/**
 * SegmentExtractor Component for IPG Taxonomy Extractor
 * 
 * Handles the extraction of individual segments (1-9) from pipe-delimited taxonomy data:
 * - Dynamic button generation with segment previews
 * - Pipe-delimited data parsing and validation
 * - Selective segment extraction with colon handling
 * - Integration with undo system for reversible operations
 * - Context-aware visibility (hidden in targeting mode)
 * - Performance monitoring and comprehensive error handling
 */

import { BaseComponent } from './BaseComponent';
import { BaseComponentProps, StateManager, SegmentExtractionOptions, OperationResult, Segment } from '../types/taxonomy.types';

export interface SegmentExtractorProps extends BaseComponentProps {
  stateManager: StateManager;
  undoSystem?: {
    addOperation: (description: string, range: Excel.Range) => Promise<void>;
  };
  onStatusUpdate?: (message: string, isError?: boolean) => void;
  onShowMessage?: (message: string) => void;
  maxButtonTextLength?: number;
}

export class SegmentExtractorComponent extends BaseComponent<SegmentExtractorProps> {
  private segmentButtons: HTMLButtonElement[] = [];
  private readonly SEGMENT_COUNT = 9;
  private readonly MAX_BUTTON_TEXT_LENGTH: number;
  
  constructor(element: HTMLElement, props: SegmentExtractorProps) {
    super(element, props);
    this.MAX_BUTTON_TEXT_LENGTH = props.maxButtonTextLength || 35;
  }
  
  protected initialize(): void {
    this.log('INFO', 'Initializing SegmentExtractor component');
    this.bindElements();
    this.subscribeToStateChanges();
    this.render();
  }

  protected render(): void {
    this.updateButtonStates();
  }

  private bindElements(): void {
    // Find all segment buttons (1-9)
    this.segmentButtons = [];
    
    for (let i = 1; i <= this.SEGMENT_COUNT; i++) {
      const btn = this.querySelector<HTMLButtonElement>(`#btn${i}`);
      if (btn) {
        this.segmentButtons.push(btn);
        
        // Setup event listener for each button
        this.addEventListener(btn, 'click', () => this.handleSegmentClick(i as Segment));
        
        // Setup keyboard navigation
        this.addEventListener(btn, 'keydown', (event: KeyboardEvent) => {
          this.handleSegmentKeyNavigation(event, i as Segment);
        });
      } else {
        this.log('WARN', `Segment button ${i} not found in DOM`);
      }
    }

    if (this.segmentButtons.length === 0) {
      this.log('ERROR', 'No segment buttons found in DOM');
      return;
    }

    this.log('DEBUG', `SegmentExtractor bound ${this.segmentButtons.length} segment buttons`);
  }

  private subscribeToStateChanges(): void {
    this.subscribeToState((event) => {
      const { changedProperties } = event;
      if (changedProperties.includes('parsedData') || 
          changedProperties.includes('currentMode') ||
          changedProperties.includes('isProcessing')) {
        this.updateButtonStates();
      }
    });
  }

  /**
   * Handle segment button click
   */
  private async handleSegmentClick(segmentNumber: Segment): Promise<void> {
    this.log('INFO', `Segment ${segmentNumber} button clicked`);
    
    const state = this.props.stateManager.getState();
    if (state.isProcessing) {
      this.log('WARN', 'Cannot extract while processing');
      return;
    }

    if (state.currentMode === 'targeting') {
      this.log('WARN', 'Segment extraction not available in targeting mode');
      return;
    }

    try {
      const result = await this.extractSegment(segmentNumber);
      
      if (result.success) {
        this.notifySuccess(
          this.getString('ui.messages.success_segment', { 
            number: segmentNumber.toString(),
            count: result.processedCount?.toString() || '0' 
          })
        );
      } else {
        this.notifyError(result.error || this.getString('ui.messages.error_segment'));
      }
    } catch (error) {
      this.log('ERROR', `Segment ${segmentNumber} extraction failed`, error);
      this.notifyError(this.getString('ui.messages.error_segment'));
    }
  }

  /**
   * Extract a specific segment from selected cells
   */
  public async extractSegment(segmentNumber: Segment, options?: SegmentExtractionOptions): Promise<OperationResult> {
    const validateData = options?.validateData ?? true;
    const addToUndo = options?.addToUndo ?? true;

    try {
      // Set processing state
      this.props.stateManager.setProcessing(true);
      this.updateButtonProcessingState(segmentNumber, true);

      // Performance monitoring
      const operationName = `extract-segment-${segmentNumber}`;
      const startTime = performance.now();

      const result = await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load('address, values, rowCount, columnCount');
        await context.sync();

        // Add to undo stack BEFORE making changes
        if (addToUndo && this.props.undoSystem) {
          await this.props.undoSystem.addOperation(
            this.getString('operations.extract_segment', { number: segmentNumber.toString() }),
            range
          );
        }

        const values = range.values as any[][];
        let processedCount = 0;

        // Process each cell
        for (let row = 0; row < values.length; row++) {
          for (let col = 0; col < values[row].length; col++) {
            const cellValue = values[row][col];
            
            if (cellValue && typeof cellValue === 'string') {
              const extractedSegment = this.extractSegmentFromText(cellValue, segmentNumber);
              
              if (extractedSegment !== null) {
                if (validateData && !this.validateSegmentData(extractedSegment)) {
                  this.log('WARN', `Invalid segment data: "${extractedSegment}"`);
                  continue;
                }

                values[row][col] = extractedSegment;
                processedCount++;
                
                this.log('DEBUG', `Extracted segment ${segmentNumber}: "${extractedSegment}" from "${cellValue}"`);
              }
            }
          }
        }

        // Update the range with new values
        range.values = values;
        await context.sync();

        return { success: true, processedCount };
      });

      // Performance logging
      const duration = performance.now() - startTime;
      this.log('DEBUG', `Segment ${segmentNumber} extraction took ${duration.toFixed(2)}ms`);

      // Trigger selection change to update UI
      if (typeof (window as any).taxonomyExtractor?.onSelectionChange === 'function') {
        await (window as any).taxonomyExtractor.onSelectionChange();
      }

      this.log('INFO', `Successfully extracted segment ${segmentNumber} from ${result.processedCount} cells`);
      return result;

    } catch (error) {
      this.log('ERROR', `Excel segment ${segmentNumber} extraction failed`, error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Unknown error occurred'
      };
    } finally {
      this.props.stateManager.setProcessing(false);
      this.updateButtonProcessingState(segmentNumber, false);
    }
  }

  /**
   * Extract specific segment from text string
   */
  private extractSegmentFromText(text: string, segmentNumber: number): string | null {
    if (!text || typeof text !== 'string') {
      return null;
    }

    try {
      // Split by colon first to remove activation ID
      const colonIndex = text.indexOf(':');
      const mainContent = colonIndex >= 0 ? text.substring(0, colonIndex) : text;
      
      // Split by pipes
      const segments = mainContent.split('|');
      
      if (segments.length >= segmentNumber) {
        const segment = segments[segmentNumber - 1].trim();
        return segment.length > 0 ? segment : null;
      }
      
      return null;
    } catch (error) {
      this.log('WARN', `Failed to extract segment ${segmentNumber} from: "${text}"`, error);
      return null;
    }
  }

  /**
   * Validate segment data
   */
  private validateSegmentData(segment: string): boolean {
    if (!segment || segment.trim().length === 0) {
      return false;
    }

    // Segments should not contain pipes or colons (would indicate parsing error)
    if (segment.includes('|') || segment.includes(':')) {
      return false;
    }

    // Reasonable length validation
    const trimmed = segment.trim();
    if (trimmed.length > 1000) { // Very long segments might indicate parsing errors
      return false;
    }

    return true;
  }

  /**
   * Update all segment button states based on current data
   */
  private updateButtonStates(): void {
    const state = this.props.stateManager.getState();
    
    // Hide all buttons in targeting mode
    if (state.currentMode === 'targeting') {
      this.hideAllButtons();
      return;
    }

    // Show buttons and update their content
    this.showAllButtons();

    if (!state.parsedData) {
      this.disableAllButtons();
      return;
    }

    const segments = [
      state.parsedData.segment1,
      state.parsedData.segment2,
      state.parsedData.segment3,
      state.parsedData.segment4,
      state.parsedData.segment5,
      state.parsedData.segment6,
      state.parsedData.segment7,
      state.parsedData.segment8,
      state.parsedData.segment9
    ];

    this.segmentButtons.forEach((button, index) => {
      if (!button) return;

      const segmentNumber = index + 1;
      const segmentText = segments[index] || '';
      
      this.updateButtonContent(button, segmentNumber, segmentText);
      button.disabled = state.isProcessing || segmentText.length === 0;
    });

    this.log('DEBUG', `Updated ${this.segmentButtons.length} segment button states`);
  }

  /**
   * Update individual button content and state
   */
  private updateButtonContent(button: HTMLButtonElement, segmentNumber: number, segmentText: string): void {
    const buttonLabel = button.querySelector('.ms-Button-label');
    if (!buttonLabel) return;

    if (segmentText && segmentText.length > 0) {
      // Truncate text if too long for button display
      const displayText = segmentText.length > this.MAX_BUTTON_TEXT_LENGTH
        ? `${segmentText.substring(0, this.MAX_BUTTON_TEXT_LENGTH - 3)}...`
        : segmentText;
      
      buttonLabel.textContent = displayText;
      button.classList.add('has-data');
      
      // Set title for full text on hover
      button.title = `${this.getString('ui.tooltips.extract_segment', { number: segmentNumber.toString() })}: ${segmentText}`;
    } else {
      buttonLabel.textContent = segmentNumber.toString();
      button.classList.remove('has-data');
      button.title = this.getString('ui.tooltips.no_segment_data', { number: segmentNumber.toString() });
    }
  }

  /**
   * Update button processing state
   */
  private updateButtonProcessingState(segmentNumber: Segment, isProcessing: boolean): void {
    const buttonIndex = segmentNumber - 1;
    const button = this.segmentButtons[buttonIndex];
    
    if (!button) return;

    const buttonLabel = button.querySelector('.ms-Button-label');
    if (!buttonLabel) return;

    if (isProcessing) {
      button.setAttribute('data-original-text', buttonLabel.textContent || '');
      buttonLabel.textContent = this.getString('ui.buttons.processing');
      button.disabled = true;
    } else {
      const originalText = button.getAttribute('data-original-text');
      if (originalText) {
        buttonLabel.textContent = originalText;
        button.removeAttribute('data-original-text');
      }
      this.updateButtonStates(); // Restore proper state
    }
  }

  /**
   * Handle keyboard navigation for segment buttons
   */
  private handleSegmentKeyNavigation(event: KeyboardEvent, segmentNumber: Segment): void {
    switch (event.key) {
      case 'Enter':
      case ' ':
        event.preventDefault();
        if (!event.currentTarget || (event.currentTarget as HTMLButtonElement).disabled) return;
        this.handleSegmentClick(segmentNumber);
        break;
        
      case 'ArrowLeft':
      case 'ArrowUp':
        event.preventDefault();
        this.focusPreviousButton(segmentNumber);
        break;
        
      case 'ArrowRight':
      case 'ArrowDown':
        event.preventDefault();
        this.focusNextButton(segmentNumber);
        break;
        
      case 'Home':
        event.preventDefault();
        this.focusFirstButton();
        break;
        
      case 'End':
        event.preventDefault();
        this.focusLastButton();
        break;
    }
  }

  /**
   * Focus management for keyboard navigation
   */
  private focusPreviousButton(currentSegment: Segment): void {
    const prevIndex = currentSegment - 2;
    if (prevIndex >= 0 && this.segmentButtons[prevIndex]) {
      this.segmentButtons[prevIndex].focus();
    }
  }

  private focusNextButton(currentSegment: Segment): void {
    const nextIndex = currentSegment;
    if (nextIndex < this.segmentButtons.length && this.segmentButtons[nextIndex]) {
      this.segmentButtons[nextIndex].focus();
    }
  }

  private focusFirstButton(): void {
    if (this.segmentButtons[0]) {
      this.segmentButtons[0].focus();
    }
  }

  private focusLastButton(): void {
    const lastButton = this.segmentButtons[this.segmentButtons.length - 1];
    if (lastButton) {
      lastButton.focus();
    }
  }

  /**
   * Button visibility management
   */
  private hideAllButtons(): void {
    this.segmentButtons.forEach(button => {
      if (button) button.style.display = 'none';
    });
  }

  private showAllButtons(): void {
    this.segmentButtons.forEach(button => {
      if (button) button.style.display = 'block';
    });
  }

  private disableAllButtons(): void {
    this.segmentButtons.forEach(button => {
      if (button) {
        button.disabled = true;
        const buttonLabel = button.querySelector('.ms-Button-label');
        if (buttonLabel) {
          buttonLabel.textContent = (this.segmentButtons.indexOf(button) + 1).toString();
        }
        button.classList.remove('has-data');
      }
    });
  }

  /**
   * Get segment text for a specific segment number
   */
  public getSegmentText(segmentNumber: Segment): string | null {
    const state = this.props.stateManager.getState();
    if (!state.parsedData) return null;

    const segments = [
      state.parsedData.segment1,
      state.parsedData.segment2,
      state.parsedData.segment3,
      state.parsedData.segment4,
      state.parsedData.segment5,
      state.parsedData.segment6,
      state.parsedData.segment7,
      state.parsedData.segment8,
      state.parsedData.segment9
    ];

    const segment = segments[segmentNumber - 1];
    return segment && segment.length > 0 ? segment : null;
  }

  /**
   * Get all available segments
   */
  public getAllSegments(): Record<string, string> {
    const state = this.props.stateManager.getState();
    if (!state.parsedData) return {};

    return {
      segment1: state.parsedData.segment1,
      segment2: state.parsedData.segment2,
      segment3: state.parsedData.segment3,
      segment4: state.parsedData.segment4,
      segment5: state.parsedData.segment5,
      segment6: state.parsedData.segment6,
      segment7: state.parsedData.segment7,
      segment8: state.parsedData.segment8,
      segment9: state.parsedData.segment9
    };
  }

  /**
   * Count available segments
   */
  public getAvailableSegmentCount(): number {
    const segments = Object.values(this.getAllSegments());
    return segments.filter(segment => segment && segment.length > 0).length;
  }

  // Notification helpers
  private notifySuccess(message: string): void {
    if (this.props.onShowMessage) {
      this.props.onShowMessage(message);
    } else {
      this.log('INFO', message);
    }
  }

  private notifyError(message: string): void {
    if (this.props.onStatusUpdate) {
      this.props.onStatusUpdate(message, true);
    } else {
      this.log('ERROR', message);
    }
  }
}