/**
 * TargetingProcessor Component for IPG Taxonomy Extractor
 * 
 * Handles ^ABC^ pattern processing with dual functionality:
 * - Trim mode: Remove ^ABC^ patterns while keeping surrounding text
 * - Keep mode: Extract only ^ABC^ patterns, discard everything else
 * - Context-aware visibility (only shown when patterns detected)
 * - Integration with undo system for reversible operations
 * - Pattern validation and comprehensive error handling
 */

import { BaseComponent } from './BaseComponent';
import { BaseComponentProps, StateManager, TargetingProcessOptions, OperationResult } from '../types/taxonomy.types';

export interface TargetingProcessorProps extends BaseComponentProps {
  stateManager: StateManager;
  undoSystem?: {
    addOperation: (description: string, range: Excel.Range) => Promise<void>;
  };
  onStatusUpdate?: (message: string, isError?: boolean) => void;
  onShowMessage?: (message: string) => void;
}

export class TargetingProcessorComponent extends BaseComponent<TargetingProcessorProps> {
  private btnTargeting!: HTMLButtonElement;     // Trim button
  private btnKeepPattern!: HTMLButtonElement;   // Keep button
  private targetingSection!: HTMLElement;
  
  // Pattern matching regex (matches VBA implementation)
  private static readonly TARGETING_PATTERN_REGEX = /\^[^^]+\^/g;
  
  protected initialize(): void {
    this.log('INFO', 'Initializing TargetingProcessor component');
    this.bindElements();
    this.subscribeToStateChanges();
    this.render();
  }

  protected render(): void {
    this.updateVisibilityAndState();
  }

  private bindElements(): void {
    // Find targeting-related elements
    this.btnTargeting = this.querySelector<HTMLButtonElement>('#btnTargeting')!;
    this.btnKeepPattern = this.querySelector<HTMLButtonElement>('#btnKeepPattern')!;
    this.targetingSection = this.querySelector<HTMLElement>('#targetingSection')!;

    if (!this.btnTargeting) {
      this.log('ERROR', 'Targeting trim button not found in DOM');
      return;
    }

    if (!this.btnKeepPattern) {
      this.log('ERROR', 'Targeting keep button not found in DOM');
      return;
    }

    if (!this.targetingSection) {
      this.log('ERROR', 'Targeting section not found in DOM');
      return;
    }

    // Setup event listeners
    this.addEventListener(this.btnTargeting, 'click', () => this.handleTrimClick());
    this.addEventListener(this.btnKeepPattern, 'click', () => this.handleKeepClick());
    
    // Setup keyboard navigation for both buttons
    this.addEventListener(this.btnTargeting, 'keydown', this.handleKeyNavigation.bind(this));
    this.addEventListener(this.btnKeepPattern, 'keydown', this.handleKeyNavigation.bind(this));

    this.log('DEBUG', 'TargetingProcessor elements bound successfully');
  }

  private subscribeToStateChanges(): void {
    this.subscribeToState((event) => {
      const { changedProperties } = event;
      if (changedProperties.includes('parsedData') || 
          changedProperties.includes('currentMode') ||
          changedProperties.includes('isProcessing')) {
        this.updateVisibilityAndState();
      }
    });
  }

  /**
   * Handle trim button click (remove ^ABC^ patterns)
   */
  private async handleTrimClick(): Promise<void> {
    this.log('INFO', 'Targeting trim button clicked');
    
    const state = this.props.stateManager.getState();
    if (state.isProcessing) {
      this.log('WARN', 'Cannot process while already processing');
      return;
    }

    try {
      const result = await this.processTargetingPatterns('trim');
      
      if (result.success) {
        this.notifySuccess(
          this.getString('ui.messages.success_targeting_trim', { 
            count: result.processedCount?.toString() || '0' 
          })
        );
      } else {
        this.notifyError(result.error || this.getString('ui.messages.error_targeting'));
      }
    } catch (error) {
      this.log('ERROR', 'Targeting trim failed', error);
      this.notifyError(this.getString('ui.messages.error_targeting'));
    }
  }

  /**
   * Handle keep button click (keep only ^ABC^ patterns)
   */
  private async handleKeepClick(): Promise<void> {
    this.log('INFO', 'Targeting keep button clicked');
    
    const state = this.props.stateManager.getState();
    if (state.isProcessing) {
      this.log('WARN', 'Cannot process while already processing');
      return;
    }

    try {
      const result = await this.processTargetingPatterns('keep');
      
      if (result.success) {
        this.notifySuccess(
          this.getString('ui.messages.success_targeting_keep', { 
            count: result.processedCount?.toString() || '0' 
          })
        );
      } else {
        this.notifyError(result.error || this.getString('ui.messages.error_targeting'));
      }
    } catch (error) {
      this.log('ERROR', 'Targeting keep failed', error);
      this.notifyError(this.getString('ui.messages.error_targeting'));
    }
  }

  /**
   * Process targeting patterns based on mode
   */
  public async processTargetingPatterns(mode: 'trim' | 'keep', options?: TargetingProcessOptions): Promise<OperationResult> {
    const validatePattern = options?.validatePattern ?? true;
    const addToUndo = options?.addToUndo ?? true;

    try {
      // Set processing state
      this.props.stateManager.setProcessing(true);
      this.updateButtonsProcessingState(true);

      const result = await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load('address, values, rowCount, columnCount');
        await context.sync();

        // Add to undo stack BEFORE making changes
        if (addToUndo && this.props.undoSystem) {
          const operationKey = mode === 'trim' ? 'operations.clean_targeting' : 'operations.keep_targeting';
          await this.props.undoSystem.addOperation(
            this.getString(operationKey),
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
              const processed = mode === 'trim' 
                ? this.trimTargetingPatterns(cellValue, validatePattern)
                : this.keepTargetingPatterns(cellValue, validatePattern);

              if (processed !== null && processed !== cellValue) {
                values[row][col] = processed;
                processedCount++;
                
                this.log('DEBUG', `${mode === 'trim' ? 'Trimmed' : 'Kept'} patterns in cell: "${cellValue}" → "${processed}"`);
              }
            }
          }
        }

        // Update the range with new values
        range.values = values;
        await context.sync();

        return { success: true, processedCount };
      });

      // Trigger selection change to update UI
      if (typeof (window as any).taxonomyExtractor?.onSelectionChange === 'function') {
        await (window as any).taxonomyExtractor.onSelectionChange();
      }

      this.log('INFO', `Successfully processed ${result.processedCount} cells in ${mode} mode`);
      return result;

    } catch (error) {
      this.log('ERROR', `Excel targeting ${mode} operation failed`, error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Unknown error occurred'
      };
    } finally {
      this.props.stateManager.setProcessing(false);
      this.updateButtonsProcessingState(false);
    }
  }

  /**
   * Remove ^ABC^ patterns from text, keep everything else
   */
  private trimTargetingPatterns(text: string, validate: boolean = true): string | null {
    if (!text || typeof text !== 'string') {
      return null;
    }

    // Check if text contains targeting patterns
    if (!TargetingProcessorComponent.TARGETING_PATTERN_REGEX.test(text)) {
      return null; // No patterns to trim
    }

    // Remove all ^ABC^ patterns with optional trailing space
    const trimmed = text.replace(/\^[^^]+\^ ?/g, '').trim();
    
    if (validate && trimmed.length === 0) {
      this.log('WARN', `Trimming would result in empty string: "${text}"`);
      return null; // Don't create empty cells
    }

    return trimmed;
  }

  /**
   * Keep only ^ABC^ patterns, remove everything else
   */
  private keepTargetingPatterns(text: string, validate: boolean = true): string | null {
    if (!text || typeof text !== 'string') {
      return null;
    }

    // Extract all targeting patterns
    const matches = text.match(TargetingProcessorComponent.TARGETING_PATTERN_REGEX);
    
    if (!matches || matches.length === 0) {
      if (validate) {
        this.log('WARN', `No targeting patterns found to keep in: "${text}"`);
      }
      return null; // No patterns to keep
    }

    // Join all patterns found with spaces
    const keptPattern = matches.join(' ').trim();
    
    if (keptPattern === text) {
      return null; // No change needed
    }

    return keptPattern;
  }

  /**
   * Validate targeting pattern format
   */
  public static validateTargetingPattern(pattern: string): boolean {
    if (!pattern || typeof pattern !== 'string') {
      return false;
    }

    return TargetingProcessorComponent.TARGETING_PATTERN_REGEX.test(pattern);
  }

  /**
   * Check if text contains targeting patterns
   */
  public static hasTargetingPatterns(text: string): boolean {
    if (!text || typeof text !== 'string') {
      return false;
    }

    return TargetingProcessorComponent.TARGETING_PATTERN_REGEX.test(text);
  }

  /**
   * Update visibility and button states
   */
  private updateVisibilityAndState(): void {
    const state = this.props.stateManager.getState();
    
    // Show/hide targeting section based on mode
    if (state.currentMode === 'targeting') {
      this.showTargetingControls(state);
    } else {
      this.hideTargetingControls();
    }
  }

  /**
   * Show targeting controls when in targeting mode
   */
  private showTargetingControls(state: any): void {
    if (this.targetingSection) {
      this.targetingSection.style.display = 'block';
    }

    // Update button labels and states
    this.updateButtonLabel(this.btnTargeting, 'ui.buttons.targeting_trim');
    this.updateButtonLabel(this.btnKeepPattern, 'ui.buttons.targeting_keep');

    // Enable/disable based on processing state
    this.btnTargeting.disabled = state.isProcessing;
    this.btnKeepPattern.disabled = state.isProcessing;

    // Add targeting mode class to body for CSS styling
    document.body.classList.add('targeting-mode');

    this.log('DEBUG', 'Targeting controls shown');
  }

  /**
   * Hide targeting controls when not in targeting mode
   */
  private hideTargetingControls(): void {
    if (this.targetingSection) {
      this.targetingSection.style.display = 'none';
    }

    // Disable buttons
    this.btnTargeting.disabled = true;
    this.btnKeepPattern.disabled = true;

    // Remove targeting mode class
    document.body.classList.remove('targeting-mode');

    this.log('DEBUG', 'Targeting controls hidden');
  }

  /**
   * Update button label text
   */
  private updateButtonLabel(button: HTMLButtonElement, labelKey: string): void {
    if (!button) return;

    const buttonLabel = button.querySelector('.ms-Button-label');
    if (buttonLabel) {
      buttonLabel.textContent = this.getString(labelKey);
    }
  }

  /**
   * Update button states during processing
   */
  private updateButtonsProcessingState(isProcessing: boolean): void {
    [this.btnTargeting, this.btnKeepPattern].forEach(button => {
      if (!button) return;

      const buttonLabel = button.querySelector('.ms-Button-label');
      if (!buttonLabel) return;

      if (isProcessing) {
        const originalText = buttonLabel.textContent || '';
        button.setAttribute('data-original-text', originalText);
        buttonLabel.textContent = this.getString('ui.buttons.processing');
        button.disabled = true;
      } else {
        const originalText = button.getAttribute('data-original-text');
        if (originalText) {
          buttonLabel.textContent = originalText;
          button.removeAttribute('data-original-text');
        }
        this.updateVisibilityAndState(); // Restore proper state
      }
    });
  }

  /**
   * Handle keyboard navigation for targeting buttons
   */
  private handleKeyNavigation(event: KeyboardEvent): void {
    switch (event.key) {
      case 'Enter':
      case ' ':
        event.preventDefault();
        const target = event.target as HTMLButtonElement;
        if (target === this.btnTargeting) {
          this.handleTrimClick();
        } else if (target === this.btnKeepPattern) {
          this.handleKeepClick();
        }
        break;
      case 'Escape':
        // Focus should return to cell selection
        event.preventDefault();
        break;
    }
  }

  /**
   * Get targeting text from current selection
   */
  public getCurrentTargetingText(): string | null {
    const state = this.props.stateManager.getState();
    return state.parsedData?.targetingText || null;
  }

  /**
   * Check if targeting mode is active
   */
  public isTargetingModeActive(): boolean {
    const state = this.props.stateManager.getState();
    return state.currentMode === 'targeting';
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