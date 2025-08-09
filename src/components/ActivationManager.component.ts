/**
 * ActivationManager Component for IPG Taxonomy Extractor
 * 
 * Handles activation ID extraction functionality with:
 * - Colon-separated activation ID detection
 * - Selective cell processing (only cells containing colons)
 * - Dynamic button state updates based on available data
 * - Integration with undo system for reversible operations
 * - Comprehensive error handling and user feedback
 */

import { BaseComponent } from './BaseComponent';
import { BaseComponentProps, StateManager, ActivationExtractionOptions, OperationResult } from '../types/taxonomy.types';

export interface ActivationManagerProps extends BaseComponentProps {
  stateManager: StateManager;
  undoSystem?: {
    addOperation: (description: string, range: Excel.Range) => Promise<void>;
  };
  onStatusUpdate?: (message: string, isError?: boolean) => void;
  onShowMessage?: (message: string) => void;
}

export class ActivationManagerComponent extends BaseComponent<ActivationManagerProps> {
  private btnActivationID!: HTMLButtonElement;
  
  protected initialize(): void {
    this.log('INFO', 'Initializing ActivationManager component');
    this.bindElements();
    this.subscribeToStateChanges();
    this.render();
  }

  protected render(): void {
    this.updateButtonState();
  }

  private bindElements(): void {
    // Find activation ID button
    this.btnActivationID = this.querySelector<HTMLButtonElement>('#btnActivationID')!;

    if (!this.btnActivationID) {
      this.log('ERROR', 'Activation ID button not found in DOM');
      return;
    }

    // Setup event listeners
    this.addEventListener(this.btnActivationID, 'click', this.handleExtractClick.bind(this));
    this.setupKeyboardNavigation();

    this.log('DEBUG', 'ActivationManager elements bound successfully');
  }

  private subscribeToStateChanges(): void {
    this.subscribeToState((event) => {
      const { changedProperties } = event;
      if (changedProperties.includes('parsedData') || 
          changedProperties.includes('currentMode') ||
          changedProperties.includes('isProcessing')) {
        this.updateButtonState();
      }
    });
  }

  /**
   * Handle activation ID button click
   */
  private async handleExtractClick(): Promise<void> {
    this.log('INFO', 'Activation ID extraction clicked');
    
    const state = this.props.stateManager.getState();
    if (state.isProcessing) {
      this.log('WARN', 'Cannot extract while processing');
      return;
    }

    if (!this.isExtractionAvailable()) {
      this.log('WARN', 'No activation ID available for extraction');
      return;
    }

    try {
      const result = await this.extractActivationIDs();
      
      if (result.success) {
        this.notifySuccess(
          this.getString('ui.messages.success_activation', { 
            count: result.processedCount?.toString() || '0' 
          })
        );
      } else {
        this.notifyError(result.error || this.getString('ui.messages.error_activation'));
      }
    } catch (error) {
      this.log('ERROR', 'Activation ID extraction failed', error);
      this.notifyError(this.getString('ui.messages.error_activation'));
    }
  }

  /**
   * Extract activation IDs from selected cells
   */
  public async extractActivationIDs(options?: ActivationExtractionOptions): Promise<OperationResult> {
    const validateData = options?.validateData ?? true;
    const addToUndo = options?.addToUndo ?? true;

    try {
      // Set processing state
      this.props.stateManager.setProcessing(true);
      this.updateButtonProcessingState(true);

      const result = await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load('address, values, rowCount, columnCount');
        await context.sync();

        // Add to undo stack BEFORE making changes
        if (addToUndo && this.props.undoSystem) {
          await this.props.undoSystem.addOperation(
            this.getString('operations.extract_activation'),
            range
          );
        }

        const values = range.values as any[][];
        let processedCount = 0;

        // Process each cell
        for (let row = 0; row < values.length; row++) {
          for (let col = 0; col < values[row].length; col++) {
            const cellValue = values[row][col];
            
            // Only process string cells that contain colons
            if (cellValue && typeof cellValue === 'string' && cellValue.includes(':')) {
              const extractedId = this.extractActivationIDFromText(cellValue);
              
              if (extractedId) {
                if (validateData && !this.validateActivationID(extractedId)) {
                  this.log('WARN', `Invalid activation ID format: ${extractedId}`);
                  continue;
                }

                values[row][col] = extractedId;
                processedCount++;
                
                this.log('DEBUG', `Extracted activation ID: "${extractedId}" from "${cellValue}"`);
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

      this.log('INFO', `Successfully extracted ${result.processedCount} activation IDs`);
      return result;

    } catch (error) {
      this.log('ERROR', 'Excel activation ID extraction failed', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Unknown error occurred'
      };
    } finally {
      this.props.stateManager.setProcessing(false);
      this.updateButtonProcessingState(false);
    }
  }

  /**
   * Extract activation ID from text using colon delimiter
   */
  private extractActivationIDFromText(text: string): string | null {
    if (!text || typeof text !== 'string') {
      return null;
    }

    const colonIndex = text.indexOf(':');
    if (colonIndex < 0 || colonIndex >= text.length - 1) {
      return null;
    }

    const activationId = text.substring(colonIndex + 1).trim();
    return activationId.length > 0 ? activationId : null;
  }

  /**
   * Validate activation ID format
   */
  private validateActivationID(activationId: string): boolean {
    if (!activationId || activationId.trim().length === 0) {
      return false;
    }

    // Basic validation - activation IDs should not contain pipes or additional colons
    if (activationId.includes('|') || activationId.includes(':')) {
      return false;
    }

    // Length validation (reasonable bounds)
    const trimmed = activationId.trim();
    if (trimmed.length < 1 || trimmed.length > 100) {
      return false;
    }

    return true;
  }

  /**
   * Check if activation ID extraction is available
   */
  private isExtractionAvailable(): boolean {
    const state = this.props.stateManager.getState();
    
    if (state.currentMode === 'targeting') {
      return false; // Not available in targeting mode
    }

    if (!state.parsedData || state.parsedData.selectedCellCount === 0) {
      return false; // No selection
    }

    return state.parsedData.activationId.length > 0;
  }

  /**
   * Update button state based on current data
   */
  private updateButtonState(): void {
    if (!this.btnActivationID) return;

    const state = this.props.stateManager.getState();
    const buttonLabel = this.btnActivationID.querySelector('.ms-Button-label');
    
    if (!buttonLabel) {
      this.log('WARN', 'Activation ID button label not found');
      return;
    }

    // Hide button in targeting mode
    if (state.currentMode === 'targeting') {
      this.btnActivationID.style.display = 'none';
      return;
    } else {
      this.btnActivationID.style.display = 'block';
    }

    // Update button text and state based on available data
    if (state.parsedData && state.parsedData.activationId) {
      buttonLabel.textContent = state.parsedData.activationId;
      this.btnActivationID.disabled = state.isProcessing;
      this.btnActivationID.classList.add('has-data');
      
      // Set title for better UX
      this.btnActivationID.title = this.getString('ui.tooltips.extract_activation', {
        id: state.parsedData.activationId
      });
    } else {
      buttonLabel.textContent = this.getString('ui.buttons.activation_id');
      this.btnActivationID.disabled = true;
      this.btnActivationID.classList.remove('has-data');
      this.btnActivationID.title = this.getString('ui.tooltips.no_activation_id');
    }

    this.log('DEBUG', `Updated activation button state: ${this.btnActivationID.disabled ? 'disabled' : 'enabled'}`);
  }

  /**
   * Update button state during processing
   */
  private updateButtonProcessingState(isProcessing: boolean): void {
    if (!this.btnActivationID) return;

    const buttonLabel = this.btnActivationID.querySelector('.ms-Button-label');
    if (!buttonLabel) return;

    if (isProcessing) {
      this.btnActivationID.setAttribute('data-original-text', buttonLabel.textContent || '');
      buttonLabel.textContent = this.getString('ui.buttons.processing');
      this.btnActivationID.disabled = true;
    } else {
      const originalText = this.btnActivationID.getAttribute('data-original-text');
      if (originalText) {
        buttonLabel.textContent = originalText;
        this.btnActivationID.removeAttribute('data-original-text');
      }
      this.updateButtonState(); // Restore proper state
    }
  }

  /**
   * Handle activation (Enter/Space key press)
   */
  protected handleActivation(): void {
    if (!this.btnActivationID?.disabled) {
      this.handleExtractClick();
    }
  }

  /**
   * Get the current activation ID from parsed data
   */
  public getCurrentActivationID(): string | null {
    const state = this.props.stateManager.getState();
    return state.parsedData?.activationId || null;
  }

  /**
   * Check if activation ID is available in current selection
   */
  public hasActivationID(): boolean {
    const currentId = this.getCurrentActivationID();
    return currentId !== null && currentId.length > 0;
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