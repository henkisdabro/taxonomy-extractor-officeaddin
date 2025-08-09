/**
 * UndoSystem Component for IPG Taxonomy Extractor
 * 
 * Manages the multi-step undo functionality with:
 * - LIFO stack operations (matches VBA implementation)
 * - Dynamic button captions based on operation count
 * - Automatic capacity management (10 operations max)
 * - Excel range restoration with precise cell addressing
 * - Comprehensive error handling and status updates
 */

import { BaseComponent } from './BaseComponent';
import { BaseComponentProps, UndoOperation, UndoData, StateManager } from '../types/taxonomy.types';

export interface UndoSystemProps extends BaseComponentProps {
  stateManager: StateManager;
  onStatusUpdate?: (message: string, isError?: boolean) => void;
  onShowMessage?: (message: string) => void;
}

export class UndoSystemComponent extends BaseComponent<UndoSystemProps> {
  private btnUndo!: HTMLButtonElement;
  private lblUndoWarning!: HTMLElement;
  private nextOperationId = 1;
  
  protected initialize(): void {
    this.log('INFO', 'Initializing UndoSystem component');
    this.bindElements();
    this.subscribeToStateChanges();
    this.render();
  }

  protected render(): void {
    this.updateButtonState();
  }

  private bindElements(): void {
    // Find undo button (assumes it exists in DOM)
    this.btnUndo = this.querySelector<HTMLButtonElement>('#btnUndo')!;
    this.lblUndoWarning = this.querySelector<HTMLElement>('#lblUndoWarning')!;

    if (!this.btnUndo) {
      this.log('ERROR', 'Undo button not found in DOM');
      return;
    }

    if (!this.lblUndoWarning) {
      this.log('WARN', 'Undo warning label not found in DOM');
    }

    // Setup event listeners
    this.addEventListener(this.btnUndo, 'click', this.handleUndoClick.bind(this));
    this.setupKeyboardNavigation();

    this.log('DEBUG', 'UndoSystem elements bound successfully');
  }

  private subscribeToStateChanges(): void {
    this.subscribeToState((event) => {
      const { changedProperties } = event;
      if (changedProperties.includes('undoStack')) {
        this.updateButtonState();
      }
    });
  }

  /**
   * Handle undo button click
   */
  private async handleUndoClick(): Promise<void> {
    this.log('INFO', 'Undo button clicked');
    
    const state = this.props.stateManager.getState();
    if (state.undoStack.length === 0) {
      this.log('WARN', 'No undo operations available');
      return;
    }

    if (state.isProcessing) {
      this.log('WARN', 'Cannot undo while processing');
      return;
    }

    try {
      await this.executeUndo();
    } catch (error) {
      this.log('ERROR', 'Undo operation failed', error);
      this.notifyError(this.getString('ui.messages.error_undo'));
    }
  }

  /**
   * Execute the undo operation
   */
  private async executeUndo(): Promise<void> {
    // Set processing state
    this.props.stateManager.setProcessing(true);
    this.updateButtonProcessingState(true);

    try {
      await Excel.run(async (context) => {
        const lastOperation = this.props.stateManager.popUndoOperation();
        
        if (!lastOperation) {
          this.log('WARN', 'No operation to undo');
          return;
        }

        this.log('INFO', `Undoing operation: ${lastOperation.description}`);

        // Restore all cell values
        for (const change of lastOperation.cellChanges) {
          try {
            const range = context.workbook.worksheets.getActiveWorksheet()
              .getRange(change.cellAddress);
            range.values = [[change.originalValue]];
            
            this.log('DEBUG', `Restored cell ${change.cellAddress} to: ${change.originalValue}`);
          } catch (cellError) {
            this.log('ERROR', `Failed to restore cell ${change.cellAddress}`, cellError);
          }
        }

        await context.sync();

        this.notifySuccess(
          this.getString('ui.messages.success_undo', {
            operation: lastOperation.description,
            count: lastOperation.cellCount.toString()
          })
        );

        this.log('INFO', `Successfully undone operation ${lastOperation.operationId}`);
      });
    } catch (error) {
      this.log('ERROR', 'Excel undo operation failed', error);
      throw error;
    } finally {
      this.props.stateManager.setProcessing(false);
      this.updateButtonProcessingState(false);
    }
  }

  /**
   * Add a new operation to the undo stack
   */
  public async addOperation(description: string, range: Excel.Range): Promise<void> {
    try {
      const cellChanges: UndoData[] = [];

      await Excel.run(async (context) => {
        // Load the values to capture current state
        range.load(['values', 'address', 'rowCount', 'columnCount']);
        await context.sync();

        const values = range.values;
        const baseAddress = range.address;

        // Extract worksheet name and base range
        const addressParts = baseAddress.split('!');
        const worksheetName = addressParts.length > 1 ? addressParts[0] : '';
        const rangeAddress = addressParts.length > 1 ? addressParts[1] : baseAddress;

        // Parse the range to get starting row and column
        const rangeParts = rangeAddress.split(':');
        const startCell = rangeParts[0];
        const startMatch = startCell.match(/([A-Z]+)(\d+)/);

        if (!startMatch) {
          throw new Error(`Invalid cell address format: ${startCell}`);
        }

        const startCol = this.columnToNumber(startMatch[1]);
        const startRow = parseInt(startMatch[2]);

        // Build individual cell addresses and capture values
        for (let row = 0; row < range.rowCount; row++) {
          for (let col = 0; col < range.columnCount; col++) {
            const cellCol = this.numberToColumn(startCol + col);
            const cellRow = startRow + row;
            const cellAddress = worksheetName 
              ? `${worksheetName}!${cellCol}${cellRow}`
              : `${cellCol}${cellRow}`;

            cellChanges.push({
              cellAddress,
              originalValue: values[row][col] as string | number | boolean
            });
          }
        }
      });

      const operation: UndoOperation = {
        description,
        cellChanges,
        cellCount: cellChanges.length,
        operationId: this.nextOperationId++,
        timestamp: new Date()
      };

      this.props.stateManager.addUndoOperation(operation);

      this.log('INFO', `Added undo operation: ${description} (${cellChanges.length} cells)`);

    } catch (error) {
      this.log('ERROR', 'Failed to add undo operation', error);
      throw error;
    }
  }

  /**
   * Update button state based on current undo stack
   */
  private updateButtonState(): void {
    const state = this.props.stateManager.getState();
    const undoCount = state.undoStack.length;
    
    if (!this.btnUndo) return;

    const undoLabel = this.btnUndo.querySelector('.ms-Button-label');
    if (!undoLabel) {
      this.log('WARN', 'Undo button label not found');
      return;
    }

    if (undoCount === 0) {
      undoLabel.textContent = this.getString('ui.buttons.undo_last');
      this.btnUndo.disabled = true;
      this.btnUndo.classList.remove('undo-available');
      if (this.lblUndoWarning) {
        this.lblUndoWarning.style.display = 'none';
      }
    } else {
      undoLabel.textContent = undoCount === 1 
        ? this.getString('ui.buttons.undo_last') 
        : this.getString('ui.buttons.undo_multiple', { count: undoCount.toString() });
      
      this.btnUndo.disabled = state.isProcessing;
      this.btnUndo.classList.add('undo-available');

      // Show warning if approaching capacity
      if (this.lblUndoWarning) {
        const MAX_OPERATIONS = 10; // TODO: Get from constants
        if (undoCount >= MAX_OPERATIONS) {
          this.lblUndoWarning.style.display = 'block';
        } else {
          this.lblUndoWarning.style.display = 'none';
        }
      }
    }

    this.log('DEBUG', `Updated undo button: ${undoCount} operations available`);
  }

  /**
   * Update button state during processing
   */
  private updateButtonProcessingState(isProcessing: boolean): void {
    if (!this.btnUndo) return;

    const undoLabel = this.btnUndo.querySelector('.ms-Button-label');
    if (!undoLabel) return;

    if (isProcessing) {
      this.btnUndo.setAttribute('data-original-text', undoLabel.textContent || '');
      undoLabel.textContent = this.getString('ui.buttons.processing');
      this.btnUndo.disabled = true;
    } else {
      const originalText = this.btnUndo.getAttribute('data-original-text');
      if (originalText) {
        undoLabel.textContent = originalText;
        this.btnUndo.removeAttribute('data-original-text');
      }
      this.updateButtonState(); // Restore proper state
    }
  }

  /**
   * Handle activation (Enter/Space key press)
   */
  protected handleActivation(): void {
    if (!this.btnUndo?.disabled) {
      this.handleUndoClick();
    }
  }

  /**
   * Clear all undo operations
   */
  public clearOperations(): void {
    this.props.stateManager.clearUndoStack();
    this.log('INFO', 'Cleared all undo operations');
  }

  /**
   * Get the number of available undo operations
   */
  public getOperationCount(): number {
    return this.props.stateManager.getState().undoStack.length;
  }

  // Utility methods for Excel address conversion
  private columnToNumber(column: string): number {
    let result = 0;
    for (let i = 0; i < column.length; i++) {
      result = result * 26 + (column.charCodeAt(i) - 64);
    }
    return result;
  }

  private numberToColumn(num: number): string {
    let result = '';
    while (num > 0) {
      num--;
      result = String.fromCharCode(65 + (num % 26)) + result;
      num = Math.floor(num / 26);
    }
    return result;
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