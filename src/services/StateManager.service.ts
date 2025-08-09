/**
 * State Manager for IPG Taxonomy Extractor
 * 
 * Centralized state management with:
 * - Immutable state updates
 * - Change detection
 * - Subscription system
 * - Performance optimization
 * - Type safety
 */

import { AppState, ParsedCellData, UndoOperation, StateChangeEvent, StateChangeListener } from '../types/taxonomy.types';

export class StateManager {
  private static instance: StateManager;
  private state: AppState;
  private listeners: Set<StateChangeListener> = new Set();
  private isUpdating: boolean = false;

  public static getInstance(): StateManager {
    if (!StateManager.instance) {
      StateManager.instance = new StateManager();
    }
    return StateManager.instance;
  }

  private constructor() {
    this.state = this.getInitialState();
  }

  /**
   * Get the initial application state
   */
  private getInitialState(): AppState {
    return {
      selectedCellCount: 0,
      parsedData: null,
      undoStack: [],
      currentMode: 'normal',
      isProcessing: false,
      isInitialized: false,
    };
  }

  /**
   * Get current state (immutable)
   */
  public getState(): Readonly<AppState> {
    return { ...this.state };
  }

  /**
   * Update state with change detection and notifications
   */
  public setState(updates: Partial<AppState>): void {
    if (this.isUpdating) {
      console.warn('[StateManager] Recursive state update detected, ignoring');
      return;
    }

    const previousState = { ...this.state };
    const newState = { ...this.state, ...updates };

    // Check if any properties actually changed
    const changedProperties = this.getChangedProperties(previousState, newState);
    
    if (changedProperties.length === 0) {
      return; // No changes, skip update
    }

    this.state = newState;
    this.isUpdating = true;

    try {
      // Notify all listeners
      const event: StateChangeEvent = {
        previousState,
        currentState: { ...newState },
        changedProperties,
      };

      this.listeners.forEach(listener => {
        try {
          listener(event);
        } catch (error) {
          console.error('[StateManager] Listener error:', error);
        }
      });
    } finally {
      this.isUpdating = false;
    }
  }

  /**
   * Subscribe to state changes
   */
  public subscribe(listener: StateChangeListener): () => void {
    this.listeners.add(listener);
    
    // Return unsubscribe function
    return () => {
      this.listeners.delete(listener);
    };
  }

  /**
   * Reset state to initial values
   */
  public reset(): void {
    this.setState(this.getInitialState());
  }

  /**
   * Update parsed data with validation
   */
  public setParsedData(data: ParsedCellData | null): void {
    if (data && data.selectedCellCount < 0) {
      console.warn('[StateManager] Invalid selectedCellCount, correcting to 0');
      data = { ...data, selectedCellCount: 0 };
    }

    const mode = data?.hasTargetingPattern ? 'targeting' : 'normal';
    
    this.setState({
      parsedData: data,
      selectedCellCount: data?.selectedCellCount || 0,
      currentMode: mode,
    });
  }

  /**
   * Add operation to undo stack with capacity management
   */
  public addUndoOperation(operation: UndoOperation): void {
    const currentUndoStack = [...this.state.undoStack];
    currentUndoStack.push(operation);

    // Remove oldest operation if over limit
    const MAX_UNDO_OPERATIONS = 10;
    if (currentUndoStack.length > MAX_UNDO_OPERATIONS) {
      currentUndoStack.shift();
    }

    this.setState({ undoStack: currentUndoStack });
  }

  /**
   * Remove last undo operation and return it
   */
  public popUndoOperation(): UndoOperation | null {
    if (this.state.undoStack.length === 0) {
      return null;
    }

    const currentUndoStack = [...this.state.undoStack];
    const lastOperation = currentUndoStack.pop();

    this.setState({ undoStack: currentUndoStack });
    return lastOperation || null;
  }

  /**
   * Clear all undo operations
   */
  public clearUndoStack(): void {
    this.setState({ undoStack: [] });
  }

  /**
   * Set processing state
   */
  public setProcessing(isProcessing: boolean): void {
    this.setState({ isProcessing });
  }

  /**
   * Mark as initialized
   */
  public setInitialized(): void {
    this.setState({ isInitialized: true });
  }

  /**
   * Get the number of active listeners (for debugging)
   */
  public getListenerCount(): number {
    return this.listeners.size;
  }

  /**
   * Compare states and return changed properties
   */
  private getChangedProperties(prevState: AppState, newState: AppState): (keyof AppState)[] {
    const changedProperties: (keyof AppState)[] = [];

    for (const key of Object.keys(newState) as (keyof AppState)[]) {
      if (!this.deepEqual(prevState[key], newState[key])) {
        changedProperties.push(key);
      }
    }

    return changedProperties;
  }

  /**
   * Deep equality check for complex objects
   */
  private deepEqual(a: any, b: any): boolean {
    if (a === b) return true;
    
    if (a instanceof Date && b instanceof Date) {
      return a.getTime() === b.getTime();
    }
    
    if (!a || !b || (typeof a !== 'object' && typeof b !== 'object')) {
      return a === b;
    }
    
    if (a === null || a === undefined || b === null || b === undefined) {
      return false;
    }
    
    if (a.prototype !== b.prototype) return false;
    
    let keys = Object.keys(a);
    if (keys.length !== Object.keys(b).length) {
      return false;
    }
    
    return keys.every(k => this.deepEqual(a[k], b[k]));
  }

  /**
   * Get state summary for debugging
   */
  public getStateSummary(): string {
    const { selectedCellCount, currentMode, undoStack, isProcessing, isInitialized } = this.state;
    return `StateManager: ${selectedCellCount} cells, ${currentMode} mode, ${undoStack.length} undo ops, processing=${isProcessing}, init=${isInitialized}`;
  }
}