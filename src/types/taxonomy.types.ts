/**
 * Type definitions for IPG Taxonomy Extractor
 * 
 * Comprehensive type safety for Excel operations, UI state, and business logic
 */

// Excel API type definitions
export type CellValue = string | number | boolean | null;

export interface ExcelRange {
  address: string;
  values: CellValue[][];
  rowCount: number;
  columnCount: number;
}

export interface SafeRangeResult<T> {
  success: boolean;
  data?: T;
  error?: string;
}

// Undo system types (from original VBA implementation)
export interface UndoData {
  cellAddress: string;
  originalValue: string | number | boolean;
}

export interface UndoOperation {
  description: string;
  cellChanges: UndoData[];
  cellCount: number;
  operationId: number;
  timestamp: Date;
}

// Parsed cell data structure
export interface ParsedCellData {
  originalText: string;
  truncatedDisplay: string;
  selectedCellCount: number;
  segment1: string;
  segment2: string;
  segment3: string;
  segment4: string;
  segment5: string;
  segment6: string;
  segment7: string;
  segment8: string;
  segment9: string;
  activationId: string;
  hasTargetingPattern: boolean;
  targetingText: string;
}

// Application state management
export interface AppState {
  selectedCellCount: number;
  parsedData: ParsedCellData | null;
  undoStack: UndoOperation[];
  currentMode: 'normal' | 'targeting';
  isProcessing: boolean;
  isInitialized: boolean;
}

// Component types
export interface ComponentProps {
  [key: string]: any;
}

export interface BaseComponentProps extends ComponentProps {
  stateManager?: StateManager;
  localization?: LocalizationService;
}

// Events
export interface StateChangeEvent {
  previousState: AppState;
  currentState: AppState;
  changedProperties: (keyof AppState)[];
}

export type StateChangeListener = (event: StateChangeEvent) => void;

// Service types
export interface LocalizationService {
  getString(key: string, params?: { [key: string]: string | number }): string;
  initialize(): Promise<void>;
  getCurrentLocale(): string;
  getIsInitialized(): boolean;
}

export interface StateManager {
  getState(): Readonly<AppState>;
  setState(updates: Partial<AppState>): void;
  subscribe(listener: StateChangeListener): () => void;
  reset(): void;
}

// Operation results
export interface OperationResult {
  success: boolean;
  processedCount?: number;
  error?: string;
  data?: any;
}

// Segment extraction specific types
export interface SegmentExtractionOptions {
  segmentNumber: number;
  validateData?: boolean;
  addToUndo?: boolean;
}

export interface ActivationExtractionOptions {
  validateData?: boolean;
  addToUndo?: boolean;
}

export interface TargetingProcessOptions {
  mode: 'trim' | 'keep';
  validatePattern?: boolean;
  addToUndo?: boolean;
}

// Error types
export interface TaxonomyError extends Error {
  code: string;
  context: string;
  severity: 'low' | 'medium' | 'high';
}

// Constants
export const CONSTANTS = {
  MAX_UNDO_OPERATIONS: 10,
  SEGMENT_COUNT: 9,
  MAX_BUTTON_TEXT_LENGTH: 35,
  TARGETING_PATTERN_REGEX: /\^[^^]+\^/g,
  PIPE_DELIMITER: '|',
  COLON_DELIMITER: ':',
} as const;

// Utility types
export type Segment = 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9;
export type LogLevel = 'ERROR' | 'WARN' | 'INFO' | 'DEBUG';
export type OperationType = 'extract_segment' | 'extract_activation' | 'clean_targeting' | 'keep_targeting';

// Re-export for convenience
export { LocalizationService as ILocalizationService } from '../services/Localization.service';