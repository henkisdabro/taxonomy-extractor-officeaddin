/*
 * IPG Taxonomy Extractor v2.0.0
 * Modern Office Add-in port of the VBA-based taxonomy extraction tool
 * 
 * Features:
 * - Real-time selection change handling
 * - Multi-step undo system (10 operations LIFO)
 * - Dynamic button previews
 * - Targeting acronym detection and removal
 * - Smart data validation
 * - Comprehensive error handling and logging
 */

// Import styles
import './taskpane.css';

// Enhanced logging system
class Logger {
  private static readonly LOG_LEVELS = {
    ERROR: 0,
    WARN: 1,
    INFO: 2,
    DEBUG: 3
  };

  private static currentLevel = Logger.LOG_LEVELS.INFO;

  static error(message: string, data?: any): void {
    if (Logger.currentLevel >= Logger.LOG_LEVELS.ERROR) {
      console.error(`[IPG Taxonomy] ERROR: ${message}`, data || '');
    }
  }

  static warn(message: string, data?: any): void {
    if (Logger.currentLevel >= Logger.LOG_LEVELS.WARN) {
      console.warn(`[IPG Taxonomy] WARN: ${message}`, data || '');
    }
  }

  static info(message: string, data?: any): void {
    if (Logger.currentLevel >= Logger.LOG_LEVELS.INFO) {
      console.log(`[IPG Taxonomy] INFO: ${message}`, data || '');
    }
  }

  static debug(message: string, data?: any): void {
    if (Logger.currentLevel >= Logger.LOG_LEVELS.DEBUG) {
      console.log(`[IPG Taxonomy] DEBUG: ${message}`, data || '');
    }
  }

  static setLevel(level: keyof typeof Logger.LOG_LEVELS): void {
    Logger.currentLevel = Logger.LOG_LEVELS[level];
  }
}

// Performance monitoring
class PerformanceMonitor {
  private static operations: Map<string, number> = new Map();

  static startOperation(name: string): void {
    this.operations.set(name, performance.now());
  }

  static endOperation(name: string): number {
    const start = this.operations.get(name);
    if (start) {
      const duration = performance.now() - start;
      Logger.debug(`Operation '${name}' took ${duration.toFixed(2)}ms`);
      this.operations.delete(name);
      return duration;
    }
    return 0;
  }
}

// Data structures matching VBA implementation
interface UndoData {
  cellAddress: string;
  originalValue: string | number | boolean;
}

interface UndoOperation {
  description: string;
  cellChanges: UndoData[];
  cellCount: number;
  operationId: number;
  timestamp: Date;
}

interface ParsedCellData {
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

// Dark mode detection and application
class DarkModeManager {
  private static isDarkMode(): boolean {
    // Check multiple indicators for dark mode
    const prefersColorScheme = window.matchMedia('(prefers-color-scheme: dark)').matches;
    const bodyComputedStyle = window.getComputedStyle(document.body);
    const backgroundColor = bodyComputedStyle.backgroundColor;
    
    // Excel Desktop in dark mode often sets a dark background
    const isDarkBackground = this.isColorDark(backgroundColor);
    
    // Additional check: look for dark theme indicators in the document
    const documentElement = document.documentElement;
    const htmlComputedStyle = window.getComputedStyle(documentElement);
    const htmlBackground = htmlComputedStyle.backgroundColor;
    const isHtmlDarkBackground = this.isColorDark(htmlBackground);
    
    // Check if Office.js provides theme information
    let officeTheme = false;
    try {
      if (typeof Office !== 'undefined' && Office.context && Office.context.officeTheme) {
        const theme = Office.context.officeTheme;
        // Check if any theme colors suggest dark mode
        officeTheme = theme.bodyBackgroundColor === '#1e1e1e' || theme.bodyBackgroundColor === '#2d2d2d';
      }
    } catch (e) {
      // Office theme not available
    }
    
    Logger.info(`Dark mode detection: prefers-color-scheme=${prefersColorScheme}, background=${backgroundColor}, isDarkBackground=${isDarkBackground}, htmlBackground=${htmlBackground}, isHtmlDarkBackground=${isHtmlDarkBackground}, officeTheme=${officeTheme}`);
    
    return prefersColorScheme || isDarkBackground || isHtmlDarkBackground || officeTheme;
  }
  
  private static isColorDark(color: string): boolean {
    // Convert color to RGB and check if it's dark
    const rgb = color.match(/\d+/g);
    if (rgb) {
      const r = parseInt(rgb[0]);
      const g = parseInt(rgb[1]);
      const b = parseInt(rgb[2]);
      // Calculate luminance
      const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
      return luminance < 0.5;
    }
    return false;
  }
  
  static applyDarkModeStyles(force: boolean = false): void {
    if (this.isDarkMode() || force || document.body.classList.contains('dark-mode')) {
      Logger.info('Applying dark mode styles programmatically');
      document.body.classList.add('dark-mode');
      
      // Remove any existing injected dark styles first
      const existingStyles = document.querySelectorAll('style[data-dark-mode="true"]');
      existingStyles.forEach(style => style.remove());
      
      // Apply styles programmatically for better compatibility
      const darkStyles = `
        .dark-mode {
          background-color: #1e1e1e !important;
          color: #ffffff !important;
        }
        
        .dark-mode .ms-Grid {
          background-color: #1e1e1e !important;
        }
        
        .dark-mode .app-header {
          background: linear-gradient(135deg, #0f6cbd 0%, #004578 100%) !important;
          color: #ffffff !important;
        }
        
        .dark-mode .ms-Callout {
          background-color: #2d2d2d !important;
          border-color: #404040 !important;
          color: #ffffff !important;
        }
        
        .dark-mode .ms-Callout--info {
          border-left-color: #0078d4 !important;
          background-color: #1a2f3d !important;
        }
        
        .dark-mode .segment-btn,
        .dark-mode .ms-Button--default {
          background-color: #2d2d2d !important;
          border-color: #555555 !important;
          color: #ffffff !important;
        }
        
        .dark-mode .segment-btn:hover:not(:disabled),
        .dark-mode .ms-Button--default:hover:not(:disabled) {
          background-color: #383838 !important;
          border-color: #888888 !important;
        }
        
        .dark-mode .segment-btn.ms-Button--primary,
        .dark-mode .ms-Button--primary {
          background-color: #4a4a4a !important;
          border-color: #666666 !important;
          color: #ffffff !important;
        }
        
        .dark-mode .segment-btn.ms-Button--primary:hover:not(:disabled),
        .dark-mode .ms-Button--primary:hover:not(:disabled) {
          background-color: #555555 !important;
          border-color: #777777 !important;
        }
        
        .dark-mode .segment-btn:disabled,
        .dark-mode .ms-Button--default:disabled {
          background-color: #252525 !important;
          border-color: #3a3a3a !important;
          color: #666666 !important;
        }
        
        .dark-mode .ms-Button--compound {
          background-color: #2d2d2d !important;
          border-color: #404040 !important;
        }
        
        .dark-mode .ms-Button--compound .ms-Button-label {
          color: #ffffff !important;
        }
        
        .dark-mode .ms-Button--compound .ms-Button-description {
          color: #cccccc !important;
        }
        
        .dark-mode .ms-Button--compound:hover:not(:disabled) {
          background-color: #383838 !important;
          border-color: #888888 !important;
        }
        
        .dark-mode .ms-Button--compound:disabled {
          background-color: #252525 !important;
          border-color: #3a3a3a !important;
        }
        
        .dark-mode .ms-Button--compound:disabled .ms-Button-label {
          color: #666666 !important;
        }
        
        .dark-mode .ms-Button--compound:disabled .ms-Button-description {
          color: #555555 !important;
        }
        
        .dark-mode .undo-section {
          background-color: #2d2d2d !important;
          border-color: #404040 !important;
        }
        
        .dark-mode .undo-section .ms-font-m {
          color: #ffffff !important;
        }
        
        .dark-mode #btnUndo:not(:disabled) {
          background-color: #3d2f1f !important;
          border-color: #f7930e !important;
          color: #ffb84d !important;
        }
        
        .dark-mode #btnUndo:hover:not(:disabled) {
          background-color: #4a3626 !important;
          border-color: #e8890a !important;
        }
        
        .dark-mode #btnUndo:disabled {
          background-color: #252525 !important;
          border-color: #3a3a3a !important;
          color: #666666 !important;
        }
        
        .dark-mode .ms-MessageBar--success {
          background-color: #1a3d17 !important;
          border-color: #92c353 !important;
          color: #c3f0c0 !important;
        }
        
        .dark-mode .ms-MessageBar--error {
          background-color: #3d1f1a !important;
          border-color: #d13438 !important;
          color: #ffcccc !important;
        }
        
        .dark-mode .ms-MessageBar--info {
          background-color: #1a2f3d !important;
          border-color: #0078d4 !important;
          color: #b3d9ff !important;
        }
        
        .dark-mode #lblUndoWarning {
          background-color: #3d1f1a !important;
          border-color: #d13438 !important;
          color: #ffcccc !important;
        }
        
        .dark-mode .ms-font-m,
        .dark-mode .ms-font-xl,
        .dark-mode .ms-font-s {
          color: #ffffff !important;
        }
        
        .dark-mode .ms-fontColor-neutralSecondary {
          color: #cccccc !important;
        }
      `;
      
      // Inject styles
      const styleElement = document.createElement('style');
      styleElement.setAttribute('data-dark-mode', 'true');
      styleElement.textContent = darkStyles;
      document.head.appendChild(styleElement);
      
      Logger.info('Dark mode styles applied successfully');
    } else {
      Logger.info('Light mode detected, using default styles');
    }
  }
  
  static setupDarkModeMonitoring(): void {
    // Monitor for changes in color scheme preference
    const mediaQuery = window.matchMedia('(prefers-color-scheme: dark)');
    mediaQuery.addEventListener('change', () => {
      Logger.info('Color scheme preference changed, reapplying styles');
      // Remove existing dark mode class and reapply
      document.body.classList.remove('dark-mode');
      this.applyDarkModeStyles();
    });
  }
  
  // Debug function to manually force dark mode
  static forceDarkMode(): void {
    Logger.info('Manually forcing dark mode');
    document.body.classList.add('dark-mode');
    this.applyDarkModeStyles(true);
  }
  
  // Debug function to toggle dark mode
  static toggleDarkMode(): void {
    if (document.body.classList.contains('dark-mode')) {
      Logger.info('Manually removing dark mode');
      document.body.classList.remove('dark-mode');
      const existingStyles = document.querySelectorAll('style[data-dark-mode="true"]');
      existingStyles.forEach(style => style.remove());
    } else {
      this.forceDarkMode();
    }
  }
}

// Global state management
class TaxonomyExtractor {
  private static readonly MAX_UNDO_OPERATIONS = 10;
  private undoStack: UndoOperation[] = [];
  private nextOperationId = 1;

  // UI elements
  private lblInstructions!: HTMLElement;
  private lblCellCount!: HTMLElement;
  private segmentButtons!: HTMLButtonElement[];
  private btnActivationID!: HTMLButtonElement;
  private btnTargeting!: HTMLButtonElement;
  private btnUndo!: HTMLButtonElement;
  private lblUndoWarning!: HTMLElement;
  private statusBar!: HTMLElement;
  private statusText!: HTMLElement;

  constructor() {
    this.initializeUI();
    this.setupEventHandlers();
    this.initializeDarkMode();
  }
  
  private initializeDarkMode(): void {
    Logger.info('Initializing dark mode support...');
    
    // Apply dark mode styles immediately and again after delays to ensure they take effect
    DarkModeManager.applyDarkModeStyles();
    
    setTimeout(() => {
      DarkModeManager.applyDarkModeStyles();
    }, 100);
    
    setTimeout(() => {
      DarkModeManager.applyDarkModeStyles();
      DarkModeManager.setupDarkModeMonitoring();
    }, 500);
    
    // Force dark mode if we detect Windows 11 dark task pane
    setTimeout(() => {
      this.forceDetectDarkMode();
    }, 1000);
  }
  
  private forceDetectDarkMode(): void {
    const taskPaneContainer = document.querySelector('.ms-Grid') || document.body;
    const computedStyle = window.getComputedStyle(taskPaneContainer);
    const backgroundColor = computedStyle.backgroundColor;
    
    Logger.info(`Force dark mode detection - container background: ${backgroundColor}`);
    
    // If the background is dark but we haven't applied dark mode, force it
    if (!document.body.classList.contains('dark-mode')) {
      const rgb = backgroundColor.match(/\d+/g);
      if (rgb) {
        const r = parseInt(rgb[0]);
        const g = parseInt(rgb[1]);
        const b = parseInt(rgb[2]);
        const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
        
        if (luminance < 0.5) {
          Logger.info('Forcing dark mode due to dark container background');
          document.body.classList.add('dark-mode');
          DarkModeManager.applyDarkModeStyles(true);
        }
      }
    }
  }

  private initializeUI(): void {
    Logger.info('Initializing UI elements...');
    
    // Get UI elements
    this.lblInstructions = document.getElementById('lblInstructions')!;
    this.lblCellCount = document.getElementById('lblCellCount')!;
    this.btnActivationID = document.getElementById('btnActivationID') as HTMLButtonElement;
    this.btnTargeting = document.getElementById('btnTargeting') as HTMLButtonElement;
    this.btnUndo = document.getElementById('btnUndo') as HTMLButtonElement;
    this.lblUndoWarning = document.getElementById('lblUndoWarning')!;
    this.statusBar = document.getElementById('statusBar')!;
    this.statusText = document.getElementById('statusText')!;

    // Verify critical elements
    if (!this.lblInstructions) Logger.error('lblInstructions not found');
    if (!this.lblCellCount) Logger.error('lblCellCount not found'); 
    if (!this.btnActivationID) Logger.error('btnActivationID not found');
    if (!this.btnTargeting) Logger.error('btnTargeting not found');
    if (!this.btnUndo) Logger.error('btnUndo not found');
    if (!this.statusBar) Logger.error('statusBar not found');
    if (!this.statusText) Logger.error('statusText not found');

    // Get all segment buttons
    this.segmentButtons = [];
    for (let i = 1; i <= 9; i++) {
      const btn = document.getElementById(`btn${i}`) as HTMLButtonElement;
      if (btn) {
        this.segmentButtons.push(btn);
        Logger.debug(`Segment button ${i} found`);
      } else {
        Logger.error(`Segment button ${i} not found`);
        this.segmentButtons.push(null as any); // Keep array indices consistent
      }
    }
    
    Logger.info(`UI initialization complete. Found ${this.segmentButtons.filter(b => b).length} segment buttons`);
  }

  private setupEventHandlers(): void {
    Logger.info('Setting up event handlers...');
    
    // Segment button event handlers
    this.segmentButtons.forEach((button, index) => {
      if (button) {
        button.addEventListener('click', () => {
          Logger.info(`Segment button ${index + 1} clicked`);
          this.extractPipeSegment(index + 1);
        });
        Logger.debug(`Segment button ${index + 1} event handler registered`);
      } else {
        Logger.error(`Segment button ${index + 1} not found`);
      }
    });

    // Activation ID button
    if (this.btnActivationID) {
      this.btnActivationID.addEventListener('click', () => {
        Logger.info('Activation ID button clicked');
        this.extractActivationID();
      });
      Logger.debug('Activation ID button event handler registered');
    } else {
      Logger.error('Activation ID button not found');
    }

    // Targeting button
    if (this.btnTargeting) {
      this.btnTargeting.addEventListener('click', () => {
        Logger.info('Targeting button clicked');
        this.cleanTargetingAcronyms();
      });
      Logger.debug('Targeting button event handler registered');
    } else {
      Logger.error('Targeting button not found');
    }

    // Undo button
    if (this.btnUndo) {
      this.btnUndo.addEventListener('click', () => {
        Logger.info('Undo button clicked');
        this.undoLastOperation();
      });
      Logger.debug('Undo button event handler registered');
    } else {
      Logger.error('Undo button not found');
    }
    
    Logger.info('Event handlers setup complete');
  }

  // Parse cell data (replicates VBA ParseFirstCellData function)
  private parseCellData(range: Excel.Range): ParsedCellData {
    const values = range.values as any[][];
    let firstTextCell = '';
    let cellCount = 0;

    // Find first non-empty text cell (matching VBA logic)
    for (let row = 0; row < values.length; row++) {
      for (let col = 0; col < values[row].length; col++) {
        const cellValue = values[row][col];
        if (cellValue && typeof cellValue === 'string' && cellValue.trim()) {
          if (!firstTextCell) {
            firstTextCell = cellValue.trim();
          }
          cellCount++;
        }
      }
    }

    const result: ParsedCellData = {
      originalText: firstTextCell,
      truncatedDisplay: firstTextCell, // Show full text without truncation
      selectedCellCount: cellCount,
      segment1: '',
      segment2: '',
      segment3: '',
      segment4: '',
      segment5: '',
      segment6: '',
      segment7: '',
      segment8: '',
      segment9: '',
      activationId: '',
      hasTargetingPattern: false,
      targetingText: ''
    };

    if (!firstTextCell) {
      return result;
    }

    // Check for targeting pattern first (^ABC^ format)
    const targetingMatch = firstTextCell.match(/\^[^^]+\^ ?/);
    if (targetingMatch && !firstTextCell.includes('|')) {
      result.hasTargetingPattern = true;
      result.targetingText = targetingMatch[0];
      return result;
    }

    // Parse pipe-delimited data (matching VBA logic)
    if (firstTextCell.includes('|')) {
      // Split by colon first to separate activation ID
      const colonParts = firstTextCell.split(':');
      const mainContent = colonParts[0];
      
      if (colonParts.length > 1) {
        result.activationId = colonParts[1].trim();
      }

      // Split main content by pipes
      const segments = mainContent.split('|');
      
      // Assign segments (VBA has 1-based indexing, we use 0-based)
      if (segments.length > 0) result.segment1 = segments[0]?.trim() || '';
      if (segments.length > 1) result.segment2 = segments[1]?.trim() || '';
      if (segments.length > 2) result.segment3 = segments[2]?.trim() || '';
      if (segments.length > 3) result.segment4 = segments[3]?.trim() || '';
      if (segments.length > 4) result.segment5 = segments[4]?.trim() || '';
      if (segments.length > 5) result.segment6 = segments[5]?.trim() || '';
      if (segments.length > 6) result.segment7 = segments[6]?.trim() || '';
      if (segments.length > 7) result.segment8 = segments[7]?.trim() || '';
      if (segments.length > 8) result.segment9 = segments[8]?.trim() || '';
    }

    return result;
  }

  // Update UI based on parsed data (replicates VBA UserForm behavior)
  private updateInterface(parsedData: ParsedCellData): void {
    // Update instructions and cell count
    if (parsedData.selectedCellCount === 0) {
      this.lblInstructions.textContent = 'Select cells containing pipe-delimited taxonomy data to begin extraction.';
      this.lblCellCount.textContent = 'No cells selected';
    } else if (parsedData.hasTargetingPattern) {
      this.lblInstructions.textContent = `Targeting pattern detected: ${parsedData.targetingText}`;
      this.lblCellCount.textContent = `${parsedData.selectedCellCount} cell(s) selected`;
    } else if (parsedData.originalText.includes('|')) {
      this.lblInstructions.textContent = parsedData.truncatedDisplay;
      this.lblCellCount.textContent = `${parsedData.selectedCellCount} cell(s) selected - Taxonomy data detected`;
    } else {
      this.lblInstructions.textContent = 'Selected cells do not contain pipe-delimited taxonomy data.';
      this.lblCellCount.textContent = `${parsedData.selectedCellCount} cell(s) selected`;
    }

    // Update segment buttons (matching VBA dynamic captions)
    const segments = [
      parsedData.segment1, parsedData.segment2, parsedData.segment3,
      parsedData.segment4, parsedData.segment5, parsedData.segment6,
      parsedData.segment7, parsedData.segment8, parsedData.segment9
    ];

    this.segmentButtons.forEach((button, index) => {
      const segment = segments[index];
      const buttonLabel = button.querySelector('.ms-Button-label')!;
      
      if (segment && parsedData.originalText.includes('|')) {
        buttonLabel.textContent = `${index + 1}: ${segment.length > 20 ? segment.substring(0, 20) + '...' : segment}`;
        button.disabled = false;
        button.classList.remove('ms-Button--default');
        button.classList.add('ms-Button--primary');
      } else {
        buttonLabel.textContent = `${index + 1}: N/A`;
        button.disabled = true;
        button.classList.remove('ms-Button--primary');
        button.classList.add('ms-Button--default');
      }
    });

    // Update activation ID button
    const activationLabel = this.btnActivationID.querySelector('.ms-Button-label')!;
    if (parsedData.activationId) {
      activationLabel.textContent = `ID: ${parsedData.activationId}`;
      this.btnActivationID.disabled = false;
    } else {
      activationLabel.textContent = 'ID: N/A';
      this.btnActivationID.disabled = true;
    }

    // Handle targeting button visibility (VBA overlay behavior)
    const targetingLabel = this.btnTargeting.querySelector('.ms-Button-label')!;
    const targetingSection = document.getElementById('targetingSection')!;
    if (parsedData.hasTargetingPattern) {
      targetingLabel.textContent = `Trim: ${parsedData.targetingText}`;
      targetingSection.style.display = 'block';
      this.btnTargeting.disabled = false;
      
      // Hide segment buttons when targeting pattern detected
      this.segmentButtons.forEach(btn => btn.style.display = 'none');
      this.btnActivationID.style.display = 'none';
    } else {
      targetingSection.style.display = 'none';
      
      // Show segment buttons
      this.segmentButtons.forEach(btn => btn.style.display = 'block');
      this.btnActivationID.style.display = 'block';
    }
  }

  // Handle selection changes (replicates VBA modeless behavior)
  public async onSelectionChange(): Promise<void> {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load('address, values, rowCount, columnCount');
        await context.sync();

        const parsedData = this.parseCellData(range);
        this.updateInterface(parsedData);
      });
    } catch (error) {
      console.error('Selection change error:', error);
    }
  }

  // Add operation to undo stack (replicates VBA undo system)
  private async addUndoOperation(description: string, range: Excel.Range): Promise<void> {
    const cellChanges: UndoData[] = [];
    const values = range.values as any[][];
    
    // Load all cell addresses at once
    const cells: Excel.Range[] = [];
    for (let row = 0; row < values.length; row++) {
      for (let col = 0; col < values[row].length; col++) {
        const cellValue = values[row][col];
        if (cellValue !== null && cellValue !== undefined && cellValue !== '') {
          const cell = range.getCell(row, col);
          cell.load('address');
          cells.push(cell);
        }
      }
    }
    
    // Sync once to get all addresses
    await range.context.sync();
    
    // Now build the undo data
    let cellIndex = 0;
    for (let row = 0; row < values.length; row++) {
      for (let col = 0; col < values[row].length; col++) {
        const cellValue = values[row][col];
        if (cellValue !== null && cellValue !== undefined && cellValue !== '') {
          cellChanges.push({
            cellAddress: cells[cellIndex].address,
            originalValue: cellValue
          });
          cellIndex++;
        }
      }
    }

    const operation: UndoOperation = {
      description,
      cellChanges,
      cellCount: cellChanges.length,
      operationId: this.nextOperationId++,
      timestamp: new Date()
    };

    this.undoStack.push(operation);

    // Remove oldest operation if over limit
    if (this.undoStack.length > TaxonomyExtractor.MAX_UNDO_OPERATIONS) {
      this.undoStack.shift();
    }

    this.updateUndoButton();
  }

  // Update undo button state (replicates VBA dynamic button captions)
  private updateUndoButton(): void {
    const undoLabel = this.btnUndo.querySelector('.ms-Button-label')!;
    
    if (this.undoStack.length === 0) {
      undoLabel.textContent = 'Undo Last';
      this.btnUndo.disabled = true;
      this.lblUndoWarning.style.display = 'none';
    } else {
      undoLabel.textContent = this.undoStack.length === 1 ? 'Undo Last' : `Undo Last (${this.undoStack.length})`;
      this.btnUndo.disabled = false;
      
      // Show warning when at capacity
      if (this.undoStack.length >= TaxonomyExtractor.MAX_UNDO_OPERATIONS) {
        this.lblUndoWarning.style.display = 'block';
      } else {
        this.lblUndoWarning.style.display = 'none';
      }
    }
  }

  // Show status message
  private showStatus(message: string, isError: boolean = false): void {
    this.statusText.textContent = message;
    this.statusBar.className = `ms-MessageBar ${isError ? 'ms-MessageBar--error' : 'ms-MessageBar--success'}`;
    this.statusBar.style.display = 'block';
    
    // Auto-hide after 3 seconds
    setTimeout(() => {
      this.statusBar.style.display = 'none';
    }, 3000);
  }

  // Extract pipe segment (core extraction logic from VBA)
  public async extractPipeSegment(segmentNumber: number): Promise<void> {
    Logger.info(`Starting extraction of segment ${segmentNumber}`);
    PerformanceMonitor.startOperation(`extract-segment-${segmentNumber}`);
    
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load('address, values');
        await context.sync();

        Logger.debug(`Selected range: ${range.address}, values: ${JSON.stringify(range.values)}`);

        // Add to undo stack BEFORE making changes
        await this.addUndoOperation(`Extract Segment ${segmentNumber}`, range);

        const values = range.values as any[][];
        let processedCount = 0;

        // Process each cell (matching VBA loop logic)
        for (let row = 0; row < values.length; row++) {
          for (let col = 0; col < values[row].length; col++) {
            const cellValue = values[row][col];
            
            if (cellValue && typeof cellValue === 'string' && cellValue.includes('|')) {
              const extracted = this.extractSegmentFromCell(cellValue, segmentNumber);
              if (extracted !== null) {
                values[row][col] = extracted;
                processedCount++;
                Logger.debug(`Extracted "${extracted}" from "${cellValue}"`);
              }
            }
          }
        }

        Logger.info(`Processed ${processedCount} cells`);

        // Update the range with new values
        range.values = values;
        await context.sync();

        // Update UI and show status
        await this.onSelectionChange();
        this.showStatus(`Extracted segment ${segmentNumber} from ${processedCount} cell(s)`);
        
        PerformanceMonitor.endOperation(`extract-segment-${segmentNumber}`);
      });
    } catch (error) {
      Logger.error('Extract segment error:', error);
      this.showStatus('Error extracting segment', true);
      PerformanceMonitor.endOperation(`extract-segment-${segmentNumber}`);
    }
  }

  // Extract segment from individual cell (replicates VBA string manipulation)
  private extractSegmentFromCell(cellText: string, segmentNumber: number): string | null {
    try {
      // Split by colon first to remove activation ID
      const colonIndex = cellText.indexOf(':');
      const mainContent = colonIndex >= 0 ? cellText.substring(0, colonIndex) : cellText;
      
      // Split by pipes
      const segments = mainContent.split('|');
      
      if (segments.length >= segmentNumber) {
        return segments[segmentNumber - 1].trim();
      }
      
      return null;
    } catch (error) {
      return null;
    }
  }

  // Extract activation ID (replicates VBA activation ID logic)
  public async extractActivationID(): Promise<void> {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load('address, values');
        await context.sync();

        // Add to undo stack BEFORE making changes
        await this.addUndoOperation('Extract Activation ID', range);

        const values = range.values as any[][];
        let processedCount = 0;

        // Process each cell
        for (let row = 0; row < values.length; row++) {
          for (let col = 0; col < values[row].length; col++) {
            const cellValue = values[row][col];
            
            if (cellValue && typeof cellValue === 'string' && cellValue.includes(':')) {
              const colonIndex = cellValue.indexOf(':');
              if (colonIndex >= 0 && colonIndex < cellValue.length - 1) {
                const activationId = cellValue.substring(colonIndex + 1).trim();
                values[row][col] = activationId;
                processedCount++;
              }
            }
          }
        }

        // Update the range with new values
        range.values = values;
        await context.sync();

        // Update UI and show status
        await this.onSelectionChange();
        this.showStatus(`Extracted activation ID from ${processedCount} cell(s)`);
      });
    } catch (error) {
      console.error('Extract activation ID error:', error);
      this.showStatus('Error extracting activation ID', true);
    }
  }

  // Clean targeting acronyms (replicates VBA v1.5.0 feature)
  public async cleanTargetingAcronyms(): Promise<void> {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load('address, values');
        await context.sync();

        // Add to undo stack BEFORE making changes
        await this.addUndoOperation('Clean Targeting Acronyms', range);

        const values = range.values as any[][];
        let processedCount = 0;

        // Process each cell with targeting pattern
        for (let row = 0; row < values.length; row++) {
          for (let col = 0; col < values[row].length; col++) {
            const cellValue = values[row][col];
            
            if (cellValue && typeof cellValue === 'string') {
              // Use regex to remove targeting patterns (^ABC^ format)
              const cleaned = cellValue.replace(/\^[^^]+\^ ?/g, '').trim();
              if (cleaned !== cellValue) {
                values[row][col] = cleaned;
                processedCount++;
              }
            }
          }
        }

        // Update the range with new values
        range.values = values;
        await context.sync();

        // Update UI and show status
        await this.onSelectionChange();
        this.showStatus(`Cleaned targeting acronyms from ${processedCount} cell(s)`);
      });
    } catch (error) {
      console.error('Clean targeting acronyms error:', error);
      this.showStatus('Error cleaning targeting acronyms', true);
    }
  }

  // Undo last operation (replicates VBA undo system)
  public async undoLastOperation(): Promise<void> {
    if (this.undoStack.length === 0) return;

    try {
      // Show processing state
      const undoLabel = this.btnUndo.querySelector('.ms-Button-label')!;
      const originalText = undoLabel.textContent;
      undoLabel.textContent = 'Processing...';
      this.btnUndo.disabled = true;

      await Excel.run(async (context) => {
        const lastOperation = this.undoStack.pop()!;
        
        // Restore each cell to its original value
        for (const change of lastOperation.cellChanges) {
          try {
            const cell = context.workbook.worksheets.getActiveWorksheet().getRange(change.cellAddress);
            cell.values = [[change.originalValue]];
          } catch (cellError) {
            console.warn(`Could not restore cell ${change.cellAddress}:`, cellError);
          }
        }

        await context.sync();

        // Update UI
        this.updateUndoButton();
        await this.onSelectionChange();
        this.showStatus(`Undid: ${lastOperation.description} (${lastOperation.cellCount} cells restored)`);
      });
    } catch (error) {
      console.error('Undo operation error:', error);
      this.showStatus('Error during undo operation', true);
      
      // Restore button state on error
      this.updateUndoButton();
    }
  }
}

// Global instances
let taxonomyExtractor: TaxonomyExtractor;

// Make DarkModeManager available globally for debugging
(window as any).DarkModeManager = DarkModeManager;

// Office.js initialization
Office.onReady((info) => {
  Logger.info('Office.js is ready', info);
  
  if (info.host === Office.HostType.Excel) {
    Logger.info('Excel host detected, initializing taxonomy extractor...');
    
    // Initialize the taxonomy extractor
    taxonomyExtractor = new TaxonomyExtractor();
    Logger.info('TaxonomyExtractor instance created');
    
    // Register selection change event (replicates VBA modeless behavior)
    Excel.run(async (context) => {
      Logger.info('Registering selection change event handler');
      context.workbook.worksheets.onSelectionChanged.add(async () => {
        Logger.debug('Selection changed event triggered');
        await taxonomyExtractor.onSelectionChange();
      });
      await context.sync();
      Logger.info('Selection change event handler registered successfully');
    }).catch((error) => {
      Logger.error('Failed to register selection change event handler:', error);
    });

    // Trigger initial selection analysis
    Logger.info('Triggering initial selection analysis');
    taxonomyExtractor.onSelectionChange().catch((error) => {
      Logger.error('Initial selection analysis failed:', error);
    });
    
    Logger.info('IPG Taxonomy Extractor initialization complete');
  } else {
    Logger.warn('Not running in Excel, taxonomy extractor will not be initialized');
  }
});