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
  private lblCellCount!: HTMLElement;
  private segmentButtons!: HTMLButtonElement[];
  private btnActivationID!: HTMLButtonElement;
  private btnTargeting!: HTMLButtonElement;
  private btnKeepPattern!: HTMLButtonElement;
  private btnUndo!: HTMLButtonElement;
  private lblUndoWarning!: HTMLElement;
  private statusBar!: HTMLElement;
  private statusText!: HTMLElement;

  constructor() {
    this.initializeUI();
    this.setupEventHandlers();
    // Dark mode disabled - using CSS-only light mode default
    // this.initializeDarkMode();
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
    this.lblCellCount = document.getElementById('lblCellCount')!;
    this.btnActivationID = document.getElementById('btnActivationID') as HTMLButtonElement;
    this.btnTargeting = document.getElementById('btnTargeting') as HTMLButtonElement;
    this.btnKeepPattern = document.getElementById('btnKeepPattern') as HTMLButtonElement;
    this.btnUndo = document.getElementById('btnUndo') as HTMLButtonElement;
    this.lblUndoWarning = document.getElementById('lblUndoWarning')!;
    this.statusBar = document.getElementById('statusBar')!;
    this.statusText = document.getElementById('statusText')!;

    // Verify critical elements
    if (!this.lblCellCount) Logger.error('lblCellCount not found'); 
    if (!this.btnActivationID) Logger.error('btnActivationID not found');
    if (!this.btnTargeting) Logger.error('btnTargeting not found');
    if (!this.btnKeepPattern) Logger.error('btnKeepPattern not found');
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

    // Targeting button (Trim)
    if (this.btnTargeting) {
      this.btnTargeting.addEventListener('click', () => {
        Logger.info('Targeting button clicked');
        this.cleanTargetingAcronyms();
      });
      Logger.debug('Targeting button event handler registered');
    } else {
      Logger.error('Targeting button not found');
    }

    // Keep Pattern button
    if (this.btnKeepPattern) {
      this.btnKeepPattern.addEventListener('click', () => {
        Logger.info('Keep Pattern button clicked');
        this.keepTargetingPattern();
      });
      Logger.debug('Keep Pattern button event handler registered');
    } else {
      Logger.error('Keep Pattern button not found');
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

    // Development mode buttons (only show in dev environment)
    // Add delay to ensure DOM is fully loaded
    setTimeout(() => {
      this.setupDevelopmentMode();
    }, 100);
    
    // Expose dev functions globally for HTML onclick handlers
    (window as any).devSimulate = () => {
      console.log('🚀 devSimulate called from HTML');
      this.simulateSelection('FY24_26|Q1-4|Tourism WA|WA |Always On Remarketing| 4LAOSO | SOC|Facebook_Instagram|Conversions:DJTDOM060725');
    };
    
    (window as any).devTargeting = () => {
      console.log('🎯 devTargeting called from HTML');
      this.simulateTargeting('^ABC^ test string');
    };
    
    (window as any).devClear = () => {
      console.log('🧹 devClear called from HTML');
      this.clearSelection();
    };

    // Alternative names that the robust handler expects
    (window as any).devSimulateData = (window as any).devSimulate;
    (window as any).devSimulateTargeting = (window as any).devTargeting;
    (window as any).devClearData = (window as any).devClear;
    
    console.log('✅ Global dev functions available: devSimulate(), devTargeting(), devClear()');
    console.log('✅ Also available: devSimulateData(), devSimulateTargeting(), devClearData()');
    
    Logger.info('Event handlers setup complete');
  }

  // Setup development mode (only visible in development environment)
  private setupDevelopmentMode(): void {
    // Only enable development mode in local development (localhost or dev server)
    const isDevEnvironment = window.location.hostname === 'localhost' || 
                             window.location.hostname === '127.0.0.1' ||
                             window.location.port === '3001' ||
                             window.location.href.includes('localhost') ||
                             window.location.href.includes('3001');
    
    Logger.info(`Development mode check: hostname=${window.location.hostname}, port=${window.location.port}, href=${window.location.href}, isDev=${isDevEnvironment}`);
    
    // Force dev mode for localhost testing - this ensures dev buttons are always visible during development
    const forceDevForLocalhost = window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1';
    
    if (isDevEnvironment || forceDevForLocalhost) {
      // Add dev-mode class to body to show dev section
      document.body.classList.add('dev-mode');
      Logger.info('Added dev-mode class to body');
      
      // ELITE APPROACH: Multi-layered event handling with global exposure
      this.setupRobustDevHandlers();
      
      Logger.info('Development mode enabled with robust handlers');
    } else {
      Logger.debug('Production mode - development features hidden');
    }
  }

  // Elite-level robust development handler setup
  private setupRobustDevHandlers(): void {
    try {
      // Strategy 1: Use the global functions that have intelligent waiting built-in
      // These are already defined at module level with waitForInstance logic
      
      // Strategy 2: DOM ready state checking with multiple attachment attempts
      const attachHandlers = () => {
        console.log('🔧 Attempting to attach dev button handlers...');
        
        // Simulate button
        const btnDevSimulate = document.getElementById('btnDevSimulate') as HTMLButtonElement;
        if (btnDevSimulate) {
          console.log('✅ Found btnDevSimulate, attaching handlers');
          
          // Multiple attachment strategies
          btnDevSimulate.onclick = (e) => {
            e.preventDefault();
            console.log('📱 Simulate button clicked (onclick)');
            (window as any).devSimulateData();
          };
          
          btnDevSimulate.addEventListener('click', (e) => {
            e.preventDefault();
            console.log('📱 Simulate button clicked (addEventListener)');
            (window as any).devSimulateData();
          });
        } else {
          console.error('❌ btnDevSimulate not found');
        }
        
        // Targeting button  
        const btnDevTargeting = document.getElementById('btnDevTargeting') as HTMLButtonElement;
        if (btnDevTargeting) {
          console.log('✅ Found btnDevTargeting, attaching handlers');
          
          btnDevTargeting.onclick = (e) => {
            e.preventDefault();
            console.log('🎯 Targeting button clicked (onclick)');
            (window as any).devSimulateTargeting();
          };
          
          btnDevTargeting.addEventListener('click', (e) => {
            e.preventDefault();
            console.log('🎯 Targeting button clicked (addEventListener)');
            (window as any).devSimulateTargeting();
          });
        } else {
          console.error('❌ btnDevTargeting not found');
        }
        
        // Clear button
        const btnDevClear = document.getElementById('btnDevClear') as HTMLButtonElement;
        if (btnDevClear) {
          console.log('✅ Found btnDevClear, attaching handlers');
          
          btnDevClear.onclick = (e) => {
            e.preventDefault();
            console.log('🧹 Clear button clicked (onclick)');
            (window as any).devClearData();
          };
          
          btnDevClear.addEventListener('click', (e) => {
            e.preventDefault();
            console.log('🧹 Clear button clicked (addEventListener)');
            (window as any).devClearData();
          });
        } else {
          console.error('❌ btnDevClear not found');
        }
      };
      
      // Strategy 3: Multiple timing attempts
      if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', attachHandlers);
      } else {
        attachHandlers();
      }
      
      // Strategy 4: Backup timer-based attachment
      setTimeout(attachHandlers, 100);
      setTimeout(attachHandlers, 500);
      setTimeout(attachHandlers, 1000);
      
      console.log('🎉 Robust dev handlers setup complete. Try: devSimulateData(), devSimulateTargeting(), devClearData()');
      
    } catch (error) {
      console.error('💥 Error setting up dev handlers:', error);
    }
  }

  // Simulate cell selection with sample data (dev mode only)
  private simulateSelection(sampleData: string): void {
    // Parse the sample data using the same logic as real data
    const segments = sampleData.split('|');
    const lastSegment = segments[segments.length - 1] || '';
    const colonIndex = lastSegment.indexOf(':');
    let activationId = '';
    
    // Extract activation ID if present
    if (colonIndex !== -1) {
      activationId = lastSegment.substring(colonIndex + 1);
      segments[segments.length - 1] = lastSegment.substring(0, colonIndex);
    }
    
    // Create mock parsed data with actual segment parsing
    const mockData: ParsedCellData = {
      originalText: sampleData,
      truncatedDisplay: sampleData,
      selectedCellCount: 1,
      segment1: segments[0]?.trim() || '',
      segment2: segments[1]?.trim() || '', 
      segment3: segments[2]?.trim() || '',
      segment4: segments[3]?.trim() || '',
      segment5: segments[4]?.trim() || '',
      segment6: segments[5]?.trim() || '',
      segment7: segments[6]?.trim() || '',
      segment8: segments[7]?.trim() || '',
      segment9: segments[8]?.trim() || '',
      activationId: activationId.trim(),
      hasTargetingPattern: false, // Set to false so we can see the segment buttons
      targetingText: ''
    };
    
    // Update UI with mock data
    this.updateInterface(mockData);
    Logger.info('Simulated selection with sample data:', mockData);
  }

  // Simulate targeting pattern selection (dev mode only)
  private simulateTargeting(sampleData: string): void {
    // Create mock targeting data
    const mockData: ParsedCellData = {
      originalText: sampleData,
      truncatedDisplay: sampleData,
      selectedCellCount: 1,
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
      hasTargetingPattern: true,
      targetingText: '^ABC^'
    };
    
    // Update UI with targeting data
    this.updateInterface(mockData);
    Logger.info('Simulated targeting pattern with sample data:', mockData);
  }

  // Clear selection (dev mode only)
  private clearSelection(): void {
    const emptyData: ParsedCellData = {
      originalText: '',
      truncatedDisplay: '',
      selectedCellCount: 0,
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
    
    this.updateInterface(emptyData);
    Logger.info('Cleared selection simulation');
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
    // Update cell count with status
    if (parsedData.selectedCellCount === 0) {
      this.lblCellCount.textContent = 'No cells selected';
    } else if (parsedData.hasTargetingPattern) {
      this.lblCellCount.textContent = `${parsedData.selectedCellCount} cell(s) selected - Targeting pattern detected`;
    } else if (parsedData.originalText.includes('|')) {
      this.lblCellCount.textContent = `${parsedData.selectedCellCount} cell(s) selected - Taxonomy data detected`;
    } else {
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
        buttonLabel.textContent = segment.length > 35 ? segment.substring(0, 35) + '...' : segment;
        button.disabled = false;
        button.classList.remove('ms-Button--default');
        button.classList.add('ms-Button--primary');
      } else {
        buttonLabel.textContent = 'N/A';
        button.disabled = true;
        button.classList.remove('ms-Button--primary');
        button.classList.add('ms-Button--default');
      }
    });

    // Update activation ID button
    const activationLabel = this.btnActivationID.querySelector('.ms-Button-label')!;
    if (parsedData.activationId) {
      activationLabel.textContent = parsedData.activationId;
      this.btnActivationID.disabled = false;
    } else {
      activationLabel.textContent = 'N/A';
      this.btnActivationID.disabled = true;
    }

    // Handle targeting buttons visibility (VBA overlay behavior)
    const targetingLabel = this.btnTargeting.querySelector('.ms-Button-label')!;
    const keepPatternLabel = this.btnKeepPattern.querySelector('.ms-Button-label')!;
    const targetingSection = document.getElementById('targetingSection')!;
    const sectionHeader = document.querySelector('.section-header') as HTMLElement;
    const specialButtons = document.querySelector('.special-buttons') as HTMLElement;
    
    if (parsedData.hasTargetingPattern) {
      targetingLabel.textContent = parsedData.targetingText;
      keepPatternLabel.textContent = parsedData.targetingText;
      targetingSection.style.display = 'block';
      this.btnTargeting.disabled = false;
      this.btnKeepPattern.disabled = false;
      
      // Add targeting-mode class to body to enable CSS targeting rules
      document.body.classList.add('targeting-mode');
      
      // Hide segment buttons when targeting pattern detected
      this.segmentButtons.forEach(btn => btn.style.display = 'none');
      this.btnActivationID.style.display = 'none';
      
      // Hide Extract Segments header and activation ID section when targeting
      if (sectionHeader) sectionHeader.style.display = 'none';
      if (specialButtons) specialButtons.style.display = 'none';
    } else {
      targetingLabel.textContent = 'N/A';
      keepPatternLabel.textContent = 'N/A';
      targetingSection.style.display = 'none';
      this.btnTargeting.disabled = true;
      this.btnKeepPattern.disabled = true;
      
      // Remove targeting-mode class from body
      document.body.classList.remove('targeting-mode');
      
      // Show segment buttons
      this.segmentButtons.forEach(btn => btn.style.display = 'block');
      this.btnActivationID.style.display = 'block';
      
      // Show Extract Segments header and activation ID section
      if (sectionHeader) sectionHeader.style.display = 'block';
      if (specialButtons) specialButtons.style.display = 'block';
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
      this.btnUndo.classList.remove('undo-available');
      this.lblUndoWarning.style.display = 'none';
    } else {
      undoLabel.textContent = this.undoStack.length === 1 ? 'Undo Last' : `Undo Last (${this.undoStack.length})`;
      this.btnUndo.disabled = false;
      this.btnUndo.classList.add('undo-available');
      
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

  // Keep only targeting pattern (^ABC^), remove everything else
  public async keepTargetingPattern(): Promise<void> {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load('address, values');
        await context.sync();

        // Add to undo stack BEFORE making changes
        await this.addUndoOperation('Keep Targeting Pattern', range);

        const values = range.values as any[][];
        let processedCount = 0;

        // Process each cell with targeting pattern
        for (let row = 0; row < values.length; row++) {
          for (let col = 0; col < values[row].length; col++) {
            const cellValue = values[row][col];
            
            if (cellValue && typeof cellValue === 'string') {
              // Extract only the targeting patterns (^ABC^ format)
              const matches = cellValue.match(/\^[^^]+\^/g);
              if (matches && matches.length > 0) {
                // Join all patterns found with spaces
                const keptPattern = matches.join(' ').trim();
                if (keptPattern !== cellValue) {
                  values[row][col] = keptPattern;
                  processedCount++;
                }
              }
            }
          }
        }

        // Update the range with new values
        range.values = values;
        await context.sync();

        // Update UI and show status
        await this.onSelectionChange();
        this.showStatus(`Kept targeting patterns from ${processedCount} cell(s)`);
      });
    } catch (error) {
      console.error('Keep targeting pattern error:', error);
      this.showStatus('Error keeping targeting pattern', true);
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

// Check if we're in development environment
const isDevEnv = () => {
  const isDev = window.location.hostname === 'localhost' || 
                window.location.hostname === '127.0.0.1' ||
                window.location.port === '3001' ||
                window.location.href.includes('localhost') ||
                window.location.href.includes('3001');
  
  // Debug logging to help troubleshoot
  console.log(`🔍 Dev environment check: hostname=${window.location.hostname}, port=${window.location.port}, isDev=${isDev}`);
  
  return isDev;
};

// Helper function to wait for taxonomyExtractor or create fallback instance (DEV ONLY)
const waitForInstance = (callback: () => void, maxAttempts: number = 50) => {
  if (!isDevEnv()) {
    console.warn('🚫 Dev functions disabled in production environment');
    return;
  }
  
  let attempts = 0;
  const checkInstance = () => {
    attempts++;
    if (taxonomyExtractor) {
      console.log(`✅ Instance ready after ${attempts} attempts`);
      callback();
    } else if (attempts < maxAttempts) {
      console.log(`⏳ Waiting for instance... attempt ${attempts}`);
      setTimeout(checkInstance, 100);
    } else {
      console.log('⚠️ Office.js timeout - creating standalone instance for browser testing');
      // Create a fallback instance for standalone browser testing
      taxonomyExtractor = new TaxonomyExtractor();
      console.log('🔧 Fallback taxonomyExtractor instance created');
      callback();
    }
  };
  checkInstance();
};

// GLOBAL DEV FUNCTIONS - Only exposed in development environment
if (isDevEnv()) {
  (window as any).devSimulate = () => {
    console.log('🚀 devSimulate called');
    waitForInstance(() => {
      (taxonomyExtractor as any).simulateSelection('FY24_26|Q1-4|Tourism WA|WA |Always On Remarketing| 4LAOSO | SOC|Facebook_Instagram|Conversions:DJTDOM060725');
    });
  };

  (window as any).devTargeting = () => {
    console.log('🎯 devTargeting called');
    waitForInstance(() => {
      (taxonomyExtractor as any).simulateTargeting('^ABC^ test string');
    });
  };

  (window as any).devClear = () => {
    console.log('🧹 devClear called');
    waitForInstance(() => {
      (taxonomyExtractor as any).clearSelection();
    });
  };

  // Alternative names for the robust handler
  (window as any).devSimulateData = (window as any).devSimulate;
  (window as any).devSimulateTargeting = (window as any).devTargeting;
  (window as any).devClearData = (window as any).devClear;

  console.log('🌍 Global dev functions exposed:', ['devSimulate', 'devTargeting', 'devClear', 'devSimulateData', 'devSimulateTargeting', 'devClearData']);
  
  // FORCE DEV MODE UI - Make dev section visible immediately in browser testing
  document.addEventListener('DOMContentLoaded', () => {
    document.body.classList.add('dev-mode');
    console.log('🟡 Dev mode UI enabled - yellow section should be visible');
    
    // Attach direct event handlers for immediate functionality - NO TIMEOUT
    const attachDevHandlers = () => {
      const btnDevSimulate = document.getElementById('btnDevSimulate') as HTMLButtonElement;
      const btnDevTargeting = document.getElementById('btnDevTargeting') as HTMLButtonElement;
      const btnDevClear = document.getElementById('btnDevClear') as HTMLButtonElement;
      
      if (btnDevSimulate) {
        btnDevSimulate.onclick = () => {
          console.log('📱 Direct dev simulate clicked');
          (window as any).devSimulate();
        };
        console.log('✅ Dev simulate handler attached');
      }
      
      if (btnDevTargeting) {
        btnDevTargeting.onclick = () => {
          console.log('🎯 Direct dev targeting clicked');
          (window as any).devTargeting();
        };
        console.log('✅ Dev targeting handler attached');
      }
      
      if (btnDevClear) {
        btnDevClear.onclick = () => {
          console.log('🧹 Direct dev clear clicked');
          (window as any).devClear();
        };
        console.log('✅ Dev clear handler attached');
      }
      
      console.log('⚡ Direct dev button handlers attached immediately');
    };
    
    // Try immediately, then retry with setTimeout as fallback
    attachDevHandlers();
    setTimeout(attachDevHandlers, 50);
    setTimeout(attachDevHandlers, 200);
  });
  
  // Also try immediately in case DOM is already loaded
  if (document.readyState === 'loading') {
    // DOM still loading, event listener above will handle it
  } else {
    document.body.classList.add('dev-mode');
    console.log('🟡 Dev mode UI enabled immediately - yellow section should be visible');
    
    // Attach handlers immediately if DOM is ready
    const attachDevHandlers = () => {
      const btnDevSimulate = document.getElementById('btnDevSimulate') as HTMLButtonElement;
      const btnDevTargeting = document.getElementById('btnDevTargeting') as HTMLButtonElement;
      const btnDevClear = document.getElementById('btnDevClear') as HTMLButtonElement;
      
      if (btnDevSimulate && !btnDevSimulate.onclick) {
        btnDevSimulate.onclick = () => (window as any).devSimulate();
        console.log('✅ Immediate dev simulate handler attached');
      }
      if (btnDevTargeting && !btnDevTargeting.onclick) {
        btnDevTargeting.onclick = () => (window as any).devTargeting();
        console.log('✅ Immediate dev targeting handler attached');
      }
      if (btnDevClear && !btnDevClear.onclick) {
        btnDevClear.onclick = () => (window as any).devClear();
        console.log('✅ Immediate dev clear handler attached');
      }
    };
    
    attachDevHandlers();
  }
} else {
  console.log('🔒 Production mode: Dev functions disabled');
}

// Office.js initialization
Office.onReady((info) => {
  Logger.info('Office.js is ready', info);
  
  if (info.host === Office.HostType.Excel) {
    Logger.info('Excel host detected, initializing taxonomy extractor...');
    
    // Initialize the taxonomy extractor
    taxonomyExtractor = new TaxonomyExtractor();
    Logger.info('TaxonomyExtractor instance created');
    console.log('🎉 taxonomyExtractor instance created - global dev functions now ready!');
    
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