/**
 * ErrorHandler Service for IPG Taxonomy Extractor
 * 
 * Comprehensive error management with:
 * - Error categorization and severity levels
 * - User-friendly error messages with localization
 * - Error reporting and telemetry integration
 * - Recovery strategies and graceful degradation
 * - Development vs production error handling
 */

import { localization } from './Localization.service';
import { accessibility } from './AccessibilityService';

export type ErrorSeverity = 'low' | 'medium' | 'high' | 'critical';
export type ErrorCategory = 'office_api' | 'excel_operation' | 'validation' | 'network' | 'localization' | 'accessibility' | 'component' | 'unknown';

export interface TaxonomyError extends Error {
  category: ErrorCategory;
  severity: ErrorSeverity;
  context: string;
  userMessage?: string;
  technicalDetails?: any;
  timestamp: Date;
  recoverable: boolean;
  errorId: string;
}

export interface ErrorHandlerOptions {
  enableTelemetry?: boolean;
  enableUserNotifications?: boolean;
  enableAccessibilityAnnouncements?: boolean;
  developmentMode?: boolean;
}

export interface ErrorRecoveryStrategy {
  category: ErrorCategory;
  severity: ErrorSeverity;
  action: 'retry' | 'fallback' | 'ignore' | 'reload' | 'notify_only';
  maxRetries?: number;
  fallbackAction?: () => Promise<void>;
}

export class ErrorHandlerService {
  private static instance: ErrorHandlerService;
  private options: ErrorHandlerOptions;
  private errorCount: Map<ErrorCategory, number> = new Map();
  private recentErrors: TaxonomyError[] = [];
  private recoveryStrategies: Map<string, ErrorRecoveryStrategy> = new Map();
  private errorIdCounter = 1;

  public static getInstance(options?: ErrorHandlerOptions): ErrorHandlerService {
    if (!ErrorHandlerService.instance) {
      ErrorHandlerService.instance = new ErrorHandlerService(options);
    }
    return ErrorHandlerService.instance;
  }

  private constructor(options: ErrorHandlerOptions = {}) {
    this.options = {
      enableTelemetry: true,
      enableUserNotifications: true,
      enableAccessibilityAnnouncements: true,
      developmentMode: process.env.NODE_ENV === 'development',
      ...options
    };

    this.setupRecoveryStrategies();
    this.initializeGlobalErrorHandling();
  }

  /**
   * Setup default recovery strategies for different error types
   */
  private setupRecoveryStrategies(): void {
    const strategies: ErrorRecoveryStrategy[] = [
      {
        category: 'office_api',
        severity: 'high',
        action: 'retry',
        maxRetries: 3
      },
      {
        category: 'excel_operation',
        severity: 'medium',
        action: 'retry',
        maxRetries: 2
      },
      {
        category: 'validation',
        severity: 'low',
        action: 'notify_only'
      },
      {
        category: 'network',
        severity: 'medium',
        action: 'fallback',
        maxRetries: 3
      },
      {
        category: 'localization',
        severity: 'low',
        action: 'fallback'
      },
      {
        category: 'accessibility',
        severity: 'medium',
        action: 'fallback'
      },
      {
        category: 'component',
        severity: 'high',
        action: 'reload'
      }
    ];

    strategies.forEach(strategy => {
      const key = `${strategy.category}-${strategy.severity}`;
      this.recoveryStrategies.set(key, strategy);
    });
  }

  /**
   * Initialize global error handling
   */
  private initializeGlobalErrorHandling(): void {
    // Handle unhandled promise rejections
    window.addEventListener('unhandledrejection', (event) => {
      this.handleError(this.createError(
        'Unhandled Promise Rejection',
        'unknown',
        'high',
        'Global promise rejection handler',
        { reason: event.reason }
      ));
      
      if (this.options.developmentMode) {
        console.error('Unhandled promise rejection:', event.reason);
      }
    });

    // Handle general JavaScript errors
    window.addEventListener('error', (event) => {
      this.handleError(this.createError(
        event.message || 'JavaScript Error',
        'unknown',
        'medium',
        'Global error handler',
        {
          filename: event.filename,
          lineno: event.lineno,
          colno: event.colno,
          error: event.error
        }
      ));
    });
  }

  /**
   * Create a standardized taxonomy error
   */
  public createError(
    message: string,
    category: ErrorCategory,
    severity: ErrorSeverity,
    context: string,
    technicalDetails?: any,
    originalError?: Error
  ): TaxonomyError {
    const error = new Error(message) as TaxonomyError;
    error.category = category;
    error.severity = severity;
    error.context = context;
    error.technicalDetails = technicalDetails;
    error.timestamp = new Date();
    error.recoverable = severity !== 'critical';
    error.errorId = `TAX-${Date.now()}-${this.errorIdCounter++}`;

    // Set user-friendly message based on category
    error.userMessage = this.getUserMessage(category, severity, message);

    // Preserve original stack trace if available
    if (originalError) {
      error.stack = originalError.stack;
      // Use safe property assignment for cause (TypeScript compatibility)
      (error as any).cause = originalError;
    }

    return error;
  }

  /**
   * Handle an error with appropriate response strategy
   */
  public async handleError(error: TaxonomyError, retryCount: number = 0): Promise<boolean> {
    // Record error statistics
    this.recordError(error);

    // Log error details
    this.logError(error);

    // Get recovery strategy
    const strategy = this.getRecoveryStrategy(error);

    // Execute recovery strategy
    const recovered = await this.executeRecoveryStrategy(error, strategy, retryCount);

    // Notify user if configured and error is significant
    if (this.options.enableUserNotifications && error.severity !== 'low') {
      this.notifyUser(error, recovered);
    }

    // Announce to screen readers for accessibility
    if (this.options.enableAccessibilityAnnouncements && error.severity === 'high') {
      this.announceError(error);
    }

    return recovered;
  }

  /**
   * Execute error recovery strategy
   */
  private async executeRecoveryStrategy(
    error: TaxonomyError,
    strategy: ErrorRecoveryStrategy,
    retryCount: number
  ): Promise<boolean> {
    try {
      switch (strategy.action) {
        case 'retry':
          if (retryCount < (strategy.maxRetries || 1)) {
            // Exponential backoff for retries
            const delay = Math.pow(2, retryCount) * 1000;
            await this.delay(delay);
            return true; // Indicate retry should be attempted
          }
          return false;

        case 'fallback':
          if (strategy.fallbackAction) {
            await strategy.fallbackAction();
            return true;
          }
          return false;

        case 'ignore':
          return true;

        case 'reload':
          if (error.severity === 'critical') {
            this.scheduleReload();
          }
          return false;

        case 'notify_only':
        default:
          return false;
      }
    } catch (recoveryError) {
      // Recovery itself failed
      this.logError(this.createError(
        'Recovery strategy failed',
        'component',
        'high',
        'ErrorHandler.executeRecoveryStrategy',
        { originalError: error, recoveryError }
      ));
      return false;
    }
  }

  /**
   * Get appropriate recovery strategy for error
   */
  private getRecoveryStrategy(error: TaxonomyError): ErrorRecoveryStrategy {
    const key = `${error.category}-${error.severity}`;
    const strategy = this.recoveryStrategies.get(key);
    
    if (strategy) {
      return strategy;
    }

    // Default strategy based on severity
    return {
      category: error.category,
      severity: error.severity,
      action: error.severity === 'critical' ? 'reload' : 'notify_only'
    };
  }

  /**
   * Get user-friendly error message
   */
  private getUserMessage(category: ErrorCategory, severity: ErrorSeverity, technicalMessage: string): string {
    try {
      // Try to get localized error message
      const key = `errors.${category}_${severity}`;
      const localizedMessage = localization.getString(key);
      
      if (localizedMessage !== key) {
        return localizedMessage;
      }

      // Fallback to category-specific message
      const categoryKey = `errors.${category}`;
      const categoryMessage = localization.getString(categoryKey);
      
      if (categoryMessage !== categoryKey) {
        return categoryMessage;
      }

      // Generic fallback messages
      switch (severity) {
        case 'critical':
          return 'A critical error occurred. The application may need to restart.';
        case 'high':
          return 'An error occurred while processing your request. Please try again.';
        case 'medium':
          return 'There was a problem completing the operation. Some features may not work correctly.';
        case 'low':
        default:
          return 'A minor issue occurred but the application should continue working normally.';
      }
    } catch (localizationError) {
      // Localization itself failed, use technical message
      return technicalMessage || 'An unexpected error occurred.';
    }
  }

  /**
   * Record error for statistics and analysis
   */
  private recordError(error: TaxonomyError): void {
    // Update error counts
    const currentCount = this.errorCount.get(error.category) || 0;
    this.errorCount.set(error.category, currentCount + 1);

    // Keep recent errors (last 50)
    this.recentErrors.push(error);
    if (this.recentErrors.length > 50) {
      this.recentErrors.shift();
    }

    // Send telemetry if enabled (placeholder for future implementation)
    if (this.options.enableTelemetry) {
      this.sendTelemetry(error);
    }
  }

  /**
   * Log error with appropriate level
   */
  private logError(error: TaxonomyError): void {
    const logMessage = `[${error.errorId}] ${error.category.toUpperCase()}: ${error.message}`;
    const logData = {
      severity: error.severity,
      context: error.context,
      timestamp: error.timestamp.toISOString(),
      technicalDetails: error.technicalDetails
    };

    switch (error.severity) {
      case 'critical':
      case 'high':
        console.error(logMessage, logData);
        break;
      case 'medium':
        console.warn(logMessage, logData);
        break;
      case 'low':
      default:
        if (this.options.developmentMode) {
          console.log(logMessage, logData);
        }
        break;
    }
  }

  /**
   * Notify user of error
   */
  private notifyUser(error: TaxonomyError, recovered: boolean): void {
    const message = error.userMessage || error.message;
    const title = recovered ? 'Issue Resolved' : 'Error Occurred';
    
    // In a real implementation, this would show a toast notification
    // For now, we'll use console in development mode
    if (this.options.developmentMode) {
      console.log(`User Notification - ${title}: ${message}`);
    }
  }

  /**
   * Announce error to screen readers
   */
  private announceError(error: TaxonomyError): void {
    try {
      const message = `Error: ${error.userMessage || error.message}`;
      accessibility.announce(message, 'assertive');
    } catch (accessibilityError) {
      // Accessibility announcement failed, log but don't throw
      console.warn('Failed to announce error to screen readers:', accessibilityError);
    }
  }

  /**
   * Send error telemetry (placeholder for future implementation)
   */
  private sendTelemetry(error: TaxonomyError): void {
    if (!this.options.enableTelemetry) return;

    // In a real implementation, this would send to analytics service
    if (this.options.developmentMode) {
      console.log('Telemetry:', {
        errorId: error.errorId,
        category: error.category,
        severity: error.severity,
        context: error.context,
        timestamp: error.timestamp
      });
    }
  }

  /**
   * Schedule application reload for critical errors
   */
  private scheduleReload(): void {
    setTimeout(() => {
      if (confirm('A critical error occurred. Would you like to reload the application?')) {
        window.location.reload();
      }
    }, 1000);
  }

  /**
   * Utility: Delay for retry logic
   */
  private delay(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  /**
   * Get error statistics
   */
  public getErrorStatistics(): {
    totalErrors: number;
    errorsByCategory: Record<ErrorCategory, number>;
    recentErrors: TaxonomyError[];
    criticalErrorCount: number;
  } {
    const errorsByCategory = {} as Record<ErrorCategory, number>;
    this.errorCount.forEach((count, category) => {
      errorsByCategory[category] = count;
    });

    const criticalErrorCount = this.recentErrors.filter(e => e.severity === 'critical').length;

    return {
      totalErrors: this.recentErrors.length,
      errorsByCategory,
      recentErrors: [...this.recentErrors],
      criticalErrorCount
    };
  }

  /**
   * Clear error history
   */
  public clearErrorHistory(): void {
    this.errorCount.clear();
    this.recentErrors = [];
  }

  /**
   * Check if system is in a healthy state
   */
  public isHealthy(): boolean {
    const stats = this.getErrorStatistics();
    return stats.criticalErrorCount === 0 && stats.totalErrors < 10;
  }
}

// Export singleton instance
export const errorHandler = ErrorHandlerService.getInstance();