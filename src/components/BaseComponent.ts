/**
 * Base Component for IPG Taxonomy Extractor
 * 
 * Provides common functionality for all UI components:
 * - Event handler management
 * - State subscription
 * - Localization access
 * - Lifecycle management
 * - Error handling
 */

import { BaseComponentProps, StateChangeListener, LogLevel } from '../types/taxonomy.types';
import { localization } from '../services/Localization.service';
import { accessibility } from '../services/AccessibilityService';
import { errorHandler } from '../services/ErrorHandler.service';

export abstract class BaseComponent<TProps extends BaseComponentProps = BaseComponentProps> {
  protected element: HTMLElement;
  protected props: TProps;
  protected eventHandlers: Map<string, EventListener> = new Map();
  protected isDestroyed: boolean = false;
  private stateUnsubscribe?: () => void;

  constructor(element: HTMLElement, props: TProps) {
    if (!element) {
      throw new Error('BaseComponent requires a valid HTMLElement');
    }

    this.element = element;
    this.props = { ...props };
    this.initialize();
  }

  /**
   * Initialize the component - override in subclasses
   */
  protected abstract initialize(): void;

  /**
   * Render the component - override in subclasses
   */
  protected abstract render(): void;

  /**
   * Safely add event listener with automatic cleanup
   */
  protected addEventListener(
    target: HTMLElement | Window | Document,
    event: string, 
    handler: EventListener,
    options?: boolean | AddEventListenerOptions
  ): void {
    if (this.isDestroyed) return;

    const key = `${target.constructor.name}-${event}`;
    target.addEventListener(event, handler, options);
    this.eventHandlers.set(key, handler);
  }

  /**
   * Subscribe to state changes with automatic cleanup
   */
  protected subscribeToState(listener: StateChangeListener): void {
    if (this.props.stateManager && !this.isDestroyed) {
      this.stateUnsubscribe = this.props.stateManager.subscribe(listener);
    }
  }

  /**
   * Get localized string with fallback
   */
  protected getString(key: string, params?: { [key: string]: string | number }): string {
    try {
      return localization.getString(key, params);
    } catch (error) {
      this.log('WARN', `Failed to get localized string for key: ${key}`, error);
      return key; // Fallback to key
    }
  }

  /**
   * Enhanced accessibility methods
   */

  /**
   * Set ARIA label with localization support
   */
  protected setAriaLabel(element: HTMLElement, labelKey: string, params?: { [key: string]: string | number }): void {
    accessibility.setAriaLabel(element, labelKey, params);
  }

  /**
   * Set ARIA description with localization support
   */
  protected setAriaDescription(element: HTMLElement, descriptionKey: string, params?: { [key: string]: string | number }): void {
    accessibility.setAriaDescription(element, descriptionKey, params);
  }

  /**
   * Announce message to screen readers
   */
  protected announce(message: string, priority: 'polite' | 'assertive' | 'off' = 'polite'): void {
    accessibility.announce(message, priority);
  }

  /**
   * Announce using localization key
   */
  protected announceKey(key: string, params?: { [key: string]: string | number }, priority: 'polite' | 'assertive' | 'off' = 'polite'): void {
    accessibility.announceKey(key, params, priority);
  }

  /**
   * Enhance button with accessibility features
   */
  protected enhanceButtonAccessibility(button: HTMLButtonElement, options: {
    labelKey: string;
    labelParams?: { [key: string]: string | number };
    descriptionKey?: string;
    descriptionParams?: { [key: string]: string | number };
    role?: string;
    expanded?: boolean;
    pressed?: boolean;
  }): void {
    accessibility.enhanceButton(button, options);
  }

  /**
   * Update button accessibility state
   */
  protected updateButtonAccessibilityState(button: HTMLButtonElement, options: {
    pressed?: boolean;
    expanded?: boolean;
    disabled?: boolean;
    labelKey?: string;
    labelParams?: { [key: string]: string | number };
  }): void {
    accessibility.updateButtonState(button, options);
  }

  /**
   * Manage focus with accessibility considerations
   */
  protected manageFocus(element: HTMLElement, options?: { savePrevious?: boolean; restoreOnEscape?: boolean }): void {
    accessibility.manageFocus(element, options);
  }

  /**
   * Enhanced error handling methods
   */

  /**
   * Handle error with proper categorization and user feedback
   */
  protected handleError(
    error: Error | string,
    category: 'office_api' | 'excel_operation' | 'validation' | 'component' = 'component',
    severity: 'low' | 'medium' | 'high' | 'critical' = 'medium',
    context?: string
  ): void {
    const taxonomyError = typeof error === 'string' 
      ? errorHandler.createError(error, category, severity, context || this.constructor.name)
      : errorHandler.createError(error.message, category, severity, context || this.constructor.name, undefined, error);

    errorHandler.handleError(taxonomyError);
  }

  /**
   * Safe async operation wrapper with error handling
   */
  protected async safeExecute<T>(
    operation: () => Promise<T>,
    errorMessage: string = 'Operation failed',
    category: 'office_api' | 'excel_operation' | 'validation' | 'component' = 'component',
    fallbackValue?: T
  ): Promise<T | undefined> {
    try {
      return await operation();
    } catch (error) {
      this.handleError(error instanceof Error ? error : new Error(String(error)), category, 'medium');
      return fallbackValue;
    }
  }

  /**
   * Validate component state before operations
   */
  protected validateState(): boolean {
    if (this.isDestroyed) {
      this.handleError('Attempted operation on destroyed component', 'component', 'high');
      return false;
    }

    if (!this.element) {
      this.handleError('Component element is not available', 'component', 'high');
      return false;
    }

    return true;
  }

  /**
   * Log with component context
   */
  protected log(level: LogLevel, message: string, data?: any): void {
    const componentName = this.constructor.name;
    const prefixedMessage = `[${componentName}] ${message}`;
    
    switch (level) {
      case 'ERROR':
        console.error(prefixedMessage, data || '');
        break;
      case 'WARN':
        console.warn(prefixedMessage, data || '');
        break;
      case 'INFO':
        console.log(prefixedMessage, data || '');
        break;
      case 'DEBUG':
        console.debug(prefixedMessage, data || '');
        break;
    }
  }

  /**
   * Safely query for child elements
   */
  protected querySelector<T extends HTMLElement = HTMLElement>(selector: string): T | null {
    try {
      return this.element.querySelector(selector) as T;
    } catch (error) {
      this.log('WARN', `Failed to query selector: ${selector}`, error);
      return null;
    }
  }

  /**
   * Safely query for all matching child elements
   */
  protected querySelectorAll<T extends HTMLElement = HTMLElement>(selector: string): T[] {
    try {
      return Array.from(this.element.querySelectorAll(selector)) as T[];
    } catch (error) {
      this.log('WARN', `Failed to query selector all: ${selector}`, error);
      return [];
    }
  }

  /**
   * Update component props and re-render
   */
  public updateProps(newProps: Partial<TProps>): void {
    if (this.isDestroyed) return;

    this.props = { ...this.props, ...newProps };
    this.render();
  }

  /**
   * Get current props (readonly)
   */
  public getProps(): Readonly<TProps> {
    return { ...this.props };
  }

  /**
   * Check if component is destroyed
   */
  public getIsDestroyed(): boolean {
    return this.isDestroyed;
  }

  /**
   * Destroy component and clean up all resources
   */
  public destroy(): void {
    if (this.isDestroyed) return;

    // Remove all event listeners
    this.eventHandlers.forEach((handler, key) => {
      try {
        // Parse the key to get target type and event
        const [targetType, event] = key.split('-');
        let target: EventTarget | null = null;

        switch (targetType) {
          case 'HTMLElement':
          case 'HTMLButtonElement':
          case 'HTMLDivElement':
          case 'HTMLSpanElement':
            target = this.element;
            break;
          case 'Window':
            target = window;
            break;
          case 'Document':
            target = document;
            break;
          default:
            target = this.element; // Fallback to element
            break;
        }

        if (target) {
          target.removeEventListener(event, handler);
        }
      } catch (error) {
        this.log('WARN', `Failed to remove event listener: ${key}`, error);
      }
    });

    this.eventHandlers.clear();

    // Unsubscribe from state changes
    if (this.stateUnsubscribe) {
      this.stateUnsubscribe();
      this.stateUnsubscribe = undefined;
    }

    // Mark as destroyed
    this.isDestroyed = true;

    this.log('DEBUG', 'Component destroyed');
  }

  /**
   * Setup keyboard navigation support
   */
  protected setupKeyboardNavigation(): void {
    this.addEventListener(this.element, 'keydown', (event: Event) => {
      const keyboardEvent = event as KeyboardEvent;
      switch (keyboardEvent.key) {
        case 'Enter':
        case ' ':
          keyboardEvent.preventDefault();
          this.handleActivation();
          break;
        case 'Escape':
          this.handleCancel?.();
          break;
        case 'Tab':
          // Let browser handle tab navigation
          break;
      }
    });
  }

  /**
   * Handle activation (Enter/Space) - override in subclasses
   */
  protected handleActivation(): void {
    // Default implementation - override in subclasses
    this.log('DEBUG', 'Activation not implemented');
  }

  /**
   * Handle cancel (Escape) - override in subclasses
   */
  protected handleCancel?(): void;

  /**
   * Safe element property setter
   */
  protected setElementProperty(property: string, value: any): void {
    try {
      (this.element as any)[property] = value;
    } catch (error) {
      this.log('WARN', `Failed to set element property: ${property}`, error);
    }
  }

  /**
   * Safe element text content setter
   */
  protected setTextContent(text: string): void {
    try {
      this.element.textContent = text;
    } catch (error) {
      this.log('WARN', 'Failed to set text content', error);
    }
  }

  /**
   * Safe element attribute setter
   */
  protected setAttribute(name: string, value: string): void {
    try {
      this.element.setAttribute(name, value);
    } catch (error) {
      this.log('WARN', `Failed to set attribute: ${name}`, error);
    }
  }

  /**
   * Safe element class list manipulation
   */
  protected addClass(className: string): void {
    try {
      this.element.classList.add(className);
    } catch (error) {
      this.log('WARN', `Failed to add class: ${className}`, error);
    }
  }

  protected removeClass(className: string): void {
    try {
      this.element.classList.remove(className);
    } catch (error) {
      this.log('WARN', `Failed to remove class: ${className}`, error);
    }
  }

  protected toggleClass(className: string, force?: boolean): void {
    try {
      this.element.classList.toggle(className, force);
    } catch (error) {
      this.log('WARN', `Failed to toggle class: ${className}`, error);
    }
  }
}