/**
 * AccessibilityService for IPG Taxonomy Extractor
 * 
 * Implements WCAG 2.1 AA compliance standards:
 * - Screen reader announcements
 * - ARIA label management
 * - Keyboard navigation support
 * - High contrast mode detection
 * - Focus management and restoration
 * - Live region updates for dynamic content
 */

import { localization } from './Localization.service';

export type AnnouncementPriority = 'polite' | 'assertive' | 'off';

export interface AccessibilityOptions {
  enableAnnouncements?: boolean;
  enableFocusManagement?: boolean;
  enableHighContrast?: boolean;
  ariaLiveRegionId?: string;
  debugMode?: boolean;
}

export class AccessibilityService {
  private static instance: AccessibilityService;
  private options: AccessibilityOptions;
  private liveRegion: HTMLElement | null = null;
  private focusHistory: HTMLElement[] = [];
  private isHighContrastMode: boolean = false;
  private mediaQueries: MediaQueryList[] = [];

  public static getInstance(options?: AccessibilityOptions): AccessibilityService {
    if (!AccessibilityService.instance) {
      AccessibilityService.instance = new AccessibilityService(options);
    }
    return AccessibilityService.instance;
  }

  private constructor(options: AccessibilityOptions = {}) {
    this.options = {
      enableAnnouncements: true,
      enableFocusManagement: true,
      enableHighContrast: true,
      ariaLiveRegionId: 'accessibility-announcements',
      debugMode: false,
      ...options
    };

    this.initialize();
  }

  /**
   * Initialize accessibility features
   */
  private initialize(): void {
    if (this.options.enableAnnouncements) {
      this.setupLiveRegion();
    }

    if (this.options.enableHighContrast) {
      this.detectHighContrastMode();
      this.setupHighContrastWatchers();
    }

    this.log('Accessibility service initialized', this.options);
  }

  /**
   * Setup ARIA live region for announcements
   */
  private setupLiveRegion(): void {
    // Check if live region already exists
    this.liveRegion = document.getElementById(this.options.ariaLiveRegionId!);
    
    if (!this.liveRegion) {
      // Create live region
      this.liveRegion = document.createElement('div');
      this.liveRegion.id = this.options.ariaLiveRegionId!;
      this.liveRegion.setAttribute('aria-live', 'polite');
      this.liveRegion.setAttribute('aria-atomic', 'true');
      this.liveRegion.className = 'sr-only'; // Screen reader only
      
      // Style for screen reader only content
      this.liveRegion.style.cssText = `
        position: absolute !important;
        width: 1px !important;
        height: 1px !important;
        padding: 0 !important;
        margin: -1px !important;
        overflow: hidden !important;
        clip: rect(0, 0, 0, 0) !important;
        white-space: nowrap !important;
        border: 0 !important;
      `;
      
      document.body.appendChild(this.liveRegion);
      this.log('ARIA live region created');
    }
  }

  /**
   * Announce message to screen readers
   */
  public announce(message: string, priority: AnnouncementPriority = 'polite'): void {
    if (!this.options.enableAnnouncements || !this.liveRegion) {
      return;
    }

    // Update aria-live attribute based on priority
    this.liveRegion.setAttribute('aria-live', priority);
    
    // Clear and set new message
    this.liveRegion.textContent = '';
    
    // Use setTimeout to ensure screen reader picks up the change
    setTimeout(() => {
      if (this.liveRegion) {
        this.liveRegion.textContent = message;
        this.log(`Announced: "${message}" (${priority})`);
      }
    }, 10);
  }

  /**
   * Announce using localized string key
   */
  public announceKey(key: string, params?: { [key: string]: string | number }, priority: AnnouncementPriority = 'polite'): void {
    const message = localization.getString(key, params);
    this.announce(message, priority);
  }

  /**
   * Set ARIA attributes on element
   */
  public setAriaAttributes(element: HTMLElement, attributes: Record<string, string>): void {
    Object.entries(attributes).forEach(([key, value]) => {
      const ariaKey = key.startsWith('aria-') ? key : `aria-${key}`;
      element.setAttribute(ariaKey, value);
    });
    
    this.log(`Set ARIA attributes on ${element.tagName}`, attributes);
  }

  /**
   * Set ARIA label with localization support
   */
  public setAriaLabel(element: HTMLElement, labelKey: string, params?: { [key: string]: string | number }): void {
    const label = localization.getString(labelKey, params);
    element.setAttribute('aria-label', label);
    
    this.log(`Set ARIA label: "${label}" on ${element.tagName}`);
  }

  /**
   * Set ARIA description with localization support
   */
  public setAriaDescription(element: HTMLElement, descriptionKey: string, params?: { [key: string]: string | number }): void {
    const description = localization.getString(descriptionKey, params);
    element.setAttribute('aria-describedby', description);
    
    this.log(`Set ARIA description: "${description}" on ${element.tagName}`);
  }

  /**
   * Manage focus for keyboard navigation
   */
  public manageFocus(element: HTMLElement, options?: { savePrevious?: boolean; restoreOnEscape?: boolean }): void {
    if (!this.options.enableFocusManagement) {
      return;
    }

    const opts = { savePrevious: true, restoreOnEscape: true, ...options };

    if (opts.savePrevious) {
      const currentFocus = document.activeElement as HTMLElement;
      if (currentFocus && currentFocus !== element) {
        this.focusHistory.push(currentFocus);
      }
    }

    element.focus();
    
    if (opts.restoreOnEscape) {
      const escapeHandler = (event: KeyboardEvent) => {
        if (event.key === 'Escape') {
          this.restoreFocus();
          element.removeEventListener('keydown', escapeHandler);
        }
      };
      element.addEventListener('keydown', escapeHandler);
    }

    this.log(`Focus managed to ${element.tagName}`);
  }

  /**
   * Restore previous focus
   */
  public restoreFocus(): boolean {
    if (this.focusHistory.length === 0) {
      return false;
    }

    const previousElement = this.focusHistory.pop();
    if (previousElement && document.contains(previousElement)) {
      previousElement.focus();
      this.log(`Focus restored to ${previousElement.tagName}`);
      return true;
    }

    return false;
  }

  /**
   * Create accessible button with proper ARIA attributes
   */
  public enhanceButton(button: HTMLButtonElement, options: {
    labelKey: string;
    labelParams?: { [key: string]: string | number };
    descriptionKey?: string;
    descriptionParams?: { [key: string]: string | number };
    role?: string;
    expanded?: boolean;
    pressed?: boolean;
  }): void {
    
    // Set ARIA label
    this.setAriaLabel(button, options.labelKey, options.labelParams);
    
    // Set description if provided
    if (options.descriptionKey) {
      this.setAriaDescription(button, options.descriptionKey, options.descriptionParams);
    }

    // Set role if specified
    if (options.role) {
      button.setAttribute('role', options.role);
    }

    // Set state attributes
    if (options.expanded !== undefined) {
      button.setAttribute('aria-expanded', options.expanded.toString());
    }

    if (options.pressed !== undefined) {
      button.setAttribute('aria-pressed', options.pressed.toString());
    }

    // Ensure button has proper keyboard handling
    if (!button.hasAttribute('tabindex')) {
      button.setAttribute('tabindex', '0');
    }

    this.log(`Enhanced button accessibility for ${button.id || button.className}`);
  }

  /**
   * Create accessible region with proper landmarks
   */
  public createAccessibleRegion(element: HTMLElement, options: {
    role?: string;
    labelKey?: string;
    labelParams?: { [key: string]: string | number };
    live?: AnnouncementPriority;
  }): void {
    
    if (options.role) {
      element.setAttribute('role', options.role);
    }

    if (options.labelKey) {
      this.setAriaLabel(element, options.labelKey, options.labelParams);
    }

    if (options.live) {
      element.setAttribute('aria-live', options.live);
      element.setAttribute('aria-atomic', 'true');
    }

    this.log(`Created accessible region: ${options.role || 'generic'}`);
  }

  /**
   * Detect high contrast mode
   */
  private detectHighContrastMode(): void {
    // Method 1: CSS media query
    const highContrastMediaQuery = window.matchMedia('(prefers-contrast: high)');
    this.isHighContrastMode = highContrastMediaQuery.matches;

    // Method 2: Windows high contrast detection
    if (!this.isHighContrastMode && window.matchMedia) {
      const windowsHighContrast = window.matchMedia('(-ms-high-contrast: active)');
      this.isHighContrastMode = windowsHighContrast.matches;
    }

    // Method 3: Background color detection (fallback)
    if (!this.isHighContrastMode) {
      const testElement = document.createElement('div');
      testElement.style.cssText = `
        background-color: rgb(31, 31, 31);
        position: absolute;
        top: -9999px;
        left: -9999px;
      `;
      document.body.appendChild(testElement);
      
      const computedColor = window.getComputedStyle(testElement).backgroundColor;
      this.isHighContrastMode = computedColor !== 'rgb(31, 31, 31)';
      
      document.body.removeChild(testElement);
    }

    if (this.isHighContrastMode) {
      document.body.classList.add('high-contrast-mode');
      this.log('High contrast mode detected and applied');
    }
  }

  /**
   * Setup high contrast mode watchers
   */
  private setupHighContrastWatchers(): void {
    // Watch for changes in contrast preference
    const contrastMediaQuery = window.matchMedia('(prefers-contrast: high)');
    const windowsHighContrast = window.matchMedia('(-ms-high-contrast: active)');

    const handleContrastChange = () => {
      const wasHighContrast = this.isHighContrastMode;
      this.detectHighContrastMode();
      
      if (wasHighContrast !== this.isHighContrastMode) {
        this.announceKey('ui.announcements.contrast_mode_changed', {
          mode: this.isHighContrastMode ? 'high' : 'normal'
        });
      }
    };

    contrastMediaQuery.addEventListener('change', handleContrastChange);
    windowsHighContrast.addEventListener('change', handleContrastChange);

    this.mediaQueries.push(contrastMediaQuery, windowsHighContrast);
  }

  /**
   * Get current accessibility status
   */
  public getAccessibilityStatus(): {
    announcements: boolean;
    focusManagement: boolean;
    highContrast: boolean;
    liveRegionPresent: boolean;
    focusHistoryLength: number;
  } {
    return {
      announcements: this.options.enableAnnouncements!,
      focusManagement: this.options.enableFocusManagement!,
      highContrast: this.isHighContrastMode,
      liveRegionPresent: this.liveRegion !== null,
      focusHistoryLength: this.focusHistory.length
    };
  }

  /**
   * Check if high contrast mode is active
   */
  public isHighContrast(): boolean {
    return this.isHighContrastMode;
  }

  /**
   * Update button state for screen readers
   */
  public updateButtonState(button: HTMLButtonElement, options: {
    pressed?: boolean;
    expanded?: boolean;
    disabled?: boolean;
    labelKey?: string;
    labelParams?: { [key: string]: string | number };
  }): void {
    
    if (options.pressed !== undefined) {
      button.setAttribute('aria-pressed', options.pressed.toString());
    }

    if (options.expanded !== undefined) {
      button.setAttribute('aria-expanded', options.expanded.toString());
    }

    if (options.disabled !== undefined) {
      button.disabled = options.disabled;
      button.setAttribute('aria-disabled', options.disabled.toString());
    }

    if (options.labelKey) {
      this.setAriaLabel(button, options.labelKey, options.labelParams);
    }
  }

  /**
   * Cleanup accessibility resources
   */
  public destroy(): void {
    // Remove live region
    if (this.liveRegion && this.liveRegion.parentNode) {
      this.liveRegion.parentNode.removeChild(this.liveRegion);
      this.liveRegion = null;
    }

    // Remove media query listeners
    this.mediaQueries.forEach(mq => {
      mq.removeEventListener('change', this.detectHighContrastMode);
    });
    this.mediaQueries = [];

    // Clear focus history
    this.focusHistory = [];

    this.log('Accessibility service destroyed');
  }

  /**
   * Debug logging
   */
  private log(message: string, data?: any): void {
    if (this.options.debugMode) {
      console.log(`[AccessibilityService] ${message}`, data || '');
    }
  }
}

// Export singleton instance
export const accessibility = AccessibilityService.getInstance();