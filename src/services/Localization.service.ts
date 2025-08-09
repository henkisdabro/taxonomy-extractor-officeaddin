/**
 * Lightweight Localization Service for IPG Taxonomy Extractor
 * 
 * Features:
 * - String interpolation with parameters
 * - Fallback to key if translation missing
 * - Minimal bundle size impact
 * - TypeScript support
 */

interface LocaleStrings {
  [key: string]: string;
}

interface InterpolationParams {
  [key: string]: string | number;
}

export class LocalizationService {
  private static instance: LocalizationService;
  private currentLocale: string = 'en-US';
  private strings: Map<string, LocaleStrings> = new Map();
  private isInitialized: boolean = false;

  public static getInstance(): LocalizationService {
    if (!LocalizationService.instance) {
      LocalizationService.instance = new LocalizationService();
    }
    return LocalizationService.instance;
  }

  private constructor() {
    // Private constructor for singleton pattern
  }

  /**
   * Initialize the localization service with default locale
   */
  public async initialize(): Promise<void> {
    if (this.isInitialized) return;

    try {
      // Import the default English locale
      const enUS = await import('../locales/en-US.json');
      this.strings.set('en-US', enUS.default || enUS);
      this.isInitialized = true;
    } catch (error) {
      console.warn('[Localization] Failed to load default locale, using fallback keys');
      this.isInitialized = true; // Still mark as initialized to prevent repeated attempts
    }
  }

  /**
   * Get a localized string with optional parameter interpolation
   * @param key - The translation key
   * @param params - Optional parameters for string interpolation
   * @returns The localized string or the key if not found
   */
  public getString(key: string, params?: InterpolationParams): string {
    const strings = this.strings.get(this.currentLocale);
    let result = strings?.[key] || key;

    // Perform parameter interpolation if params provided
    if (params) {
      Object.entries(params).forEach(([paramKey, value]) => {
        const placeholder = `{${paramKey}}`;
        result = result.replace(new RegExp(placeholder, 'g'), String(value));
      });
    }

    return result;
  }

  /**
   * Set the current locale (future use)
   * @param locale - The locale code (e.g., 'en-US', 'es-ES')
   */
  public setLocale(locale: string): void {
    this.currentLocale = locale;
  }

  /**
   * Get the current locale
   * @returns The current locale code
   */
  public getCurrentLocale(): string {
    return this.currentLocale;
  }

  /**
   * Check if the service is initialized
   * @returns True if initialized
   */
  public getIsInitialized(): boolean {
    return this.isInitialized;
  }
}

// Export singleton instance for convenience
export const localization = LocalizationService.getInstance();