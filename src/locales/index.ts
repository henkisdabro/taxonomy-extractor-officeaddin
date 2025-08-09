/**
 * Locale management index file
 * 
 * This file provides centralized access to all available locales
 * and manages the locale loading system.
 */

export { LocalizationService, localization } from '../services/Localization.service';

/**
 * Available locale codes
 */
export const AVAILABLE_LOCALES = {
  'en-AU': 'English (Australia)',
  'en-US': 'English (United States)' // Fallback
  // Future locales can be added here:
  // 'es-ES': 'Spanish (Spain)',
  // 'fr-FR': 'French (France)',
} as const;

export type LocaleCode = keyof typeof AVAILABLE_LOCALES;

/**
 * Get the default locale based on Office environment
 * @returns The default locale code
 */
export function getDefaultLocale(): LocaleCode {
  try {
    // Try to get Office locale if available
    if (typeof Office !== 'undefined' && Office.context?.displayLanguage) {
      const officeLocale = Office.context.displayLanguage as LocaleCode;
      if (officeLocale in AVAILABLE_LOCALES) {
        return officeLocale;
      }
    }
  } catch (error) {
    // Fallback to browser locale or default
  }

  // Fallback to browser locale
  const browserLocale = (navigator.language || navigator.languages?.[0]) as LocaleCode;
  if (browserLocale in AVAILABLE_LOCALES) {
    return browserLocale;
  }

  // Ultimate fallback to Australian English
  return 'en-AU';
}