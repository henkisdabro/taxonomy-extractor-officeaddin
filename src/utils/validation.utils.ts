/**
 * Validation utilities for IPG Taxonomy Extractor
 * 
 * Comprehensive input validation and sanitization:
 * - Cell data validation and sanitization
 * - Taxonomy format validation
 * - Excel range validation
 * - Security-focused input sanitization
 * - Type safety and error handling
 */

import { errorHandler } from '../services/ErrorHandler.service';

export interface ValidationResult<T = any> {
  isValid: boolean;
  data?: T;
  errors: string[];
  warnings: string[];
  sanitizedData?: T;
}

export interface TaxonomyValidationRules {
  maxSegments?: number;
  maxSegmentLength?: number;
  maxActivationIdLength?: number;
  allowEmptySegments?: boolean;
  requiredSegments?: number[];
  allowSpecialCharacters?: boolean;
}

export class ValidationUtils {
  
  // Default validation rules
  private static readonly DEFAULT_RULES: TaxonomyValidationRules = {
    maxSegments: 9,
    maxSegmentLength: 100,
    maxActivationIdLength: 50,
    allowEmptySegments: true,
    requiredSegments: [],
    allowSpecialCharacters: true
  };

  /**
   * Validate and sanitize cell value for taxonomy processing
   */
  public static validateCellValue(value: any): ValidationResult<string> {
    const result: ValidationResult<string> = {
      isValid: false,
      errors: [],
      warnings: [],
      data: undefined,
      sanitizedData: undefined
    };

    try {
      // Check if value exists
      if (value === null || value === undefined) {
        result.errors.push('Cell value is null or undefined');
        return result;
      }

      // Convert to string and sanitize
      let stringValue = String(value).trim();

      // Check for empty string after trimming
      if (stringValue.length === 0) {
        result.errors.push('Cell value is empty after trimming');
        return result;
      }

      // Basic XSS protection - remove potentially dangerous characters
      const sanitized = this.sanitizeString(stringValue);

      // Check for excessive length
      if (sanitized.length > 1000) {
        result.errors.push('Cell value exceeds maximum length (1000 characters)');
        return result;
      }

      // Validate character set (allow Unicode but warn about potential issues)
      if (this.hasProblematicCharacters(sanitized)) {
        result.warnings.push('Cell contains characters that may cause processing issues');
      }

      result.isValid = true;
      result.data = stringValue;
      result.sanitizedData = sanitized;

    } catch (error) {
      result.errors.push(`Validation error: ${error instanceof Error ? error.message : 'Unknown error'}`);
      errorHandler.handleError(errorHandler.createError(
        'Cell value validation failed',
        'validation',
        'medium',
        'ValidationUtils.validateCellValue',
        { originalValue: value, error }
      ));
    }

    return result;
  }

  /**
   * Validate taxonomy data format
   */
  public static validateTaxonomyData(data: string, rules?: Partial<TaxonomyValidationRules>): ValidationResult<{
    segments: string[];
    activationId: string;
    isValid: boolean;
  }> {
    const validationRules = { ...this.DEFAULT_RULES, ...rules };
    const result: ValidationResult = {
      isValid: false,
      errors: [],
      warnings: [],
      data: undefined
    };

    try {
      // First validate the input string
      const cellValidation = this.validateCellValue(data);
      if (!cellValidation.isValid) {
        result.errors.push(...cellValidation.errors);
        result.warnings.push(...cellValidation.warnings);
        return result;
      }

      const sanitizedData = cellValidation.sanitizedData!;

      // Split by colon to separate main content and activation ID
      const colonParts = sanitizedData.split(':');
      const mainContent = colonParts[0];
      const activationId = colonParts.length > 1 ? colonParts[1].trim() : '';

      // Validate activation ID length
      if (activationId.length > validationRules.maxActivationIdLength!) {
        result.errors.push(`Activation ID exceeds maximum length (${validationRules.maxActivationIdLength} characters)`);
      }

      // Split main content by pipes
      const segments = mainContent.split('|').map(segment => segment.trim());

      // Validate segment count
      if (segments.length > validationRules.maxSegments!) {
        result.warnings.push(`Found ${segments.length} segments, maximum expected is ${validationRules.maxSegments}`);
      }

      // Validate individual segments
      const validatedSegments: string[] = [];
      
      for (let i = 0; i < segments.length; i++) {
        const segment = segments[i];
        
        // Check segment length
        if (segment.length > validationRules.maxSegmentLength!) {
          result.errors.push(`Segment ${i + 1} exceeds maximum length (${validationRules.maxSegmentLength} characters)`);
          continue;
        }

        // Check for empty segments if not allowed
        if (!validationRules.allowEmptySegments && segment.length === 0) {
          result.errors.push(`Segment ${i + 1} is empty (not allowed by current rules)`);
          continue;
        }

        // Sanitize segment
        const sanitizedSegment = this.sanitizeString(segment);
        validatedSegments.push(sanitizedSegment);
      }

      // Check required segments
      if (validationRules.requiredSegments && validationRules.requiredSegments.length > 0) {
        for (const requiredIndex of validationRules.requiredSegments) {
          if (requiredIndex <= validatedSegments.length) {
            const segment = validatedSegments[requiredIndex - 1];
            if (!segment || segment.length === 0) {
              result.errors.push(`Required segment ${requiredIndex} is missing or empty`);
            }
          } else {
            result.errors.push(`Required segment ${requiredIndex} is not present (only ${validatedSegments.length} segments found)`);
          }
        }
      }

      // If no errors, mark as valid
      if (result.errors.length === 0) {
        result.isValid = true;
        result.data = {
          segments: validatedSegments,
          activationId: this.sanitizeString(activationId),
          isValid: true
        };
      }

    } catch (error) {
      result.errors.push(`Taxonomy validation error: ${error instanceof Error ? error.message : 'Unknown error'}`);
      errorHandler.handleError(errorHandler.createError(
        'Taxonomy data validation failed',
        'validation',
        'medium',
        'ValidationUtils.validateTaxonomyData',
        { data, rules, error }
      ));
    }

    return result;
  }

  /**
   * Validate targeting pattern (^ABC^ format)
   */
  public static validateTargetingPattern(text: string): ValidationResult<string[]> {
    const result: ValidationResult<string[]> = {
      isValid: false,
      errors: [],
      warnings: [],
      data: undefined
    };

    try {
      // First validate the input string
      const cellValidation = this.validateCellValue(text);
      if (!cellValidation.isValid) {
        result.errors.push(...cellValidation.errors);
        return result;
      }

      const sanitizedText = cellValidation.sanitizedData!;

      // Check for targeting pattern
      const targetingRegex = /\^[^^]+\^/g;
      const matches = sanitizedText.match(targetingRegex);

      if (!matches || matches.length === 0) {
        result.errors.push('No targeting patterns found in text');
        return result;
      }

      // Validate each pattern
      const validPatterns: string[] = [];
      
      for (const pattern of matches) {
        // Check pattern length
        if (pattern.length < 3) { // Minimum ^A^
          result.warnings.push(`Pattern "${pattern}" is very short`);
        }
        
        if (pattern.length > 20) { // Reasonable maximum
          result.warnings.push(`Pattern "${pattern}" is unusually long`);
        }

        // Extract content between ^ characters
        const content = pattern.slice(1, -1);
        
        // Validate content
        if (content.length === 0) {
          result.errors.push(`Empty targeting pattern found: ${pattern}`);
          continue;
        }

        // Check for nested ^ characters
        if (content.includes('^')) {
          result.errors.push(`Invalid nested ^ character in pattern: ${pattern}`);
          continue;
        }

        validPatterns.push(pattern);
      }

      if (result.errors.length === 0) {
        result.isValid = true;
        result.data = validPatterns;
      }

    } catch (error) {
      result.errors.push(`Targeting pattern validation error: ${error instanceof Error ? error.message : 'Unknown error'}`);
      errorHandler.handleError(errorHandler.createError(
        'Targeting pattern validation failed',
        'validation',
        'low',
        'ValidationUtils.validateTargetingPattern',
        { text, error }
      ));
    }

    return result;
  }

  /**
   * Validate Excel range address
   */
  public static validateExcelRange(range: any): ValidationResult<{
    address: string;
    rowCount: number;
    columnCount: number;
  }> {
    const result: ValidationResult = {
      isValid: false,
      errors: [],
      warnings: [],
      data: undefined
    };

    try {
      if (!range) {
        result.errors.push('Range object is null or undefined');
        return result;
      }

      // Check required properties
      if (!range.address || typeof range.address !== 'string') {
        result.errors.push('Range address is missing or invalid');
        return result;
      }

      if (typeof range.rowCount !== 'number' || range.rowCount <= 0) {
        result.errors.push('Range rowCount is missing or invalid');
        return result;
      }

      if (typeof range.columnCount !== 'number' || range.columnCount <= 0) {
        result.errors.push('Range columnCount is missing or invalid');
        return result;
      }

      // Check for reasonable limits
      if (range.rowCount > 1000000) {
        result.errors.push('Range has too many rows (maximum 1,000,000)');
        return result;
      }

      if (range.columnCount > 100) {
        result.errors.push('Range has too many columns (maximum 100)');
        return result;
      }

      // Warn about large selections
      const totalCells = range.rowCount * range.columnCount;
      if (totalCells > 10000) {
        result.warnings.push(`Large selection (${totalCells} cells) may impact performance`);
      }

      result.isValid = true;
      result.data = {
        address: range.address,
        rowCount: range.rowCount,
        columnCount: range.columnCount
      };

    } catch (error) {
      result.errors.push(`Range validation error: ${error instanceof Error ? error.message : 'Unknown error'}`);
      errorHandler.handleError(errorHandler.createError(
        'Excel range validation failed',
        'validation',
        'high',
        'ValidationUtils.validateExcelRange',
        { range, error }
      ));
    }

    return result;
  }

  /**
   * Sanitize string input to prevent XSS and other security issues
   */
  private static sanitizeString(input: string): string {
    if (!input || typeof input !== 'string') {
      return '';
    }

    return input
      .replace(/[<>]/g, '') // Remove < and > to prevent HTML injection
      .replace(/javascript:/gi, '') // Remove javascript: protocol
      .replace(/on\w+\s*=/gi, '') // Remove event handlers
      .trim();
  }

  /**
   * Check for potentially problematic characters
   */
  private static hasProblematicCharacters(text: string): boolean {
    // Check for control characters (except common whitespace)
    const controlCharRegex = /[\x00-\x08\x0B-\x1F\x7F]/;
    
    // Check for unusual Unicode characters that might cause issues
    const problematicUnicodeRegex = /[\uFFF0-\uFFFF]/;
    
    return controlCharRegex.test(text) || problematicUnicodeRegex.test(text);
  }

  /**
   * Validate segment number for extraction
   */
  public static validateSegmentNumber(segmentNumber: any): ValidationResult<number> {
    const result: ValidationResult<number> = {
      isValid: false,
      errors: [],
      warnings: [],
      data: undefined
    };

    try {
      // Convert to number
      const num = Number(segmentNumber);

      // Check if it's a valid number
      if (isNaN(num)) {
        result.errors.push('Segment number is not a valid number');
        return result;
      }

      // Check if it's an integer
      if (!Number.isInteger(num)) {
        result.errors.push('Segment number must be an integer');
        return result;
      }

      // Check range
      if (num < 1 || num > 9) {
        result.errors.push('Segment number must be between 1 and 9');
        return result;
      }

      result.isValid = true;
      result.data = num;

    } catch (error) {
      result.errors.push(`Segment number validation error: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }

    return result;
  }

  /**
   * Batch validate multiple cell values
   */
  public static validateCellBatch(values: any[][]): ValidationResult<{
    validCells: number;
    totalCells: number;
    errors: Array<{ row: number; col: number; error: string }>;
    warnings: Array<{ row: number; col: number; warning: string }>;
  }> {
    const result: ValidationResult = {
      isValid: false,
      errors: [],
      warnings: [],
      data: {
        validCells: 0,
        totalCells: 0,
        errors: [],
        warnings: []
      }
    };

    try {
      let validCells = 0;
      let totalCells = 0;

      for (let row = 0; row < values.length; row++) {
        for (let col = 0; col < values[row].length; col++) {
          totalCells++;
          const cellValue = values[row][col];
          
          const cellValidation = this.validateCellValue(cellValue);
          
          if (cellValidation.isValid) {
            validCells++;
          } else {
            cellValidation.errors.forEach(error => {
              result.data!.errors.push({ row, col, error });
            });
          }

          cellValidation.warnings.forEach(warning => {
            result.data!.warnings.push({ row, col, warning });
          });
        }
      }

      result.data!.validCells = validCells;
      result.data!.totalCells = totalCells;
      result.isValid = validCells > 0; // At least one valid cell

      if (validCells === 0) {
        result.errors.push('No valid cells found in batch');
      } else if (validCells < totalCells) {
        result.warnings.push(`${totalCells - validCells} cells failed validation`);
      }

    } catch (error) {
      result.errors.push(`Batch validation error: ${error instanceof Error ? error.message : 'Unknown error'}`);
      errorHandler.handleError(errorHandler.createError(
        'Cell batch validation failed',
        'validation',
        'medium',
        'ValidationUtils.validateCellBatch',
        { values, error }
      ));
    }

    return result;
  }
}

export default ValidationUtils;