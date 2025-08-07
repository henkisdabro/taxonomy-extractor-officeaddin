// @ts-check
const eslint = require('@eslint/js');
const tseslint = require('typescript-eslint');
const officeAddins = require('eslint-plugin-office-addins');
const prettier = require('eslint-plugin-prettier');

module.exports = tseslint.config(
  // Base ESLint recommended rules
  eslint.configs.recommended,
  
  // TypeScript-ESLint configurations (using stable recommended set)
  ...tseslint.configs.recommended,
  
  // Office Add-ins specific configuration
  ...officeAddins.configs.recommended,
  
  // Custom configuration for Office Add-ins development
  {
    plugins: {
      'office-addins': officeAddins,
      prettier: prettier
    },
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: 'module',
      globals: {
        // Office.js globals
        Office: 'readonly',
        Excel: 'readonly',
        Word: 'readonly',
        PowerPoint: 'readonly',
        OneNote: 'readonly',
        Outlook: 'readonly',
        
        // Browser globals
        console: 'readonly',
        window: 'readonly',
        document: 'readonly',
        
        // Node.js globals (for build scripts)
        require: 'readonly',
        module: 'readonly',
        exports: 'readonly',
        __dirname: 'readonly',
        __filename: 'readonly',
        global: 'readonly',
        process: 'readonly',
        Buffer: 'readonly'
      }
    },
    rules: {
      // Office Add-ins specific adjustments
      'no-console': 'off', // Console logging is common in Office Add-ins for debugging
      'no-debugger': 'error',
      
      // TypeScript rules tailored for Office Add-ins
      '@typescript-eslint/no-explicit-any': 'warn', // Office.js often requires any types
      '@typescript-eslint/explicit-function-return-type': 'off',
      '@typescript-eslint/explicit-module-boundary-types': 'off',
      '@typescript-eslint/no-inferrable-types': 'off',
      '@typescript-eslint/no-unused-vars': ['error', { 
        argsIgnorePattern: '^_', 
        varsIgnorePattern: '^_',
        ignoreRestSiblings: true
      }],
      
      // Code quality rules
      'eqeqeq': ['error', 'always'],
      'curly': ['error', 'all'],
      'prefer-const': 'error',
      'no-var': 'error',
      'no-unused-vars': ['error', { 
        argsIgnorePattern: '^_', 
        varsIgnorePattern: '^_',
        ignoreRestSiblings: true
      }],
      
      // Prettier integration
      'prettier/prettier': 'error',
      
      // Office.js specific overrides
      'no-undef': 'off' // TypeScript handles undefined variables
    }
  },
  
  // Specific overrides for different file types
  {
    files: ['**/*.js'],
    rules: {
      '@typescript-eslint/no-var-requires': 'off'
    }
  },
  
  // Configuration files exceptions
  {
    files: ['*.config.js', '*.config.mjs', 'webpack.config.js'],
    rules: {
      '@typescript-eslint/no-var-requires': 'off',
      'no-console': 'off'
    }
  },
  
  // Commands file exception (functions referenced in manifest.xml)
  {
    files: ['**/commands/*.ts'],
    rules: {
      'no-unused-vars': 'off',
      '@typescript-eslint/no-unused-vars': 'off'
    }
  },
  
  // Ignore patterns
  {
    ignores: [
      'dist/**/*',
      'node_modules/**/*',
      '*.min.js',
      'assets/**/*'
    ]
  }
);