# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview - Version 2.0.0

This project is a modern Office Add-in that ports the functionality of the original [VBA-based IPG Taxonomy Extractor](https://github.com/henkisdabro/excel-taxonomy-cleaner) to a cross-platform solution for Excel on the Web, Windows, and Mac.

The add-in provides a task pane interface for extracting specific segments from pipe-delimited taxonomy data. It features real-time updates based on user selection, a multi-step undo system, and advanced workflow enhancements.

## Architecture & Key Components

This is a web-based Office Add-in. The technology stack is:
-   **UI:** HTML5 and CSS, styled with Microsoft's Fluent UI.
-   **Logic:** TypeScript (compiled to JavaScript).
-   **API:** Office.js library for interacting with the Excel workbook.
-   **Configuration:** `manifest.xml` file defines the add-in's properties and ribbon integration.

### Planned File Structure (after scaffolding)
-   `manifest.xml`: Office Add-in manifest defining properties and ribbon integration
-   `src/taskpane/taskpane.html`: Task pane UI layout
-   `src/taskpane/taskpane.css`: UI styles (using Fluent UI components)
-   `src/taskpane/taskpane.ts`: Main application logic
-   `webpack.config.js`: Build configuration
-   `package.json`: Dependencies and scripts

## Current Status

**✅ COMPLETED: Production-ready Office Add-in fully implemented**

This project has been successfully ported from VBA to a modern Office Add-in with complete feature parity and enhanced functionality. The add-in is now ready for production deployment across Excel Desktop, Web, and Mac platforms.

## Getting Started

The project is fully implemented and ready for development/deployment:

1. **Install dependencies:**
   ```bash
   npm install
   ```

2. **Development commands:**
   ```bash
   npm start              # Launch development server with auto-reload
   npm run start:desktop  # Start specifically for Excel Desktop
   npm run start:web      # Start specifically for Excel Online
   npm run build          # Create production build
   npm run validate       # Validate manifest.xml
   npm run clean          # Clean dist folder
   ```

3. **Testing in Excel:**
   - Run `npm start` to launch the dev server
   - Open Excel (Desktop/Online/Mac)
   - Insert → Add-ins → Upload My Add-in → Select `manifest.xml`
   - Look for "IPG Taxonomy Extractor" in the Home tab

## Requirements
- Node.js and npm
- Excel (desktop, web, or Mac)
- Office Add-in development tools

## Implementation Architecture

### Core Technologies
- **Office.js API**: For Excel integration and workbook manipulation
- **TypeScript**: Type-safe development with interfaces for data structures
- **Fluent UI**: Microsoft's design system for consistent Office integration
- **Webpack**: Module bundling and development server

### Essential Data Structures (from VBA Analysis)

```typescript
interface UndoData {
  cellAddress: string;
  originalValue: string | number | boolean;
}

interface UndoOperation {
  description: string;           // "Extract Segment 3", "Extract Activation ID"
  cellChanges: UndoData[];       // Array of cell changes for this operation
  cellCount: number;             // Number of cells changed
  operationId: number;           // Unique identifier
  timestamp: Date;               // When operation was performed
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
}
```

### Critical VBA Functionality to Replicate

**Multi-Step Undo System (v1.6.0):**
- LIFO stack supporting up to 10 operations
- Dynamic button captions showing operation count: "Undo Last (3)"
- Captures state BEFORE making changes
- Processing feedback during undo execution
- Warning when stack reaches capacity

**Real-time Selection Handling:**
- Instant UI updates when user selects different cells
- Smart data validation (requires pipe characters)
- Dynamic button captions showing segment previews
- Cell count display for batch operations

**Smart Data Processing:**
- Pipe-delimited parsing: splits by `:` then `|`
- Handles missing segments gracefully
- Targeting acronym detection: `^ABC^` patterns
- Only processes cells containing taxonomy format

**Modeless Operation:**
- Task pane stays open during Excel interaction
- Real-time updates on selection changes
- Continuous workflow without reopening interface

### Core Implementation Patterns

**Office.js Integration:**
```typescript
Office.onReady(info => {
  // Initialize add-in and register selection change events
  Excel.run(async (context) => {
    context.workbook.worksheets.onSelectionChanged.add(onSelectionChange);
    await context.sync();
  });
});

// Real-time selection handler (replaces VBA App_SheetSelectionChange)
async function onSelectionChange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address, values, rowCount, columnCount");
    await context.sync();
    
    updateTaskPaneInterface(range);
  });
}
```

**Extraction Logic Pattern:**
```typescript
async function extractPipeSegment(segmentNumber: number) {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address, values");
    await context.sync();
    
    // Add to undo stack BEFORE making changes
    addUndoOperation(`Extract Segment ${segmentNumber}`, range);
    
    // Process each cell
    const newValues = range.values.map(row => 
      row.map(cell => extractSegmentFromCell(cell, segmentNumber))
    );
    
    range.values = newValues;
    await context.sync();
    
    // Update UI after changes
    refreshTaskPaneInterface();
  });
}
```

**Undo Stack Management:**
```typescript
class UndoManager {
  private static readonly MAX_OPERATIONS = 10;
  private operations: UndoOperation[] = [];
  
  addOperation(description: string, range: Excel.Range): void {
    // Capture current state before changes
    const operation: UndoOperation = {
      description,
      cellChanges: this.captureRangeState(range),
      cellCount: range.values.length,
      operationId: Date.now(),
      timestamp: new Date()
    };
    
    this.operations.push(operation);
    
    // Remove oldest if over limit
    if (this.operations.length > UndoManager.MAX_OPERATIONS) {
      this.operations.shift();
    }
    
    this.updateUndoButton();
  }
  
  async undoLastOperation(): Promise<void> {
    if (this.operations.length === 0) return;
    
    const lastOp = this.operations.pop()!;
    
    await Excel.run(async (context) => {
      // Restore each cell to its original value
      for (const change of lastOp.cellChanges) {
        const cell = context.workbook.worksheets.getActiveWorksheet()
          .getRange(change.cellAddress);
        cell.values = [[change.originalValue]];
      }
      await context.sync();
    });
    
    this.updateUndoButton();
  }
}
```

## Data Format Specifications

### IPG Interact Taxonomy Format
**Structure**: `segment1|segment2|segment3|segment4|segment5|segment6|segment7|segment8|segment9:activationID`

**Test Data Examples:**
```
FY24_26|Q1-4|Tourism WA|WA |Always On Remarketing| 4LAOSO | SOC|Facebook_Instagram|Conversions:DJTDOM060725
FY25|Q2|Brand Campaign|NSW|Video Content|ABC123|Social|YouTube|Awareness:TEST12345
```

**Edge Cases:**
- Missing segments: `FY24|Q1|||Campaign Name||SOC||:ID123`
- No activation ID: `FY24|Q1|Campaign|State|Type|Code|Channel|Platform|Goal`
- Targeting acronyms: `^AT^ testing string` (non-taxonomy data)

## Critical Implementation Requirements

### VBA Feature Parity Checklist

**Core Extraction Engine:**
- [ ] Parse pipe-delimited data: `segment1|segment2|...|segment9:activationID`
- [ ] Extract individual segments (1-9) with proper handling of missing segments
- [ ] Extract activation ID from text after colon
- [ ] Handle edge cases: no pipes, no colon, empty segments
- [ ] Batch processing for multiple selected cells

**Multi-Step Undo System (Essential):**
- [ ] LIFO stack with 10-operation limit
- [ ] Capture cell state BEFORE making changes
- [ ] Dynamic button text: "Undo Last" → "Undo Last (3)"
- [ ] Processing state feedback during undo
- [ ] Capacity warning at 10/10 operations
- [ ] Atomic undo operations per button click

**Real-Time Interface Updates:**
- [ ] Selection change event handling
- [ ] Dynamic button captions showing segment previews
- [ ] Smart button enabling/disabling based on data
- [ ] Cell count display for batch operations
- [ ] Instant UI refresh after extractions

**Smart Data Validation:**
- [ ] Require pipe characters for taxonomy processing
- [ ] Targeting acronym detection: `^ABC^` patterns
- [ ] Show "N/A" for unavailable segments
- [ ] Different UI states for taxonomy vs. targeting data

**Targeting Acronym Feature (v1.5.0):**
- [ ] Regex pattern matching: `\^[^^]+\^ ?`
- [ ] Smart button overlay when pattern detected
- [ ] One-click removal of targeting text
- [ ] Full undo integration

**Performance & UX:**
- [ ] Task pane remains open during Excel interaction
- [ ] No screen flickering during batch operations
- [ ] Focus management after button clicks
- [ ] Error resilience for individual cell failures

## ✅ Implementation Status - COMPLETED

All implementation phases have been successfully completed:

**✅ Phase 1: Foundation (MVP) - COMPLETED**
- ✅ Office Add-in project fully implemented
- ✅ Complete UI with 9 segment buttons + activation ID + targeting
- ✅ Advanced parsing logic for pipe-delimited data
- ✅ All segment extractions (1-9) working
- ✅ Real-time selection change event handling

**✅ Phase 2: Core Functionality - COMPLETED**
- ✅ Multi-step undo system (10 operations LIFO stack)
- ✅ All segment extraction buttons (1-9) with dynamic previews
- ✅ Activation ID extraction with colon parsing
- ✅ Real-time UI updates and button previews
- ✅ Efficient batch processing for multiple cells

**✅ Phase 3: Advanced Features - COMPLETED**
- ✅ Targeting acronym detection and removal (^ABC^ patterns)
- ✅ Smart button states and comprehensive validation
- ✅ Performance optimizations with monitoring
- ✅ Comprehensive error handling and edge cases
- ✅ Professional Fluent UI integration

**✅ Phase 4: Enhanced & Deployment - COMPLETED**
- ✅ Enhanced features beyond original VBA version
- ✅ Production-ready manifest configuration
- ✅ Cross-platform compatibility (Desktop/Online/Mac)
- ✅ Comprehensive documentation and user guides

## 🚀 Enhanced Features Beyond VBA

The Office Add-in version includes several improvements over the original VBA:

**Modern Architecture:**
- TypeScript with strict type checking
- Comprehensive logging system with different log levels
- Performance monitoring for operations
- Enhanced error handling with user-friendly messages

**Cross-Platform Support:**
- Works on Excel Desktop (Windows/Mac)
- Compatible with Excel Online (web)
- Responsive design that adapts to different screen sizes

**Enhanced UX:**
- Professional Fluent UI design system
- Real-time status messages with auto-hide
- Dynamic button states and previews
- Improved accessibility features

**Developer Experience:**
- Hot reload development server
- Automatic SSL certificate generation
- Webpack-based build system
- Comprehensive validation tools

## 🛠️ Developer Instructions for Contributors

### Understanding the Codebase

**Core Architecture:**
- `src/taskpane/taskpane.ts`: Main application logic with TaxonomyExtractor class
- `src/taskpane/taskpane.html`: UI layout with Fluent UI components
- `src/taskpane/taskpane.css`: Styling with Microsoft Fluent UI framework
- `manifest.xml`: Office Add-in manifest defining permissions and integration points

**Key Classes:**
- `TaxonomyExtractor`: Main application class handling all functionality
- `Logger`: Enhanced logging system with multiple levels (ERROR, WARN, INFO, DEBUG)
- `PerformanceMonitor`: Tracks operation performance for optimization

**Data Structures:**
- `UndoOperation`: Captures state for multi-step undo functionality
- `ParsedCellData`: Structured representation of taxonomy data parsing results

### Development Best Practices

**When making changes:**
1. **Always test across platforms**: Desktop Excel, Excel Online, and Mac
2. **Use proper TypeScript**: Leverage strict type checking and interfaces
3. **Handle errors gracefully**: Wrap Office.js calls in try-catch blocks
4. **Performance matters**: Use PerformanceMonitor for timing critical operations
5. **Respect undo system**: Always call `addUndoOperation` before modifying cells

**Critical Implementation Details:**
- Selection changes trigger real-time UI updates via `onSelectionChange`
- Undo operations must capture state BEFORE making changes (VBA pattern)
- Button states dynamically update based on parsed data availability
- Targeting acronym detection takes precedence over taxonomy parsing

**Testing Checklist:**
- [ ] Real-time selection updates work smoothly
- [ ] All 9 segment extractions function correctly
- [ ] Activation ID extraction handles various formats
- [ ] Targeting acronym removal works with ^ABC^ patterns
- [ ] Undo system maintains proper operation order (LIFO)
- [ ] Batch processing handles large cell ranges efficiently
- [ ] Error messages are user-friendly and actionable
- [ ] Task pane remains responsive during operations

**Common Development Commands:**
```bash
npm start                    # Development with hot reload
npm run validate            # Check manifest.xml compliance
npm run build               # Production build for deployment
npm run clean && npm start  # Fresh development environment
```

**Debugging in Different Environments:**
- **Excel Desktop**: Use Visual Studio Code debugger with breakpoints
- **Excel Online**: Browser dev tools (F12) for console logging and network inspection
- **Mac Excel**: Safari/Chrome dev tools when using web view

### Office.js Integration Patterns

**Selection Event Handling:**
```typescript
// Register for selection changes (modeless behavior)
context.workbook.worksheets.onSelectionChanged.add(async () => {
  await taxonomyExtractor.onSelectionChange();
});
```

**Safe Cell Operations:**
```typescript
// Always wrap in Excel.run for proper context management
await Excel.run(async (context) => {
  const range = context.workbook.getSelectedRange();
  range.load('address, values');
  await context.sync();
  // Process range data...
});
```

**Error Handling Pattern:**
```typescript
try {
  await Excel.run(async (context) => {
    // Office.js operations
  });
  this.showStatus('Success message');
} catch (error) {
  Logger.error('Operation failed', error);
  this.showStatus('User-friendly error message', true);
}
```