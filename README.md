# IPG Taxonomy Extractor v2.0.0

![TypeScript](https://img.shields.io/badge/TypeScript-5.3+-blue)
![Office.js](https://img.shields.io/badge/Office.js-1.12+-green)
![Node.js](https://img.shields.io/badge/Node.js-16+-brightgreen)
![Status](https://img.shields.io/badge/Status-Production-brightgreen)
![Cloudflare](https://img.shields.io/badge/Cloudflare-Workers-orange)

A production-ready Office Add-in for extracting segments from pipe-delimited taxonomy data across Excel Web, Desktop, and Mac platforms.

**🚀 Live Deployment:** [https://ipg-taxonomy-extractor-addin.wookstar.com](https://ipg-taxonomy-extractor-addin.wookstar.com)

## ✨ Features

### Core Functionality
- **9 Segment Extraction**: Extract any of the first 9 pipe-delimited segments
- **Activation ID Extraction**: Extract identifiers after colon characters  
- **Multi-Step Undo System**: LIFO stack supporting up to 10 operations
- **Acronym Pattern Processing**: Handle `^ABC^` patterns with Trim/Keep operations
- **Real-time Updates**: Instant UI updates when selecting different cells
- **Batch Processing**: Efficiently process multiple selected cells

### Technical Excellence
- **Cross-Platform**: Works on Excel Online, Windows, and Mac
- **Modern UI**: Fluent UI design system with responsive layout
- **Performance Optimized**: <50ms response times, handles 1000+ cells
- **Accessibility Ready**: WCAG 2.1 AA compliance foundation
- **Hot Reloading**: Development server with instant code updates

## 🚀 Quick Start

### Prerequisites
- Node.js 16.0.0 or higher
- npm 8.0.0 or higher
- Excel (Desktop, Online, or Mac)

### Development Setup
```bash
# Clone and install
git clone https://github.com/henkisdabro/taxonomy-extractor-officeaddin.git
cd taxonomy-extractor-officeaddin
npm install

# Start development server
npm run dev-server    # HTTP server at localhost:3001

# OR VS Code integrated debugging
npm start             # Press F5 in VS Code for auto-launch
```

### Using the Add-in
1. **Select cells** containing pipe-delimited taxonomy data
2. **Task pane updates** automatically showing data preview
3. **Extract segments** by clicking numbered buttons (1-9) or "Activation ID"
4. **Process patterns** when `^ABC^` patterns detected (Trim/Keep operations)
5. **Use "Undo Last"** to reverse operations

## 📊 Data Format

**Input Format:**
```
segment1|segment2|segment3|segment4|segment5|segment6|segment7|segment8|segment9:activationID
```

**Example:**
```
FY24_26|Q1-4|Tourism WA|WA|Always On Remarketing|4LAOSO|SOC|Facebook_Instagram|Conversions:DJTDOM060725
```

**Available Extractions:**
- **Segment 1**: `FY24_26`
- **Segment 3**: `Tourism WA`  
- **Segment 8**: `Facebook_Instagram`
- **Activation ID**: `DJTDOM060725`

## 🛠️ Developer Commands

```bash
# Development
npm run dev-server     # HTTP hot-reload server (recommended)
npm start              # VS Code integrated debugging with Excel launch
npm run validate       # Validate Office Add-in manifest

# Production
npm run build          # Full production build (add-in + worker)
npm run clean          # Clean build directory

# Deployment (Cloudflare Workers)
git push origin main   # Auto-deploys via GitHub integration
```

## 📁 Project Structure

```
taxonomy-extractor-officeaddin/
├── src/
│   ├── taskpane/
│   │   ├── taskpane.ts            # Core application logic
│   │   ├── taskpane.css           # Fluent UI styles
│   │   └── taskpane.html          # Main UI layout
│   ├── commands/
│   │   └── commands.ts            # Ribbon button handlers
│   ├── components/                # Component architecture (modernization)
│   ├── services/                  # State management and utilities
│   └── worker.ts                  # Cloudflare Workers handler
├── manifest.xml                   # Office Add-in configuration
├── webpack.config.js              # Production build configuration
├── webpack.dev.config.js          # Development server configuration
└── wrangler.toml                  # Cloudflare Workers deployment config
```

## 🔧 Development Features

### Hot Development Workflow
- **VS Code Integration**: Press F5 → Excel launches with add-in loaded
- **Hot Module Replacement**: Code changes reflect instantly without refresh
- **Development Simulation**: Test interface without Excel (localhost only)
- **Production-Safe**: Dev features automatically hidden in production

### Testing Capabilities
- **Simulate Taxonomy Data**: Test with sample pipe-delimited data
- **Simulate Targeting Patterns**: Test `^ABC^` pattern detection
- **Interface State Testing**: Clear selection, test different scenarios
- **Cross-Platform Testing**: Verify functionality across Excel platforms

## 🌐 Deployment

### Production (Current)
- **Platform**: Cloudflare Workers with Static Assets
- **URL**: `https://ipg-taxonomy-extractor-addin.wookstar.com`
- **Auto-Deploy**: GitHub → Cloudflare integration
- **Global CDN**: Worldwide distribution with edge caching

### Manual Installation
1. Download `manifest.xml` from this repository
2. In Excel: Home → Add-ins → More Add-ins → My Add-ins → Upload My Add-in
3. Select the `manifest.xml` file
4. Add-in loads from live Cloudflare deployment

## 🏗️ Architecture

### Core Technologies
- **TypeScript**: Type-safe development with strict mode
- **Office.js API**: Excel integration (ExcelApi 1.12)
- **Webpack**: Module bundling and hot-reload development
- **Cloudflare Workers**: Global edge deployment

### Key Components
```typescript
// Real-time selection monitoring
context.workbook.worksheets.onSelectionChanged.add(onSelectionChange);

// Multi-step undo system
class UndoManager {
  private static readonly MAX_OPERATIONS = 10;
  private operations: UndoOperation[] = [];
}

// Smart data detection
const hasPipes = cellText.includes('|');
const hasTargeting = /\^[^^]+\^ ?/.test(cellText);
```

## 📈 Performance

- **Bundle Size**: ~105KB (minified + gzipped)
- **Load Time**: <2 seconds first load
- **Response Time**: <50ms for selection updates
- **Memory Usage**: Efficient with automatic cleanup
- **Cross-Platform**: Consistent performance across Excel variants

## 🤝 Contributing

1. **Fork** the repository
2. **Create feature branch**: `git checkout -b feature/your-feature`
3. **Test thoroughly**: Verify across Excel platforms
4. **Run validation**: `npm run validate && npm run build`
5. **Submit Pull Request** with detailed description

### Development Standards
- TypeScript strict mode compliance
- Comprehensive error handling
- Performance optimization for large datasets
- Cross-platform compatibility testing

## 📄 License

MIT License - see [LICENSE](LICENSE) file for details.

## 🔗 Links

- [Live Add-in](https://ipg-taxonomy-extractor-addin.wookstar.com) - Production deployment
- [Issues](https://github.com/henkisdabro/taxonomy-extractor-officeaddin/issues) - Bug reports and feature requests
- [Original VBA Version](https://github.com/henkisdabro/excel-taxonomy-cleaner) - Legacy VBA implementation
- [Office Add-ins Documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/) - Microsoft's official documentation

---

**Ready to streamline taxonomy data extraction across all Excel platforms! 🚀**