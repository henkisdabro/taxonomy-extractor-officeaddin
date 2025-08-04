# IPG Taxonomy Extractor v2.0.0 - Office Add-in

![TypeScript](https://img.shields.io/badge/TypeScript-5.3+-blue)
![Office.js](https://img.shields.io/badge/Office.js-1.12+-green)
![Node.js](https://img.shields.io/badge/Node.js-16+-brightgreen)
![License](https://img.shields.io/badge/License-MIT-yellow)
![Status](https://img.shields.io/badge/Status-Deployed-brightgreen)
![Cloudflare](https://img.shields.io/badge/Cloudflare-Workers-orange)

A modern, production-ready Office Add-in that ports the functionality of the [VBA-based IPG Taxonomy Extractor](https://github.com/henkisdabro/excel-taxonomy-cleaner) to work seamlessly across Excel Online, Windows, and Mac platforms. Currently deployed on **Cloudflare Workers** with enhanced features and 2024/2025 compliance standards.

**🚀 Live Deployment:** [https://ipg-taxonomy-extractor-addin.wookstar.com](https://ipg-taxonomy-extractor-addin.wookstar.com)

## 🌟 Features

### ✅ **Complete VBA Feature Parity + Enhanced Features**
- **9 Segment Extraction**: Extract any of the first 9 pipe-delimited segments
- **Activation ID Extraction**: Extract unique identifiers after colon characters
- **Multi-Step Undo System**: LIFO stack supporting up to 10 operations with dynamic button captions
- **Real-time Updates**: Instant UI updates when selecting different cells
- **Enhanced Acronym Pattern Processing**: Smart detection of `^ABC^` patterns with dual functionality:
  - **Trim**: Remove `^ABC^` patterns while keeping surrounding text
  - **Keep**: Extract only `^ABC^` patterns while removing everything else
- **Batch Processing**: Process multiple selected cells efficiently

### 🚀 **Modern Enhancements**
- **Cross-Platform Compatibility**: Works on Excel Online, Windows, and Mac
- **Professional Fluent UI**: Native Microsoft design system integration
- **Enhanced Error Handling**: Comprehensive logging and error recovery
- **Performance Monitoring**: Built-in performance tracking and optimization
- **Responsive Design**: Adapts to different task pane sizes
- **Snappy UI Animations**: Ultra-fast transitions (30ms) for responsive feel
- **Smart Interface States**: Context-aware UI that adapts based on data patterns

## 🛠️ Development Setup

### Prerequisites
- **Node.js** 16.0.0 or higher
- **npm** 8.0.0 or higher
- **Excel** (Desktop, Online, or Mac)

### Quick Start

1. **Clone the repository**
   ```bash
   git clone https://github.com/henkisdabro/taxonomy-extractor-officeaddin.git
   cd taxonomy-extractor-officeaddin
   ```

2. **Install dependencies**
   ```bash
   npm install
   ```

3. **Start development server**
   ```bash
   # Primary development server (hot-reloading, easy setup)
   npm run dev-server
   
   # Alternative HTTPS server (if certificates needed)
   npm start
   ```
   
   **Development server features:**
   - **Hot reloading**: Instant CSS/TypeScript updates without refresh
   - **HTTPS server**: `https://localhost:3000/taskpane.html` with auto-generated certificates
   - **Development simulation**: Yellow dev section with test buttons (localhost only)
   - **Interface testing**: Simulate taxonomy data, targeting patterns, and selection states
   - **Production-safe**: Dev features automatically hidden in production environments

4. **Sideload in Excel**
   - Open Excel For Web (usually not possible on desktop)
   - Go to Home → Add-ins → More Add-ins → My Add-ins → Upload My Add-in
   - Select the `manifest.xml` file from the project root
   - The add-in will appear in the "IPG Tools" group on the Home tab

## 📁 Project Structure

```
taxonomy-extractor-officeaddin/
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html          # Main UI layout
│   │   ├── taskpane.css           # Fluent UI styles
│   │   └── taskpane.ts            # Core application logic
│   └── commands/
│       ├── commands.html          # Command functions
│       └── commands.ts            # Ribbon button handlers
├── manifest.xml                   # Office Add-in manifest
├── webpack.config.js              # Build configuration
├── package.json                   # Dependencies and scripts
└── README.md                      # This file
```

## 🎯 Usage

### Data Format
The add-in processes IPG Interact Taxonomy format:
```
segment1|segment2|segment3|segment4|segment5|segment6|segment7|segment8|segment9:activationID
```

### Example
```
FY24_26|Q1-4|Tourism WA|WA |Always On Remarketing| 4LAOSO | SOC|Facebook_Instagram|Conversions:DJTDOM060725
```

### Extraction Results
- **Segment 1**: `FY24_26`
- **Segment 3**: `Tourism WA`
- **Segment 5**: `Always On Remarketing`
- **Segment 8**: `Facebook_Instagram`
- **Segment 9**: `Conversions`
- **Activation ID**: `DJTDOM060725`

### Workflow
1. **Select cells** containing pipe-delimited taxonomy data
2. **Task pane updates** automatically showing segment previews
3. **Extract segments** by clicking numbered buttons (1-9) or "Activation ID"
4. **Process acronym patterns** when `^ABC^` patterns are detected:
   - Interface automatically switches to "Acronym Pattern" mode
   - **Trim**: Remove `^ABC^` patterns, keep surrounding text
   - **Keep**: Extract only `^ABC^` patterns, remove everything else
5. **Use "Undo Last"** to reverse operations (up to 10 steps)
6. **Continue processing** different ranges without closing the task pane

## 🔧 Development Commands

```bash
# Development
npm run dev-server     # Primary dev server with hot-reloading (HTTPS, port 3000)
npm start              # Alternative dev server with auto-reload
npm run start:desktop  # Start for Excel Desktop
npm run start:web      # Start for Excel Online

# Building
npm run build:dev      # Development build
npm run build          # Production build
npm run build:prod     # Production build with analysis

# Validation
npm run validate       # Validate manifest.xml
npm run lint           # Run linting (when configured)
npm run test           # Run tests (when configured)

# Utilities
npm run clean          # Clean dist folder
npm run stop           # Stop debugging session
```

### 🧪 Development Testing Features

The development server includes powerful testing capabilities:

**Dev Simulation Buttons** (visible only on localhost):
- **Simulate Taxonomy Data**: Test interface with sample pipe-delimited data
- **Simulate Targeting Pattern**: Test `^ABC^` pattern detection and processing
- **Clear Selection**: Reset interface to initial state

**Hot Development Workflow**:
1. Run `npm run dev-server`
2. Open `https://localhost:3000/taskpane.html` in browser
3. Use yellow dev section to test different interface states
4. Make code changes - see updates instantly without refresh
5. Test all extraction and targeting functionality without Excel

## 📊 Architecture

### Core Technologies
- **TypeScript**: Type-safe development with strict mode
- **Office.js API**: Excel integration and workbook manipulation
- **Webpack**: Module bundling and development server
- **Fluent UI**: Microsoft's design system

### Key Components

**Real-time Selection Handler**
```typescript
// Monitors Excel selection changes
context.workbook.worksheets.onSelectionChanged.add(onSelectionChange);
```

**Multi-Step Undo System**
```typescript
// LIFO stack with operation tracking
class UndoManager {
  private static readonly MAX_OPERATIONS = 10;
  private operations: UndoOperation[] = [];
}
```

**Data Validation Engine**
```typescript
// Smart taxonomy format detection
const hasPipes = cellText.includes('|');
const hasTargeting = /\^[^^]+\^ ?/.test(cellText);
```

## 🚀 Deployment

### 🌐 Production Deployment (Current)

**Status: LIVE ON CLOUDFLARE WORKERS** ✅

The Office Add-in is currently deployed and fully functional at:
- **Main URL**: `https://ipg-taxonomy-extractor-addin.wookstar.com`
- **Task Pane**: `https://ipg-taxonomy-extractor-addin.wookstar.com/taskpane`
- **Commands**: `https://ipg-taxonomy-extractor-addin.wookstar.com/commands`

**Deployment Architecture:**
- **Platform**: Cloudflare Workers with Static Assets
- **Auto-Deploy**: GitHub → Cloudflare integration
- **Build Time**: ~12 seconds
- **Global CDN**: Automatic worldwide distribution

**To use the deployed add-in:**
1. Download the production `manifest.xml` from the repository
2. In Excel For Web: Home → Add-ins → More Add-ins → My Add-ins → Upload My Add-in → Select `manifest.xml`
3. The add-in will load from the live Cloudflare deployment

**Future Deployment Plans:**
- **Organization Add-ins Repository**: Plans are underway to deploy this add-in through the organization's centralized add-ins repository for automatic installation across all user accounts, eliminating the need for manual sideloading

### Development Testing
1. **Install dependencies**: `npm install`
2. **Start development server**: `npm start`
   - Generates SSL certificates automatically
   - Opens browser with dev server URL
   - Enables hot reload for development
3. **Sideload in Excel**:
   - **Excel For Web**: Home → Add-ins → More Add-ins → My Add-ins → Upload My Add-in → Select `manifest.xml`
   - **Desktop Excel**: Usually not possible for custom add-ins
   - **Mac Excel**: Insert → Add-ins → My Add-ins → Upload My Add-in → Select `manifest.xml`
4. **Test with sample data**:
   ```
   FY24_26|Q1-4|Tourism WA|WA |Always On Remarketing| 4LAOSO | SOC|Facebook_Instagram|Conversions:DJTDOM060725
   FY25|Q2|Brand Campaign|NSW|Video Content|ABC123|Social|YouTube|Awareness:TEST12345
   ^AT^ testing string
   ```
5. **Verify functionality**:
   - Real-time selection updates
   - All 9 segment extractions
   - Activation ID extraction
   - Targeting acronym removal
   - Multi-step undo system

### Manual Production Deployment (Alternative)
If you want to deploy to your own infrastructure:
1. **Build for production**: `npm run build`
2. **Host on secure server**: Upload `dist` folder to HTTPS-enabled hosting
3. **Update manifest**: Change URLs in `manifest.xml` to your domain
4. **Deploy via Office 365 Admin Center** (Centralized Deployment) or distribute manifest file
5. **Validate deployment**: Use `npm run validate` to check manifest compliance

### Cloudflare Workers Deployment (Current Setup)
The project uses modern **Cloudflare Workers Static Assets** for optimal performance:
```toml
[assets]
directory = "./dist"
binding = "ASSETS"
```

**Deployment Process:**
1. Push changes to GitHub `master` branch
2. Cloudflare automatically detects changes
3. Runs `npm run build` in cloud environment
4. Deploys worker with static assets globally
5. Updates are live within 1-2 minutes

## 📈 Performance

- **Real-time Updates**: < 50ms response time for selection changes
- **Batch Processing**: Handles 1000+ cells efficiently
- **Memory Management**: Automatic cleanup of undo operations
- **Optimized Bundle**: Minified production build < 500KB

## 🔒 Security

- **HTTPS Required**: All communications use SSL/TLS
- **Trusted Certificates**: Development certificates for localhost
- **Content Security Policy**: Prevents XSS attacks
- **Manifest Validation**: Strict XML schema compliance

## 📝 Contributing

We welcome contributions! Please follow these guidelines:

### Development Workflow
1. **Fork** the repository
2. **Clone** your fork: `git clone https://github.com/YOUR_USERNAME/taxonomy-extractor-officeaddin.git`
3. **Install dependencies**: `npm install`
4. **Create feature branch**: `git checkout -b feature/your-feature-name`
5. **Make changes** and test thoroughly in Excel
6. **Run validation**: `npm run validate` and `npm run build`
7. **Commit** with descriptive message: `git commit -m 'Add: new segment validation feature'`
8. **Push** to your fork: `git push origin feature/your-feature-name`
9. **Create Pull Request** with detailed description

### Code Standards
- **TypeScript**: Use strict type checking
- **Office.js**: Follow Microsoft's best practices
- **Error handling**: Comprehensive try-catch blocks
- **Performance**: Optimize for large data sets
- **Testing**: Test across Excel Desktop, Online, and Mac

### Debugging Tips
- Use browser dev tools (F12) in Excel Online
- Enable Office.js logging: `Office.context.requirements.isSetSupported('Logging', '1.1')`
- Test with various data formats and edge cases
- Verify undo functionality works correctly

## 📄 License

MIT License - see LICENSE file for details.

## 🔗 Related Projects

- [VBA IPG Taxonomy Extractor](https://github.com/henkisdabro/excel-taxonomy-cleaner) - Original VBA version
- [Office Add-ins Documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/) - Microsoft's official docs

## 🆘 Support

For issues, questions, or feature requests:
1. Check the [Issues](https://github.com/henkisdabro/taxonomy-extractor-officeaddin/issues) page
2. Create a new issue with detailed information
3. Include your Office version and platform details

---

**Ready to streamline your taxonomy data extraction across all Excel platforms! 🚀**