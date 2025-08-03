# Development Guide - IPG Taxonomy Extractor v2.0.0

## 🔧 Advanced Development Setup

### IDE Configuration

**VS Code Extensions (Recommended):**
```json
{
  "recommendations": [
    "ms-vscode.vscode-typescript-next",
    "ms-vscode.vscode-json",
    "esbenp.prettier-vscode",
    "ms-office-addin.office-addin-debugger",
    "bradlc.vscode-tailwindcss"
  ]
}
```

**TypeScript Configuration:**
- Strict mode enabled for type safety
- ES2017 target for modern browser support
- Source maps enabled for debugging

### SSL Certificate Management

Development certificates are automatically managed:
```bash
# Verify certificates
npx office-addin-dev-certs verify

# Reinstall if needed
npx office-addin-dev-certs install

# Check certificate location
ls ~/.office-addin-dev-certs/
```

## 🏗️ Build System

### Webpack Configuration

The build system uses webpack with:
- **TypeScript compilation** with ts-loader
- **Hot module replacement** for fast development
- **Source maps** for debugging
- **Asset optimization** for production

### Build Modes

```bash
# Development (with source maps, no minification)
npm run build:dev

# Production (minified, optimized)
npm run build:prod

# Watch mode (rebuilds on changes)
npm run watch
```

## 🧪 Testing Strategy

### Manual Testing Checklist

**Core Functionality:**
- [ ] Selection change triggers UI update
- [ ] All 9 segment buttons extract correctly
- [ ] Activation ID extraction works
- [ ] Undo system maintains 10 operations
- [ ] Targeting acronym detection and removal
- [ ] Batch processing multiple cells

**UI/UX Testing:**
- [ ] Buttons enable/disable based on data
- [ ] Dynamic captions show segment previews
- [ ] Status messages appear for operations
- [ ] Responsive design in narrow task panes
- [ ] Dark mode adaptation

**Error Scenarios:**
- [ ] Empty cell selection
- [ ] Non-taxonomy data (no pipes)
- [ ] Malformed taxonomy data
- [ ] Network connectivity issues
- [ ] Excel API failures

### Test Data Sets

```typescript
// Valid taxonomy data
const testData = [
  "FY24_26|Q1-4|Tourism WA|WA |Always On Remarketing| 4LAOSO | SOC|Facebook_Instagram|Conversions:DJTDOM060725",
  "FY25|Q2|Brand Campaign|NSW|Video Content|ABC123|Social|YouTube|Awareness:TEST12345"
];

// Edge cases
const edgeCases = [
  "FY24|Q1|||Campaign Name||SOC||:ID123",  // Missing segments
  "FY24|Q1|Campaign|State|Type|Code|Channel|Platform|Goal",  // No activation ID
  "^AT^ testing string",  // Targeting acronym
  "",  // Empty string
  "No pipes here"  // No taxonomy format
];
```

## 📊 Performance Optimization

### Key Metrics

Monitor these performance indicators:
- **Selection Change Response**: < 50ms
- **Extraction Operations**: < 100ms for 100 cells
- **UI Update Latency**: < 30ms
- **Memory Usage**: < 50MB for normal operations

### Optimization Techniques

**Debouncing Selection Changes:**
```typescript
let selectionTimeout: number;
function onSelectionChange() {
  clearTimeout(selectionTimeout);
  selectionTimeout = setTimeout(() => {
    updateInterface();
  }, 50);
}
```

**Batch Processing:**
```typescript
// Process cells in batches to avoid blocking UI
const BATCH_SIZE = 100;
for (let i = 0; i < totalCells; i += BATCH_SIZE) {
  const batch = cells.slice(i, i + BATCH_SIZE);
  await processBatch(batch);
  await new Promise(resolve => setTimeout(resolve, 0)); // Yield to UI
}
```

## 🐛 Debugging

### Office Add-in Debugger

```bash
# Start debugging session
npm start

# Debug in different environments
npm run start:desktop  # Excel Desktop
npm run start:web      # Excel Online
```

### Console Logging

The application includes comprehensive logging:
```typescript
Logger.setLevel('DEBUG');  // Enable verbose logging
Logger.info('Operation completed', { cellsProcessed: 50 });
Logger.error('Failed to extract segment', error);
```

### Common Issues

**SSL Certificate Issues:**
```bash
# Regenerate certificates
npx office-addin-dev-certs uninstall
npx office-addin-dev-certs install
```

**Manifest Validation Errors:**
```bash
# Validate manifest
npm run validate

# Check for common issues:
# - Invalid URLs
# - Missing resource strings
# - Incorrect XML structure
```

## 🔒 Security Considerations

### Content Security Policy

The add-in implements strict CSP:
```html
<meta http-equiv="Content-Security-Policy" 
      content="default-src 'self' https://appsforoffice.microsoft.com; 
               script-src 'self' 'unsafe-inline' https://appsforoffice.microsoft.com;
               style-src 'self' 'unsafe-inline' https://static2.sharepointonline.com;">
```

### Data Protection

- **No data persistence**: All operations are in-memory only
- **Local processing**: No data sent to external servers
- **Secure communication**: HTTPS required for all endpoints

## 📱 Cross-Platform Compatibility

### Platform-Specific Testing

**Excel Desktop (Windows/Mac):**
- Full Office.js API support
- Native performance
- Complete feature set

**Excel Online:**
- Subset of Office.js APIs
- Network-dependent performance
- Some limitations in file access

**Excel Mobile:**
- Limited API surface
- Touch-optimized UI needed
- Reduced functionality

### API Compatibility Matrix

| Feature | Desktop | Online | Mobile |
|---------|---------|--------|--------|
| Selection Changes | ✅ | ✅ | ⚠️ |
| Range Manipulation | ✅ | ✅ | ✅ |
| Event Handlers | ✅ | ✅ | ⚠️ |
| Custom Functions | ✅ | ✅ | ❌ |

## 🚀 Deployment Pipeline

### Development → Staging → Production

1. **Development**
   ```bash
   npm run build:dev
   # Test locally with npm start
   ```

2. **Staging**
   ```bash
   npm run build:prod
   # Deploy to staging server
   # Update manifest URLs to staging
   # Test in target Office environments
   ```

3. **Production**
   ```bash
   npm run build:prod
   # Deploy to production CDN
   # Update manifest for production
   # Submit to Office Store (optional)
   ```

### Continuous Integration

Example GitHub Actions workflow:
```yaml
name: CI/CD
on: [push, pull_request]
jobs:
  test:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v3
      - uses: actions/setup-node@v3
        with:
          node-version: '18'
      - run: npm ci
      - run: npm run build:prod
      - run: npm run validate
```

## 📚 Additional Resources

### Microsoft Documentation
- [Office Add-ins Platform Overview](https://learn.microsoft.com/en-us/office/dev/add-ins/)
- [Excel JavaScript API Reference](https://learn.microsoft.com/en-us/javascript/api/excel)
- [Office Add-ins Best Practices](https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/add-in-development-best-practices)

### Community Resources
- [Office Dev Center](https://developer.microsoft.com/en-us/office)
- [Office Add-ins Community](https://github.com/OfficeDev)
- [Stack Overflow - office-js](https://stackoverflow.com/questions/tagged/office-js)

### Tools and Utilities
- [Script Lab](https://script-lab.azureedge.net/) - Office Add-in playground
- [Office Add-in Validator](https://ux.microsoft.com/addins/validation) - Online validation tool
- [Fluent UI](https://developer.microsoft.com/en-us/fluentui) - Microsoft's design system

---

This development guide ensures professional-grade development practices and helps maintain the high quality standards of the IPG Taxonomy Extractor Office Add-in.