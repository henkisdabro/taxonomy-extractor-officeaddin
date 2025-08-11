## 🔄 IMPORTANT: File Synchronization Requirements

**CRITICAL:** This file must be kept identical to `GEMINI.md` at all times.

### **Synchronization Rules:**
1. **Any updates to this file** must be immediately copied to `GEMINI.md`
2. **Any updates from other AI applications** to `GEMINI.md` must be copied here
3. **Both files must contain identical content** for consistency across AI assistants
4. **Before making changes:** Always check both files are synchronized
5. **After making changes:** Always update both files simultaneously

**Purpose:** Ensures all AI assistants (Claude, Gemini, etc.) have access to the same project knowledge and context.

---

## 🎯 IPG Taxonomy Extractor - Current Project Status (January 2025)

**Status: PRODUCTION READY - PHASE 3 MODERNIZATION IN PROGRESS** ✅

### **Production Deployment**
- **Live URL**: `https://ipg-taxonomy-extractor-addin.wookstar.com`
- **Platform**: Cloudflare Workers with Static Assets API
- **Status**: Fully functional across Excel Web, Desktop, and Mac
- **Auto-deployment**: GitHub → Cloudflare integration active

### **Core Functionality (Stable)**
- **9 Segment Extraction**: All pipe-delimited segments working
- **Activation ID Extraction**: Post-colon identifier extraction
- **Targeting Pattern Processing**: `^ABC^` pattern detection with Trim/Keep operations
- **Multi-Step Undo System**: LIFO stack supporting 10 operations
- **Real-time Interface**: Context-aware UI with selection monitoring

---

## 📋 Excel API & Microsoft Store Compliance

### **Excel API Configuration**
```xml
<Requirements>
  <Sets>
    <Set Name="ExcelApi" MinVersion="1.12"/>
  </Sets>
</Requirements>
```

**Rationale for ExcelApi 1.12:**
- **Broad compatibility** across all Office platforms (Web, Desktop, Mac, Mobile)
- **Stable foundation** for all current taxonomy extraction requirements
- **Production tested** and reliable for enterprise deployment
- **Future upgrade path** available when advanced features needed

### **Office Add-in Manifest Requirements**
```xml
<!-- CRITICAL: Four-part version format required -->
<Version>2.0.0.0</Version>
<!-- NOT: v2.0.0 (will fail Microsoft validation) -->
```

**Validation Command:**
```bash
npm run validate  # Must show "The manifest is valid."
```

---

## 🚀 Development Environment & Workflow

### **Hot-Reloading Development Server**
```bash
# Primary development workflow
npm run dev-server     # HTTP server at localhost:3001
npm start              # VS Code integrated debugging
```

**Development Features:**
- **HTTP-only server**: Avoids certificate complexity
- **Hot module replacement**: Instant code updates
- **VS Code integration**: F5 launches Excel with add-in loaded
- **Development simulation**: Yellow dev section for interface testing
- **Production-safe**: Dev features hidden outside localhost

### **Build & Deployment Pipeline**
```bash
# Quality assurance
npm run validate       # Manifest validation
npm run build         # Production build (add-in + worker)

# Git workflow
git add .
git commit -m "Descriptive commit message"
git push origin feature/branch-name
```

### **VS Code Debugging Configuration**
- **Port**: 3001 (matches webpack.dev.config.js)
- **Task**: "Debug: Excel Desktop" (F5 shortcut)
- **Auto-launch**: Excel opens with add-in pre-loaded
- **Hot-reload**: Code changes reflect immediately

---

## 🏗️ Modernization Roadmap Progress

### **✅ COMPLETED PHASES**

#### **Phase 1: Asset Foundation** ✅
- Icon variants (16px, 32px, 80px) properly configured
- Webpack asset pipeline validated
- Cloudflare Workers asset deployment confirmed

#### **Phase 2: Internationalization Setup** ✅
- Localization framework (`LocalizationService`) implemented
- Australian English (`en-AU`) as primary locale
- 30+ UI strings externalized to JSON
- Parameter interpolation support

#### **Phase 3: Component Architecture Foundation** ✅
- `BaseComponent` class with event handling and state management
- `StateManager` service for centralized state with change detection
- Comprehensive type system in `src/types/taxonomy.types.ts`
- Component migration strategy established

#### **Phase 4: Accessibility & Localization** ✅
- WCAG 2.1 AA compliance foundation
- `AccessibilityService` with screen reader support
- High contrast mode detection and adaptation
- Focus management and keyboard navigation enhancement

### **🚧 IN PROGRESS**

#### **Phase 3: Component Migration** (Current)
- **UndoSystem** component migration - NEXT
- **ActivationManager** component migration - PENDING
- **TargetingProcessor** component migration - PENDING
- **SegmentExtractor** component migration - PENDING

### **📋 UPCOMING PHASES**

#### **Phase 5: Code Quality & Error Handling** ⏳
- Comprehensive error boundaries and validation
- Testing framework implementation (Jest/Vitest)
- Type safety enforcement

#### **Phase 6: Performance & Optimization** ⏳
- Bundle analysis and tree shaking
- Modern browser targeting
- Code splitting strategy

---

## 🔧 Technical Architecture

### **Cloudflare Workers Configuration**
```toml
# wrangler.toml - Modern Static Assets API
[assets]
directory = "./dist"
binding = "ASSETS"
```

```javascript
// worker.ts - Static asset serving
interface Env {
  ASSETS: Fetcher;
}
export default {
  async fetch(request: Request, env: Env): Promise<Response> {
    return await env.ASSETS.fetch(request);
  }
}
```

### **Development Build Configuration**
```javascript
// webpack.dev.config.js - HTTP development server
module.exports = {
  devServer: {
    hot: true,
    server: "http",    // No HTTPS certificates needed
    port: 3001,        // Consistent with package.json config
    historyApiFallback: { index: "/taskpane.html" }
  }
}
```

### **Component Architecture Pattern**
```typescript
// BaseComponent foundation
export abstract class BaseComponent {
  protected localization: LocalizationService;
  protected stateManager: StateManager;
  protected accessibilityService: AccessibilityService;
  
  abstract initialize(): Promise<void>;
  abstract cleanup(): void;
}
```

---

## ⚡ Performance & Quality Metrics

### **Build Performance**
- **Total build time**: ~71 seconds (optimized)
- **Dependency installation**: ~30 seconds
- **Worker deployment**: ~5 seconds
- **Bundle size**: Add-in (~105KB) + Worker (388 bytes)

### **Runtime Performance**
- **Selection updates**: <50ms response time
- **UI transitions**: 30ms (hardware accelerated)
- **Batch processing**: 1000+ cells efficiently handled
- **Memory management**: Automatic cleanup with 10-operation undo limit

---

## 🎨 UI/UX Design Standards

### **Current Interface Specifications**
- **Typography**: 12-13px fonts for space efficiency
- **Button dimensions**: 32px height, 8px/12px padding
- **Animation timing**: 30ms cubic-bezier transitions
- **Color scheme**: CSS custom properties with dark/light mode support
- **Accessibility**: WCAG 2.1 AA compliance foundation

### **Space Optimization Results**
- **Vertical space saved**: ~100px through typography and spacing optimization
- **Button efficiency**: 6px height reduction per button
- **Section headers**: 4px margin reduction
- **Total improvement**: ~15% more content visible in task pane

---

## 🔐 Security & Compliance

### **Production Security Features**
- **HTTPS enforcement**: All communications over SSL/TLS
- **Content Security Policy**: XSS attack prevention
- **Manifest validation**: Strict Microsoft schema compliance
- **Asset integrity**: Cloudflare Workers secure asset delivery

### **Development Security**
- **HTTP development server**: Certificate-free local development
- **Production isolation**: Development features hidden outside localhost
- **Secure asset pipeline**: Webpack content hashing and optimization

---

## 📚 Developer Resources

### **Essential Commands**
```bash
# Development
npm run dev-server     # HTTP hot-reload server
npm start             # VS Code integrated debugging
npm run validate      # Manifest validation

# Production
npm run build         # Full production build
git add . && git commit -m "message" && git push
```

### **Version Management**
**CRITICAL:** When releasing new versions, update version numbers in **ALL** these locations:
1. `manifest.xml` - Line 8: `<Version>2.0.0.0</Version>`
2. `src/taskpane/taskpane.ts` - Line 552: `const manifestVersion = '2.0.0.0';`
3. `src/taskpane/taskpane.css` - Line 1: `/* IPG Taxonomy Extractor v2.0.0 - Modern Fluent UI Styles */`
4. `src/taskpane/taskpane.ts` - Line 2: `* IPG Taxonomy Extractor v2.0.0`

**Note:** Version pill automatically displays the current version from the code. Failure to update all locations will result in version inconsistencies in the UI.

### **Key Files**
- `manifest.xml` - Office Add-in configuration (**VERSION CRITICAL**)
- `src/taskpane/taskpane.ts` - Core application logic (**VERSION CRITICAL**)
- `src/taskpane/taskpane.css` - Styles (**VERSION CRITICAL**)
- `webpack.dev.config.js` - Development server configuration
- `wrangler.toml` - Cloudflare Workers deployment configuration
- `.vscode/tasks.json` - VS Code debugging integration

### **Testing URLs**
- **Development**: `http://localhost:3001/taskpane.html`
- **Production**: `https://ipg-taxonomy-extractor-addin.wookstar.com/taskpane`

---

## 🎯 Success Metrics & Goals

### **Microsoft Store Certification Progress**
| Requirement | Status |
|-------------|--------|
| Manifest validation | ✅ Passing |
| HTTPS deployment | ✅ Live on Cloudflare |
| Icon assets | ✅ All variants present |
| API compatibility | ✅ ExcelApi 1.12 stable |
| Accessibility compliance | 🚧 Foundation complete |
| Localization support | ✅ Australian English |

### **Technical Excellence Targets**
- **Bundle size reduction**: Target 40% reduction through modern browser targeting
- **Load time optimization**: Target <2 seconds first load
- **WCAG 2.1 AA compliance**: 100% accessibility coverage
- **Test coverage**: 90%+ automated test coverage (Phase 5)

**Current Status: Production-ready Office Add-in with ongoing modernization for Microsoft Store certification and enhanced enterprise deployment capabilities.**