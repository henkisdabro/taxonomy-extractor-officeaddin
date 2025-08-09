# Cloudflare Workers Deployment Requirements

## Critical Configuration for Successful Deployment

This document outlines the **exact requirements** for deploying Office Add-ins to Cloudflare Workers with Static Assets. These configurations have been tested and proven to work correctly.

## Problem Statement

Cloudflare Workers requires specific webpack and TypeScript configurations to properly bundle and deploy ES modules. Incorrect configurations result in:

- Empty worker.js files (0 bytes)
- "No event handlers were registered" deployment errors (code 10021)
- Failed deployments despite successful builds

## Required Files & Configurations

### 1. `webpack.worker.config.js` - CRITICAL CONFIGURATION

```javascript
const path = require('path');

module.exports = {
  entry: {
    worker: './src/worker.ts',
  },
  target: 'webworker',
  mode: 'production',
  output: {
    filename: '[name].js',
    path: path.resolve(__dirname, 'dist'),
    library: {
      type: 'module', // ESSENTIAL: Enables ES module output
    },
  },
  experiments: {
    outputModule: true, // ESSENTIAL: Required for ES module support
  },
  resolve: {
    extensions: ['.ts', '.js'],
  },
  module: {
    rules: [
      {
        test: /\.ts$/,
        exclude: /node_modules/,
        use: {
          loader: 'ts-loader',
          options: {
            configFile: 'tsconfig.worker.json',
            transpileOnly: false,
          },
        },
      },
    ],
  },
};
```

**Key Points:**

- `library.type: "module"` is MANDATORY for proper ES module output
- `experiments.outputModule: true` is REQUIRED for webpack to support ES modules
- Without these, worker.js will be empty (0 bytes)

### 2. `tsconfig.worker.json` - Required TypeScript Configuration

```json
{
  "compilerOptions": {
    "target": "es2022", // IMPORTANT: Must be es2022, not esnext
    "module": "es2022", // IMPORTANT: Must match target
    "lib": ["es2022"],
    "moduleResolution": "node",
    "esModuleInterop": true,
    "strict": true,
    "skipLibCheck": true,
    "forceConsistentCasingInFileNames": true,
    "allowSyntheticDefaultImports": true,
    "experimentalDecorators": true,
    "outDir": "./dist",
    "types": ["@cloudflare/workers-types"]
  },
  "include": ["src/worker.ts"],
  "exclude": ["node_modules"]
}
```

**Key Points:**

- Must use `es2022`, not `esnext` for better Cloudflare Workers compatibility
- `module` and `target` must match

### 3. `src/worker.ts` - Correct Worker Implementation

```typescript
interface Env {
  ASSETS: Fetcher;
}

export default {
  async fetch(request: Request, env: Env, ctx: ExecutionContext): Promise<Response> {
    // This worker uses Cloudflare Workers Static Assets to serve files from the dist directory.
    // The ASSETS binding provides access to static assets configured in wrangler.toml
    try {
      // Serve static assets using the new Workers Static Assets API
      return await env.ASSETS.fetch(request);
    } catch (e) {
      let pathname = new URL(request.url).pathname;
      return new Response(`"${pathname}" not found`, {
        status: 404,
        statusText: 'not found',
      });
    }
  },
};
```

**Key Points:**

- Must use `export default` with object containing `fetch` handler
- ASSETS binding must be properly typed in Env interface
- Error handling for 404 responses is recommended

### 4. `wrangler.toml` - Cloudflare Workers Static Assets Configuration

```toml
name = "ipg-taxonomy-extractor-addin"
main = "dist/worker.js"
compatibility_date = "2025-08-03"

[assets]
directory = "./dist"
binding = "ASSETS"
```

**Key Points:**

- `main` must point to the built worker.js file
- `[assets]` section enables Static Assets API (not deprecated `[site]`)
- `binding = "ASSETS"` must match the interface in worker.ts

### 5. `package.json` - Build Scripts

```json
{
  "scripts": {
    "build": "npm run clean && npm run build:addin && npm run build:worker",
    "build:addin": "webpack --mode production --config webpack.config.js",
    "build:worker": "webpack --mode production --config webpack.worker.config.js",
    "clean": "rimraf dist"
  }
}
```

**Key Points:**

- Worker build MUST use webpack, not direct TypeScript compilation
- Separate configurations ensure proper bundling for each target

## Verification Steps

### 1. Local Build Verification

```bash
npm run build:worker
```

**Expected Output:**

```
asset worker.js 388 bytes [emitted] [javascript module] [minimized] (name: worker)
```

**❌ Failed Output:**

```
asset worker.js 0 bytes [emitted] [minimized] (name: worker)
```

### 2. Worker Content Verification

```bash
cat dist/worker.js
```

**Expected:** Minified JavaScript code containing fetch handler
**❌ Failed:** Empty file or no file

### 3. Deployment Success Indicators

**✅ Successful Deployment Log:**

```
🌀 Found 1 new or modified static asset to upload. Proceeding with upload...
+ /worker.js
✨ Success! Uploaded 1 file (13 already uploaded) (1.01 sec)
Total Upload: 0.77 KiB / gzip: 0.45 KiB

Your Worker has access to the following bindings:
Binding            Resource
env.ASSETS         Assets

Uploaded ipg-taxonomy-extractor-addin (5.91 sec)
Deployed ipg-taxonomy-extractor-addin triggers (0.28 sec)
```

**❌ Failed Deployment:**

```
✘ [ERROR] A request to the Cloudflare API failed.
No event handlers were registered. This script does nothing.
[code: 10021]
```

## Common Mistakes to Avoid

1. **Missing ES Module Configuration**: Without `library.type: "module"` and `experiments.outputModule: true`, webpack produces empty output
2. **Using esnext instead of es2022**: TypeScript target `esnext` can cause compatibility issues
3. **Direct TypeScript Compilation**: Using `tsc` instead of webpack bypasses bundling and optimization
4. **Wrong wrangler.toml Configuration**: Using deprecated `[site]` instead of `[assets]`

## Troubleshooting Checklist

If deployment fails:

1. ✅ Check worker.js file size > 0 bytes
2. ✅ Verify webpack.worker.config.js has `library.type: "module"`
3. ✅ Verify webpack.worker.config.js has `experiments.outputModule: true`
4. ✅ Confirm tsconfig.worker.json uses `es2022` not `esnext`
5. ✅ Ensure build:worker script uses webpack, not tsc
6. ✅ Verify wrangler.toml uses `[assets]` not `[site]`

## Success Metrics

- **Worker Size**: ~388 bytes (minified)
- **Build Time**: ~1-2 seconds
- **Upload Size**: ~0.77 KiB
- **Deployment Time**: ~6 seconds total
- **Assets Uploaded**: 14-15 files including worker.js

This configuration has been tested and verified to work with:

- Webpack 5.101.0
- Wrangler 4.27.0
- Node.js 22.16.0
- Cloudflare Workers Static Assets API

Following these exact configurations will ensure successful deployment of Office Add-ins to Cloudflare Workers.
