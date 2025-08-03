
## Cloudflare Worker Deployment

This project is configured for deployment to Cloudflare Workers, enabling the Office Add-in to be served from Cloudflare's global network.

### Key Components for Cloudflare Deployment:

-   **`src/worker.ts`**: This TypeScript file contains the Cloudflare Worker's logic. It acts as the server for your add-in's static assets.
-   **`wrangler.toml`**: The configuration file for Cloudflare's `wrangler` CLI tool. It defines the Worker's name, entry point (`dist/worker.js`), and specifies the `dist` directory as the bucket for static assets.
-   **`webpack.worker.config.js`**: A dedicated Webpack configuration file specifically for bundling the `src/worker.ts` into `dist/worker.js`. This ensures the worker is compiled with the correct environment settings.
-   **`tsconfig.worker.json`**: A dedicated TypeScript configuration file for the Cloudflare Worker, ensuring it uses the appropriate `lib` and `types` (e.g., `@cloudflare/workers-types`) for a worker environment.

### Build Process for Cloudflare:

The `npm run build` command now orchestrates a two-step build process:

1.  **`npm run build:addin`**: Uses `webpack.config.js` to compile the main Office Add-in frontend (HTML, CSS, JavaScript) into the `dist` directory.
2.  **`npm run build:worker`**: First, `tsc --project tsconfig.worker.json` compiles `src/worker.ts` into `dist/worker.js`. Then, `wrangler deploy --dry-run` is executed to validate the worker deployment configuration (without actually deploying during local builds).

### Deployment to Cloudflare Pages:

When deploying via Cloudflare Pages (connected to GitHub):

-   **Build command**: `npm run build`
-   **Deploy command**: `npx wrangler deploy` (required by Cloudflare Pages to trigger the `wrangler` deployment after the build)
-   **Root directory**: `/`
-   **Build output directory**: `dist`

Cloudflare Pages will automatically run the build command, and then execute the deploy command, using `wrangler.toml` to understand how to deploy the `dist` folder contents and the `worker.js` script.

### File Cleanup:

-   The redundant `src/worker.js` file (a leftover from previous configurations) has been removed. The compiled worker is now exclusively generated in `dist/worker.js`.

### Dependency Warnings:

-   Warnings regarding `inflight`, `rimraf`, and `glob` are due to transitive dependencies of `clean-webpack-plugin`. As `clean-webpack-plugin` is at its latest version, these warnings are not directly actionable within this project and are expected to be resolved when its internal dependencies are updated.
