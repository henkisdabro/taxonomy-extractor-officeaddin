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
        statusText: "not found",
      });
    }
  },
};
