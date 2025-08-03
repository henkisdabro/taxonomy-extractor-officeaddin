export default {
  async fetch(request: Request, env: any, ctx: any): Promise<Response> {
    // This is a basic pass-through worker.
    // Cloudflare's `[site]` configuration will handle serving static assets from the `dist` directory.
    // You can add custom logic here for things like API endpoints or modifying headers.
    try {
      // This will pass the request to the static asset handler.
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
