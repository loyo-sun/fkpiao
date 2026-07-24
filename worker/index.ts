import handler from "vinext/server/app-router-entry";

interface Env {
  ASSETS: Fetcher;
}

interface ExecutionContext {
  waitUntil(promise: Promise<unknown>): void;
  passThroughOnException(): void;
}

export default {
  async fetch(request: Request, env: Env, ctx: ExecutionContext): Promise<Response> {
    const url = new URL(request.url);

    if (url.pathname === "/") {
      url.pathname = "/app/index.html";
      return env.ASSETS.fetch(new Request(url, request));
    }

    if (url.pathname.startsWith("/app/")) {
      return env.ASSETS.fetch(request);
    }

    return handler.fetch(request, env, ctx);
  },
};
