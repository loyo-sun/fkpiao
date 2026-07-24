import { defineConfig } from "vite";

export default defineConfig({
  base: process.env.SITES_APP_BUILD === "1" ? "/app/" : "/",
});
