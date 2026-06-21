import { defineConfig } from "vite";

// The demo lives in demo/; build it to demo-dist/ for GitHub Pages. base "./" so it
// works from any repo-subpath. Vitest uses jsdom for the DOM-based round-trip tests.
export default defineConfig({
  root: "demo",
  base: "./",
  build: { outDir: "../demo-dist", emptyOutDir: true },
  test: {
    root: ".",
    globals: true,
    environment: "jsdom",
    include: ["src/**/*.test.ts"],
  },
});
