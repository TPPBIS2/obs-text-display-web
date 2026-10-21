import { defineConfig } from "vitest/config";
import react from "@vitejs/plugin-react";

export default defineConfig(({ mode }) => ({
  base: mode === "github-pages" ? "/obs-text-display-web/" : "/",
  plugins: [react()],
  build: {
    chunkSizeWarningLimit: 520,
  },
  test: {
    environment: "node",
  },
}));
