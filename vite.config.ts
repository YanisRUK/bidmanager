import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import path from "path";

// https://vite.dev/config/
export default defineConfig(({ command }) => ({
  plugins: [react()],

  resolve: {
    alias:
      command === "serve"
        ? {
            // In dev, swap the SDK for a local stub so the app runs without
            // the real package being installed.
            "@microsoft/powerapps-code-apps": path.resolve(
              __dirname,
              "src/lib/powerAppsSdkStub.ts"
            ),
          }
        : undefined,
  },

  build: {
    rollupOptions: {
      // The Power Apps SDK is provided by the runtime host (Power Apps).
      // Mark it as external so Rollup doesn't bundle it in production.
      external: ["@microsoft/powerapps-code-apps"],
    },
  },
}));
