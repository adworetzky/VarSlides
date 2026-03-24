import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import basicSsl from "@vitejs/plugin-basic-ssl";

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react(), basicSsl()],
  server: {
    https: true,
    port: 3000,
  },
  build: {
    outDir: "dist",
    rollupOptions: {
      input: {
        taskpane: "src/taskpane/index.html",
      },
    },
  },
  resolve: {
    alias: {
      "@": "/src",
    },
  },
});
