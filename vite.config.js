import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  // Relative base so the build works regardless of the repo name /
  // whatever sub-path GitHub Pages serves it from.
  base: "./",
});
