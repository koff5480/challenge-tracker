// @ts-nocheck
import tailwindcss from "@tailwindcss/vite";
import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";

export default defineConfig({
  // 本地預覽用 "./"；GitHub Actions 時由 VITE_BASE_PATH 注入 "/challenge-tracker/"
  base: process.env.VITE_BASE_PATH || './',
  plugins: [react(), tailwindcss()],
  server: {
    allowedHosts: [".sandbox.entrydesk.com"],
  },
});
