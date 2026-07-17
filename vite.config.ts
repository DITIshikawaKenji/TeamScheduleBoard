import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

const buildTime = new Date()
  .toLocaleString("ja-JP", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  });

export default defineConfig({
  plugins: [react()],

  define: {
    __BUILD_TIME__: JSON.stringify(buildTime),
  },

  base: "/TeamScheduleBoard/",
});