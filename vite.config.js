import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// PENTING: ganti "omnifile" di bawah dengan nama repo GitHub kamu.
// Contoh: kalau repo-nya https://github.com/username/my-repo
// maka base harus "/my-repo/"
export default defineConfig({
  plugins: [react()],
  base: "/omnifile/",
});
