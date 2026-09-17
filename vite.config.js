import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// PENTING: ganti "omnifile" di bawah dengan nama repo GitHub kamu.
// Contoh: kalau repo-nya https://github.com/username/my-repo
// maka base harus "/my-repo/"
// Jika digunakan sebagai user page (username.github.io), gunakan base: "/"

export default defineConfig({
  plugins: [
    react(),
  ],
  base: '/omnifile/',
  build: {
    // Optimasi build
    minify: 'esbuild',
    sourcemap: false,
    rollupOptions: {
      output: {
        // Split chunks untuk caching yang lebih baik
        manualChunks: {
          vendor: ['react', 'react-dom'],
          lucide: ['lucide-react'],
          xlsx: ['xlsx'],
          mammoth: ['mammoth'],
          papaparse: ['papaparse'],
        },
        // Kompres nama file
        entryFileNames: 'assets/[name]-[hash].js',
        chunkFileNames: 'assets/[name]-[hash].js',
        assetFileNames: 'assets/[name]-[hash].[ext]',
      },
    },
    // Batas ukuran chunk
    chunkSizeWarningLimit: 1000,
  },
  server: {
    // Port default
    port: 5173,
    // Buka otomatis di browser
    open: true,
    // Host untuk akses dari jaringan lokal
    host: true,
  },
  preview: {
    port: 4173,
    open: true,
  },
  // Optimasi asset
  assetsInclude: ['**/*.png', '**/*.jpg', '**/*.jpeg', '**/*.svg'],
});