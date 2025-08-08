// vite.config.js
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  build: {
    outDir: 'dist',
    chunkSizeWarningLimit: 1500, // Meningkatkan batas peringatan ukuran bundle
    rollupOptions: {
      output: {
        manualChunks: {
          // Pisahkan vendor besar menjadi chunk terpisah
          react: ['react', 'react-dom'],
          supabase: ['@supabase/supabase-js'],
          preview: ['mammoth', 'xlsx', 'jszip'],
          icons: ['react-icons/fa'],
        },
      },
    },
  },
});