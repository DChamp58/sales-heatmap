import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  // Served from a GitHub Pages project site at /sales-heatmap/, so assets must
  // be referenced relative to that subpath rather than the domain root.
  base: '/sales-heatmap/',
  plugins: [react()],
});
