import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  // Use relative asset URLs (./assets/...) so the build works no matter what
  // subpath GitHub Pages serves it from (project site, custom domain, etc.)
  // rather than hard-coding an absolute /sales-heatmap/ prefix that 404s if the
  // serving path differs. The app has no client-side routing, so this is safe.
  base: './',
  plugins: [react()],
});
