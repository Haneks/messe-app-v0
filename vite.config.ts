import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react()],
  server: {
    proxy: {
      '/api/aelf': {
        target: 'https://api.aelf.org',
        changeOrigin: true,
        rewrite: (path) => path.replace(/^\/api\/aelf/, '')
      }
    }
  },
  optimizeDeps: {
    exclude: ['lucide-react'],
  },
});