import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react-swc'
import mkcert from 'vite-plugin-mkcert'

export default defineConfig({
  plugins: [react(), mkcert()],
  server: {
    port: 3000,
    https: true as any,
    host: true, 
  },
  base: './', // Essential for GitHub Pages relative paths
  build: {
    outDir: 'dist',
  }
})
