import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import basicSsl from '@vitejs/plugin-basic-ssl'

// https://vitejs.dev/config/
export default defineConfig(({ command }) => {
  const isServe = command === 'serve';
  return {
    // GitHub Pages base path for production, root for dev
    base: command === 'build' ? '/outlook-org-chart/' : '/',
    plugins: [
      react(),
      // Only use SSL in dev mode
      ...(isServe ? [basicSsl()] : [])
    ],
    server: {
      port: 3000,
    }
  };
})
