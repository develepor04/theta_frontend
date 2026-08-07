import { defineConfig, loadEnv } from "vite";
import react from "@vitejs/plugin-react";
import fs from 'fs'
import tailwindcss from "@tailwindcss/vite"

export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, process.cwd(), '');
  const backendUrl = env.BACKEND_URL || 'http://localhost:5000';
  return {
  // base:"/",
  plugins: [react(),tailwindcss()],
  server: {
    https: {
      key: fs.readFileSync('./localhost-key.pem'),
      cert: fs.readFileSync('./localhost.pem')
    },
    port: 3000,
    proxy: {
      '/api': {
        target: backendUrl,
        changeOrigin: true,
        secure: backendUrl.startsWith('https://'),
      },
    },
  },
  preview: {
    port: 3000,
    allowedHosts: ['pmo.thetadynamics.io'],
    proxy: {
      '/api': {
        target: backendUrl,
        changeOrigin: true,
        secure: backendUrl.startsWith('https://'),
      },
    },
  },
  build: {
    chunkSizeWarningLimit: 2000,
  },
};
});
