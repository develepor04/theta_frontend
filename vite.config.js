import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import fs from 'fs'
import tailwindcss from "@tailwindcss/vite"
export default defineConfig({
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
        // Local backend: 'http://localhost:5000'
        // Azure backend (same as BACKEND_URL):
        target: 'https://theta-backend-a2d7g4ash4ddhmc3.canadacentral-01.azurewebsites.net',
        changeOrigin: true,
        secure: true,
      },
    },
  },
  preview: {
    port: 3000,
    allowedHosts: ['pmo.thetadynamics.io'],
    proxy: {
      '/api': {
        target: 'https://pmo-backend.thetadynamics.io',
        changeOrigin: true,
        secure: true,
      },
    },
  },
  build: {
    chunkSizeWarningLimit: 2000,
  },
});