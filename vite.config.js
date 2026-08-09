import { defineConfig, loadEnv } from "vite";
import react from "@vitejs/plugin-react";
import fs from 'fs'
import tailwindcss from "@tailwindcss/vite"

/** Dev/preview: mirror production /env.js so index.html does not 404. */
function runtimeEnvPlugin(env) {
  const payload = {
    VITE_AZURE_CLIENT_ID: env.VITE_AZURE_CLIENT_ID || '',
    VITE_AZURE_TENANT_ID: env.VITE_AZURE_TENANT_ID || '',
    VITE_GOOGLE_CLIENT_ID: env.VITE_GOOGLE_CLIENT_ID || '',
  };
  const body = `window.__RUNTIME_ENV__=${JSON.stringify(payload)};`;
  const middleware = (req, res, next) => {
    if (req.url?.split('?')[0] !== '/env.js') return next();
    res.setHeader('Content-Type', 'application/javascript');
    res.setHeader('Cache-Control', 'no-store');
    res.end(body);
  };
  return {
    name: 'runtime-env',
    configureServer(server) {
      server.middlewares.use(middleware);
    },
    configurePreviewServer(server) {
      server.middlewares.use(middleware);
    },
  };
}

export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, process.cwd(), '');
  const backendUrl = env.BACKEND_URL || 'http://localhost:5000';
  return {
  // base:"/",
  plugins: [react(), tailwindcss(), runtimeEnvPlugin(env)],
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
