import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import { createProxyMiddleware } from 'http-proxy-middleware';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();

// Backend URL — set as Azure App Setting: BACKEND_URL
// Example: https://theta-backend-xxxxx.azurewebsites.net
const BACKEND_URL = process.env.BACKEND_URL;

if (BACKEND_URL) {
  app.use(
    '/api',
    createProxyMiddleware({
      target: BACKEND_URL,
      changeOrigin: true,
      secure: true,
      // Keep /api prefix — backend expects /api/...
      on: {
        error(err, req, res) {
          console.error('[proxy]', err.message);
          if (!res.headersSent) {
            res.writeHead(502, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ detail: 'Backend unavailable' }));
          }
        },
      },
    })
  );
  console.log(`Proxying /api → ${BACKEND_URL}`);
} else {
  console.warn(
    'BACKEND_URL is not set. /api requests will fall through to index.html. ' +
      'Set BACKEND_URL in Azure App Settings to your backend App Service URL.'
  );
}

// Serve static files from the Vite build output
app.use(express.static(path.join(__dirname, 'dist')));

// Fallback: serve index.html for SPA client-side routing (non-API routes only)
app.use((req, res) => {
  res.sendFile(path.join(__dirname, 'dist', 'index.html'));
});

const PORT = process.env.PORT || 8080;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
