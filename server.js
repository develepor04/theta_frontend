import express from 'express';
import http from 'http';
import https from 'https';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();

/**
 * Normalize Azure App Setting → absolute origin URL.
 * Rejects truncated / relative values that cause "TypeError: Invalid URL".
 */
function resolveBackendUrl(raw) {
  if (!raw) return null;

  let value = String(raw).trim();
  // Strip accidental wrapping quotes from App Settings
  if (
    (value.startsWith('"') && value.endsWith('"')) ||
    (value.startsWith("'") && value.endsWith("'"))
  ) {
    value = value.slice(1, -1).trim();
  }
  // Allow host-only values
  if (value && !/^https?:\/\//i.test(value)) {
    value = `https://${value}`;
  }
  // Drop trailing slash — we append paths ourselves
  value = value.replace(/\/+$/, '');

  let parsed;
  try {
    parsed = new URL(value);
  } catch {
    console.error(
      `[proxy] BACKEND_URL is not a valid URL: ${JSON.stringify(raw)}. ` +
        'Example: https://theta-backend-xxxxx.canadacentral-01.azurewebsites.net'
    );
    return null;
  }

  if (!parsed.hostname.includes('.')) {
    console.error(
      `[proxy] BACKEND_URL hostname looks incomplete: ${parsed.hostname}. ` +
        'Check for a truncated value in Azure App Settings (must end with .azurewebsites.net).'
    );
    return null;
  }

  return parsed.origin; // protocol + host (+ port if any)
}

const BACKEND_ORIGIN = resolveBackendUrl(process.env.BACKEND_URL);

if (BACKEND_ORIGIN) {
  const backend = new URL(BACKEND_ORIGIN);
  const transport = backend.protocol === 'https:' ? https : http;

  app.use('/api', (req, res) => {
    // Express strips the /api mount → restore full path for the backend
    const targetPath = req.originalUrl || `/api${req.url}`;

    const headers = { ...req.headers, host: backend.host };
    delete headers['content-length']; // let Node recalculate if body is piped

    const proxyReq = transport.request(
      {
        protocol: backend.protocol,
        hostname: backend.hostname,
        port: backend.port || (backend.protocol === 'https:' ? 443 : 80),
        path: targetPath,
        method: req.method,
        headers,
      },
      (proxyRes) => {
        res.writeHead(proxyRes.statusCode || 502, proxyRes.headers);
        proxyRes.pipe(res);
      }
    );

    proxyReq.on('error', (err) => {
      console.error('[proxy]', err.message);
      if (!res.headersSent) {
        res.status(502).json({ detail: 'Backend unavailable', error: err.message });
      }
    });

    req.pipe(proxyReq);
  });

  console.log(`Proxying /api → ${BACKEND_ORIGIN}`);
} else {
  console.warn(
    'BACKEND_URL is missing or invalid. /api requests will serve index.html. ' +
      'Set BACKEND_URL in Azure to e.g. https://your-backend.azurewebsites.net'
  );
}

app.use(express.static(path.join(__dirname, 'dist')));

app.use((req, res) => {
  res.sendFile(path.join(__dirname, 'dist', 'index.html'));
});

const PORT = process.env.PORT || 8080;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
