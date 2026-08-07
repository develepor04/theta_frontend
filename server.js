import express from 'express';
import fs from 'fs';
import http from 'http';
import https from 'https';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();

/** Load .env into process.env without overriding Azure/shell vars. */
function loadDotEnv(filePath) {
  if (!fs.existsSync(filePath)) return;
  for (const line of fs.readFileSync(filePath, 'utf8').split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#')) continue;
    const eq = trimmed.indexOf('=');
    if (eq === -1) continue;
    const key = trimmed.slice(0, eq).trim();
    let val = trimmed.slice(eq + 1).trim();
    if (
      (val.startsWith('"') && val.endsWith('"')) ||
      (val.startsWith("'") && val.endsWith("'"))
    ) {
      val = val.slice(1, -1);
    }
    if (process.env[key] === undefined) process.env[key] = val;
  }
}

loadDotEnv(path.join(__dirname, '.env'));

/**
 * Normalize Azure App Setting / .env → absolute origin URL.
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

  const host = parsed.hostname.toLowerCase();
  const isLocal = host === 'localhost' || host === '127.0.0.1' || host === '::1';
  if (!isLocal && !host.includes('.')) {
    console.error(
      `[proxy] BACKEND_URL hostname looks incomplete: ${parsed.hostname}. ` +
        'Check for a truncated value in Azure App Settings (must end with .azurewebsites.net).'
    );
    return null;
  }

  return parsed.origin; // protocol + host (+ port if any)
}

// Same var as Vite: BACKEND_URL (Azure App Setting overrides local .env)
const RAW_BACKEND_URL =
  process.env.BACKEND_URL ||
  process.env.API_URL ||
  process.env.VITE_API_URL ||
  'http://localhost:5000';

const BACKEND_ORIGIN = resolveBackendUrl(RAW_BACKEND_URL);

if (BACKEND_ORIGIN) {
  console.log(`Proxying /api → ${BACKEND_ORIGIN}`);
} else {
  console.warn(
    `[proxy] No valid backend URL (raw=${JSON.stringify(RAW_BACKEND_URL ?? null)}).`
  );
}

// Diagnostic — open in browser to verify the proxy env is loaded
app.get('/__health', (_req, res) => {
  const raw = RAW_BACKEND_URL == null ? null : String(RAW_BACKEND_URL);
  res.json({
    ok: true,
    backendConfigured: Boolean(BACKEND_ORIGIN),
    backendOrigin: BACKEND_ORIGIN,
    // Helps distinguish "not set" vs "set but invalid" without leaking full secrets
    backendUrlPresent: Boolean(raw && raw.trim()),
    backendUrlLength: raw ? raw.trim().length : 0,
    backendUrlEndsWithAzurewebsites: Boolean(
      raw && /\.azurewebsites\.net\/?$/i.test(raw.trim().replace(/^["']|["']$/g, ''))
    ),
  });
});

// Always handle /api — never fall through to index.html
app.use('/api', (req, res) => {
  if (!BACKEND_ORIGIN) {
    return res.status(503).json({
      detail: 'API proxy not configured',
      hint:
        'Set App Setting BACKEND_URL on the frontend App Service to your backend URL ' +
        '(e.g. https://theta-backend-xxxxx.canadacentral-01.azurewebsites.net), then restart.',
    });
  }

  const backend = new URL(BACKEND_ORIGIN);
  const transport = backend.protocol === 'https:' ? https : http;
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

app.use(express.static(path.join(__dirname, 'dist')));

app.use((req, res) => {
  res.sendFile(path.join(__dirname, 'dist', 'index.html'));
});

const PORT = process.env.PORT || 8080;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
