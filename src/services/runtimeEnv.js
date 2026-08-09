/**
 * Resolve public SPA config from Vite build-time env, then Azure runtime
 * injection (window.__RUNTIME_ENV__ from /env.js via server.js).
 *
 * Azure App Settings are NOT available to Vite at GitHub Actions build time
 * unless passed as secrets — runtime fallback lets App Settings work after deploy.
 */
export function getRuntimeEnv(key) {
  const fromVite = import.meta.env?.[key];
  if (typeof fromVite === 'string' && fromVite.trim()) return fromVite.trim();

  try {
    const fromRuntime = window.__RUNTIME_ENV__?.[key];
    if (typeof fromRuntime === 'string' && fromRuntime.trim()) return fromRuntime.trim();
  } catch {
    // SSR / non-browser — ignore
  }

  return '';
}
