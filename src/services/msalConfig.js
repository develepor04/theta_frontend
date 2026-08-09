import { PublicClientApplication } from '@azure/msal-browser';
import { getRuntimeEnv } from './runtimeEnv';

const TENANT_ID = getRuntimeEnv('VITE_AZURE_TENANT_ID');
const CLIENT_ID = getRuntimeEnv('VITE_AZURE_CLIENT_ID');

// SECURITY: VITE_AZURE_TENANT_ID must be set to your org's tenant ID.
// Leaving it unset falls back to 'common', which allows any Microsoft
// account (personal or any Azure AD tenant) to complete the OAuth flow.
// The backend DB lookup is the final gate, but setting the tenant ID
// restricts the OAuth popup to your org's directory as a first layer.
if (!TENANT_ID) {
  console.warn('[MSAL] VITE_AZURE_TENANT_ID is not set — falling back to "common". Set it to your Azure AD tenant ID to restrict login to your organisation.');
}

/**
 * Blank page for MSAL popup return — must NOT load the React app, or the
 * Theta login UI appears inside the OneDrive/Microsoft popup.
 * Register this exact URI in Azure AD → App registration → Redirect URIs:
 *   https://<your-frontend-host>/auth-redirect.html
 *   https://localhost:3000/auth-redirect.html  (local HTTPS)
 */
export const POPUP_REDIRECT_URI = `${window.location.origin}/auth-redirect.html`;

export const msalConfig = {
  auth: {
    clientId:    CLIENT_ID,
    authority:   `https://login.microsoftonline.com/${TENANT_ID || 'common'}`,
    // MSAL v5: must be the redirect-bridge page (auth-redirect.html).
    redirectUri: POPUP_REDIRECT_URI,
  },
  cache: {
    cacheLocation:        'sessionStorage',
    storeAuthStateInCookie: false,
  },
};

export const loginRequest = {
  scopes: ['User.Read'],
  redirectUri: POPUP_REDIRECT_URI,
};

export const TEAMS_DOMAIN = 'thetadynamic.io';

/** Returns true only when a real Client ID has been provided. */
export function isMsalConfigured() {
  return Boolean(CLIENT_ID && CLIENT_ID.trim().length > 10);
}

let _msalInstance = null;
let _initPromise = null;
/** Serializes interactive MSAL calls (popup/redirect) across the app. */
let _interactionTail = Promise.resolve();

export function getMsalInstance() {
  if (!_msalInstance) {
    _msalInstance = new PublicClientApplication(msalConfig);
  }
  return _msalInstance;
}

/**
 * Clear a stuck MSAL "interaction_in_progress" flag left when a popup was
 * closed on the React app instead of auth-redirect.html.
 */
export function clearStuckMsalInteraction() {
  try {
    const keys = [];
    for (let i = 0; i < sessionStorage.length; i += 1) {
      const key = sessionStorage.key(i);
      if (key && /interaction/i.test(key)) keys.push(key);
    }
    keys.forEach((key) => sessionStorage.removeItem(key));
  } catch (err) {
    console.warn('[MSAL] could not clear interaction cache', err);
  }
}

export async function initializeMsal() {
  if (_initPromise) return _initPromise;

  _initPromise = (async () => {
    const msal = getMsalInstance();
    try {
      await msal.initialize();
      const response = await msal.handleRedirectPromise();
      if (response?.account) {
        msal.setActiveAccount(response.account);
      } else {
        const accounts = msal.getAllAccounts();
        if (accounts.length > 0) {
          msal.setActiveAccount(accounts[0]);
        }
      }
    } catch (error) {
      // Stale interaction from a previous aborted popup — clear and continue.
      if (error?.errorCode === 'interaction_in_progress') {
        clearStuckMsalInteraction();
        return;
      }
      // Suppress benign no_token_request_cache_error (popup instead of redirect)
      if (error?.errorCode !== 'no_token_request_cache_error') {
        console.error('[MSAL] Initialization error:', error);
      }
    }
  })();

  return _initPromise;
}

async function runExclusiveInteraction(fn) {
  const previous = _interactionTail;
  let release;
  _interactionTail = new Promise((resolve) => { release = resolve; });
  try {
    try {
      await previous;
    } catch {
      // ignore prior failure
    }
    return await fn();
  } finally {
    release();
  }
}

/**
 * Interactive token (popup). Retries once after clearing a stuck interaction.
 */
export async function msalPopup(request, { preferSilent = true } = {}) {
  await initializeMsal();
  const msal = getMsalInstance();
  const scopes = request.scopes || ['User.Read'];
  const popupRequest = {
    ...request,
    scopes,
    redirectUri: request.redirectUri || POPUP_REDIRECT_URI,
  };

  return runExclusiveInteraction(async () => {
    const accounts = msal.getAllAccounts();
    if (preferSilent && accounts.length > 0) {
      const account = msal.getActiveAccount() || accounts[0];
      msal.setActiveAccount(account);
      try {
        const silent = await msal.acquireTokenSilent({ scopes, account });
        if (silent?.accessToken) return silent;
      } catch {
        // fall through to interactive
      }
    }

    const interactive = async () => {
      const accountsNow = msal.getAllAccounts();
      if (accountsNow.length > 0) {
        const account = msal.getActiveAccount() || accountsNow[0];
        return msal.acquireTokenPopup({ ...popupRequest, account });
      }
      return msal.loginPopup(popupRequest);
    };

    try {
      return await interactive();
    } catch (err) {
      if (err?.errorCode !== 'interaction_in_progress') throw err;
      clearStuckMsalInteraction();
      await new Promise((r) => setTimeout(r, 400));
      return interactive();
    }
  });
}
