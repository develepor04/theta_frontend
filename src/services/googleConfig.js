const GOOGLE_CLIENT_ID = import.meta.env.VITE_GOOGLE_CLIENT_ID || '';

/** Read-only Drive access for the Catalog file picker. */
export const GOOGLE_DRIVE_READONLY_SCOPE = 'https://www.googleapis.com/auth/drive.readonly';

export function isGoogleConfigured() {
  return Boolean(GOOGLE_CLIENT_ID && GOOGLE_CLIENT_ID.trim().length > 10);
}

/**
 * Opens Google's OAuth popup and resolves with an access token.
 * Requires the GIS script to be loaded (index.html).
 * @param {string} [scope]
 * @returns {Promise<string>}
 */
function requestGoogleAccessToken(scope) {
  return new Promise((resolve, reject) => {
    if (!window.google?.accounts?.oauth2) {
      reject(new Error('Google Identity Services not loaded yet. Please refresh and try again.'));
      return;
    }
    if (!GOOGLE_CLIENT_ID) {
      reject(new Error('Google sign-in is not configured (VITE_GOOGLE_CLIENT_ID missing).'));
      return;
    }

    const client = window.google.accounts.oauth2.initTokenClient({
      client_id: GOOGLE_CLIENT_ID,
      scope,
      callback: (response) => {
        if (response.error) {
          reject(new Error(response.error_description || response.error));
        } else {
          resolve(response.access_token);
        }
      },
      error_callback: (err) => {
        if (err?.type === 'popup_closed') {
          reject(Object.assign(new Error('Popup closed'), { cancelled: true }));
        } else {
          reject(new Error(err?.message || 'Google authorization failed'));
        }
      },
    });

    client.requestAccessToken({ prompt: 'select_account' });
  });
}

/** Google SSO (login / signup). */
export function googleSignInPopup() {
  return requestGoogleAccessToken('openid email profile');
}

/** Google Drive picker — read-only file access. */
export function googleDriveSignInPopup() {
  return requestGoogleAccessToken(GOOGLE_DRIVE_READONLY_SCOPE);
}
