// OAuth 2.0 Authentication via Google Identity Services (GIS)

// ── Deployer configuration ──────────────────────────────────────────
// Set your OAuth Client ID here. End users do NOT need to change this.
// Create one at: Google Cloud Console → APIs & Services → Credentials
const OAUTH_CLIENT_ID = '174725674241-vnr83t2ck5avn6l9jeo9veud8heh9ptn.apps.googleusercontent.com';
// ─────────────────────────────────────────────────────────────────────

class AuthManager {
    constructor() {
        this._accessToken = null;
        this._expiresAt = 0;       // epoch ms when the current token expires
        this._tokenClient = null;
        this._clientId = '';
        this._userEmail = '';
        this._resolveToken = null; // for the sign-in promise callback
        this._rejectToken = null;

        this.ready = this._init();
    }

    // ── Initialisation ──────────────────────────────────────────────────

    async _init() {
        // Priority: hardcoded constant > localStorage override (for self-hosters)
        this._clientId = OAUTH_CLIENT_ID || localStorage.getItem('oauth-client-id') || '';
        this._userEmail = localStorage.getItem('oauth-user-email') || '';

        // Restore persisted token if it hasn't expired yet
        this._restoreToken();

        // Handle OAuth2 redirect callback (mobile sign-in flow) — must run before
        // any early return so the token is recovered on the redirect-back page load.
        this._handleRedirectCallback();

        if (!this._clientId) {
            console.log('[Auth] No OAuth Client ID configured');
            return;
        }

        await this._waitForGIS();
        this._createTokenClient();
        console.log('[Auth] AuthManager initialised');
    }

    /** Returns true when running on a mobile browser where popups are unreliable. */
    _isMobileBrowser() {
        return /Android|iPhone|iPad|iPod/i.test(navigator.userAgent);
    }

    /**
     * Parse an OAuth2 implicit-flow redirect callback from the URL hash.
     * Google appends #access_token=…&expires_in=… after the user signs in.
     * Cleans the hash from the URL so it is not shown to the user.
     */
    _handleRedirectCallback() {
        if (!window.location.hash) return false;

        const params = new URLSearchParams(window.location.hash.substring(1));
        const token    = params.get('access_token');
        const errorVal = params.get('error');

        // Always clean the fragment from the URL bar
        history.replaceState(null, '', window.location.pathname + window.location.search);

        if (errorVal) {
            console.warn('[Auth] Redirect callback returned error:', errorVal);
            this._redirectError = errorVal;
            return false;
        }

        if (token) {
            const expiresIn = parseInt(params.get('expires_in') || '3600', 10);
            this._accessToken = token;
            this._expiresAt   = Date.now() + (expiresIn - 60) * 1000;
            this._persistToken();
            console.log('[Auth] Access token recovered from redirect callback');
            this._fetchUserInfo();
            return true;
        }

        return false;
    }

    /**
     * Redirect the browser to Google OAuth2 (implicit / token flow).
     * Used on mobile where popup-based GIS does not work reliably.
     * The returned Promise never resolves — the page navigates away immediately.
     */
    _signInWithRedirect(prompt = '') {
        // Normalise the redirect URI: strip index.html so it always matches
        // the bare directory URL registered in Google Cloud Console.
        let pathname = window.location.pathname;
        if (pathname.endsWith('/index.html')) {
            pathname = pathname.slice(0, -'index.html'.length);
        }
        const redirectUri = window.location.origin + pathname;
        const params = new URLSearchParams({
            client_id:              this._clientId,
            redirect_uri:           redirectUri,
            response_type:          'token',
            scope:                  'https://www.googleapis.com/auth/drive.file',
            include_granted_scopes: 'true',
        });
        if (prompt) params.set('prompt', prompt);

        // Mark that we expect a redirect return so app.js can resume the flow
        sessionStorage.setItem('oauth-redirect-pending', '1');

        window.location.href = `https://accounts.google.com/o/oauth2/v2/auth?${params}`;
        return new Promise(() => {}); // never resolves; page navigates away
    }

    /** Persist current token to localStorage so it survives page reloads. */
    _persistToken() {
        if (this._accessToken && this._expiresAt > Date.now()) {
            localStorage.setItem('oauth-access-token', this._accessToken);
            localStorage.setItem('oauth-token-expires', String(this._expiresAt));
        }
    }

    /** Restore token from localStorage if it hasn't expired. */
    _restoreToken() {
        const token = localStorage.getItem('oauth-access-token');
        const expires = parseInt(localStorage.getItem('oauth-token-expires') || '0', 10);
        if (token && expires > Date.now()) {
            this._accessToken = token;
            this._expiresAt = expires;
            console.log('[Auth] Restored token from storage, expires in', Math.round((expires - Date.now()) / 1000), 's');
        } else {
            // Clean up stale values
            localStorage.removeItem('oauth-access-token');
            localStorage.removeItem('oauth-token-expires');
        }
    }

    /** Clear persisted token from localStorage. */
    _clearPersistedToken() {
        localStorage.removeItem('oauth-access-token');
        localStorage.removeItem('oauth-token-expires');
    }

    /** Wait until the Google Identity Services library is loaded. */
    _waitForGIS() {
        return new Promise((resolve) => {
            if (window.google?.accounts?.oauth2) {
                resolve();
                return;
            }
            // Poll every 100 ms (the <script> tag is async)
            const id = setInterval(() => {
                if (window.google?.accounts?.oauth2) {
                    clearInterval(id);
                    resolve();
                }
            }, 100);
            // Give up after 15 s
            setTimeout(() => { clearInterval(id); resolve(); }, 15000);
        });
    }

    _createTokenClient() {
        if (!window.google?.accounts?.oauth2) return;

        this._tokenClient = google.accounts.oauth2.initTokenClient({
            client_id: this._clientId,
            scope: 'https://www.googleapis.com/auth/drive.file',
            callback: (resp) => this._onTokenResponse(resp),
            error_callback: (err) => this._onTokenError(err),
        });
    }

    // ── Token lifecycle ─────────────────────────────────────────────────

    _onTokenResponse(resp) {
        if (resp.error) {
            console.error('[Auth] Token error:', resp.error);
            this._rejectToken?.(new Error(resp.error));
            this._resolveToken = null;
            this._rejectToken = null;
            return;
        }

        this._accessToken = resp.access_token;
        this._expiresAt = Date.now() + (resp.expires_in - 60) * 1000; // buffer 60 s
        this._persistToken();
        console.log('[Auth] Token obtained, expires in', resp.expires_in, 's');

        // Fetch the user's email for display (best-effort)
        this._fetchUserInfo();

        this._resolveToken?.(this._accessToken);
        this._resolveToken = null;
        this._rejectToken = null;
    }

    _onTokenError(err) {
        console.error('[Auth] GIS error callback:', err);
        const msg = (err && err.message) || (err && err.type) || 'Unknown GIS error';
        this._rejectToken?.(new Error(msg));
        this._resolveToken = null;
        this._rejectToken = null;
    }

    async _fetchUserInfo() {
        try {
            const r = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
                headers: { Authorization: `Bearer ${this._accessToken}` }
            });
            if (r.ok) {
                const info = await r.json();
                this._userEmail = info.email || '';
                localStorage.setItem('oauth-user-email', this._userEmail);
            }
        } catch (_) { /* best effort */ }
    }

    // ── Public API ──────────────────────────────────────────────────────

    /** Interactive sign-in (shows Google consent popup, or redirects on mobile). */
    signIn() {
        if (!this._tokenClient) {
            return Promise.reject(new Error('OAuth not configured — set a Client ID first'));
        }
        // On Android/mobile, the GIS popup cannot post back to window.opener;
        // use the OAuth2 redirect (implicit) flow instead.
        if (this._isMobileBrowser()) {
            return this._signInWithRedirect('consent');
        }
        return new Promise((resolve, reject) => {
            // Timeout: if GIS callback never fires (popup blocked, origin mismatch, etc.)
            const timeout = setTimeout(() => {
                this._resolveToken = null;
                this._rejectToken = null;
                console.error('[Auth] Sign-in timed out after 120 s');
                reject(new Error('Sign-in timed out. Make sure popups are allowed and try again.'));
            }, 120000);

            this._resolveToken = (token) => { clearTimeout(timeout); resolve(token); };
            this._rejectToken  = (err)   => { clearTimeout(timeout); reject(err); };

            console.log('[Auth] Requesting access token (consent prompt)...');
            this._tokenClient.requestAccessToken({ prompt: 'consent' });
        });
    }

    /** Silent token refresh (no popup if user has an active Google session). */
    silentRefresh() {
        if (!this._tokenClient) {
            return Promise.reject(new Error('OAuth not configured'));
        }
        // On mobile the GIS popup flow is unreliable; reject here so that
        // getValidToken() falls through to the interactive redirect sign-in.
        if (this._isMobileBrowser()) {
            return Promise.reject(new Error('Silent refresh not supported on mobile'));
        }
        return new Promise((resolve, reject) => {
            const timeout = setTimeout(() => {
                this._resolveToken = null;
                this._rejectToken = null;
                reject(new Error('Silent refresh timed out'));
            }, 15000);

            this._resolveToken = (token) => { clearTimeout(timeout); resolve(token); };
            this._rejectToken  = (err)   => { clearTimeout(timeout); reject(err); };

            this._tokenClient.requestAccessToken({ prompt: '' });
        });
    }

    /**
     * Return a valid access token. Tries silent refresh if expired,
     * falls back to interactive sign-in.
     */
    async getValidToken() {
        if (this._accessToken && Date.now() < this._expiresAt) {
            return this._accessToken;
        }

        // Try silent refresh first
        try {
            return await this.silentRefresh();
        } catch (_) {
            // Silent refresh failed — fall through to interactive
        }

        return this.signIn();
    }

    isSignedIn() {
        return !!(this._accessToken && Date.now() < this._expiresAt);
    }

    get userEmail() { return this._userEmail; }
    get clientId()  { return this._clientId; }

    /** Sign out: revoke the token and clear local state. */
    async signOut() {
        if (this._accessToken) {
            try {
                google.accounts.oauth2.revoke(this._accessToken);
            } catch (_) { /* best effort */ }
        }
        this._accessToken = null;
        this._expiresAt = 0;
        this._userEmail = '';
        this._clearPersistedToken();
        localStorage.removeItem('oauth-user-email');
        localStorage.removeItem('oauth-spreadsheet-id');
        console.log('[Auth] Signed out');
    }

    /**
     * (Re)configure the OAuth Client ID. Called from Settings UI or
     * on first run.
     */
    async configure(clientId) {
        this._clientId = (clientId || '').trim();
        localStorage.setItem('oauth-client-id', this._clientId);

        // Reset token state
        this._accessToken = null;
        this._expiresAt = 0;

        if (this._clientId) {
            await this._waitForGIS();
            this._createTokenClient();
        }
    }
}

// Singleton
window.AuthManager = new AuthManager();
