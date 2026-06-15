const authority = B2C_AUTHORITY;
const clientId = B2C_CLIENT_ID;

const authorityHost = new URL(authority).hostname;
const appOrigin = "http://localhost:8080";

/** Interactive redirect — must match the origin that calls loginRedirect. */
const redirectUri = `${appOrigin}/mynordic/from-portal`;

/**
 * MSAL v5 redirect bridge for ssoSilent (must call broadcastResponseToMainFrame).
 * @see https://learn.microsoft.com/en-us/entra/msal/javascript/browser/redirect-bridge
 *
 * Register on the app registration:
 *   - http://localhost:8080/mynordic/from-portal
 *   - http://localhost:8080/mynordic/sso.html
 */
const ssoRedirectUri = `${appOrigin}/mynordic/sso.html`;

/** Register under Authentication → Front-channel logout / Redirect URI. */
const postLogoutRedirectUri = `${appOrigin}/`;

export const msalConfig = {
  auth: {
    clientId,
    authority,
    redirectUri,
    postLogoutRedirectUri,
    knownAuthorities: [authorityHost],
  },
  cache: {
    cacheLocation: "localStorage",
  },
};

export const loginRequest = {
  scopes: ["openid", "profile"],
};

export const ssoSilentRequest = {
  scopes: ["openid", "profile"],
  redirectUri: ssoRedirectUri,
};
