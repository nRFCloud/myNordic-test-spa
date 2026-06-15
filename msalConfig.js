const authority = B2C_AUTHORITY;
const clientId = B2C_CLIENT_ID;

const authorityHost = new URL(authority).hostname;
const appOrigin = "http://localhost:8080";

export const msalConfig = {
  auth: {
    clientId,
    authority,
    redirectUri: `${appOrigin}/mynordic/callback`,
    postLogoutRedirectUri: `${appOrigin}/`,
    knownAuthorities: [authorityHost],
  },
  cache: {
    cacheLocation: "sessionStorage",
  },
};

export const loginRequest = {
  scopes: ["openid", "profile"],
};
