# Testing the myNordic OAuth Flow (Azure AD B2C)

This guide explains how to stand up a minimal SPA to end-to-end test the
authorization-code flow with PKCE against an Azure AD B2C (CIAM) instance using
`@azure/msal-browser`. Use it to verify that your tenant, user-flow/policy, and
client registration are configured correctly before integrating with the
production app.

## Prerequisites

| Requirement                                | Notes                                                                                          |
| ------------------------------------------ | ---------------------------------------------------------------------------------------------- |
| Node.js >= 24                              | Any LTS release works for the test harness                                                     |
| An Azure AD B2C / Entra External ID tenant | The **authority** URL, e.g. `https://<tenant>.ciamlogin.com/`                                  |
| A registered SPA application               | Gives you a **client ID** (Application ID)                                                     |
| A redirect URI registered on the app       | Must match exactly what the SPA sends — use `http://localhost:8080/mynordic/callback` for local testing |

### Azure portal checklist

1. **App registration → Authentication → Platform → Single-page application**
   Add `http://localhost:8080/mynordic/callback` as a redirect URI.
2. **API permissions** — grant `openid` and `profile` (and `User.Read` if you
   need to call Microsoft Graph for the `mail` property, as the production app
   does).
3. **Implicit grant** — leave disabled; MSAL v3+ uses the authorization-code
   flow with PKCE by default.
4. If your tenant uses a custom **user flow / policy**, ensure the authority URL
   includes it, e.g.
   `https://<tenant>.b2clogin.com/<tenant>.onmicrosoft.com/<policy>`. For Entra
   External ID (CIAM) tenants the authority is typically just
   `https://<tenant>.ciamlogin.com/`.
5. **Token configuration → Optional claims → ID** — add `login_hint` for the nrfCloud
   SPA client. MSAL uses this claim as `logout_hint` so CIAM can end the server
   session without showing “Pick an account” (especially after myNordic portal SSO).
6. **Authentication → Redirect URI** — register the SPA **logout** URL (e.g.
   `http://localhost:8080/` or your nrfCloud origin) as a front-channel logout /
   post-logout redirect URI, matching `postLogoutRedirectUri` in `msalConfig.js`.

### Logout after myNordic portal SSO (nrfCloud)

If users sign in on myNordic and land in nrfCloud via `login_hint` / silent SSO,
then log out from nrfCloud and CIAM shows **“Pick an account”** with **no accounts
listed**, the logout request did not identify the CIAM session. Typical causes:

| Cause | Fix |
| ----- | --- |
| `logoutRedirect()` called with no `account` | Pass `msal.getAllAccounts()[0]` (or the active account). |
| No `logout_hint` on the end-session URL | Pass `logoutHint` (email from myNordic `login_hint`, or `account.idTokenClaims.login_hint`). |
| `login_hint` not in the ID token | Enable the optional claim on the **nrfCloud** app registration (see checklist item 5). |
| Missing post-logout redirect | Set `postLogoutRedirectUri` in MSAL config and register it in Azure. |

The `mynordic-oauth-test` harness persists `login_hint` from the portal redirect
and sends it on logout. Mirror that in nrfCloud production code.

## 1 — Scaffold the test SPA

```bash
mkdir mynordic-oauth-test && cd mynordic-oauth-test
npm init -y
npm pkg set type=module
npm install @azure/msal-browser vite
```

Setting `type` to `"module"` is required because the source files use ES module
`import`/`export` syntax.

## 2 — Configure environment variables

The SPA reads the Azure AD B2C authority and client ID from environment
variables at build/dev time via Vite's `define` option. The E2E tests also need
test-user credentials. Create an `.envrc` (for
[direnv](https://direnv.net/)) or export the variables in your shell:

```bash
export B2C_AUTHORITY="https://<tenant>.ciamlogin.com/"
export B2C_CLIENT_ID="<your-client-id>"

export TEST_USER_EMAIL="<test-user-email>"
export TEST_USER_PASSWORD="<test-user-password>"
```

`B2C_AUTHORITY` and `B2C_CLIENT_ID` are injected into the browser code as
compile-time constants by Vite's `define` config. The `TEST_USER_*` variables
are used by the Playwright E2E tests only.

## 3 — Create the configuration

Create `msalConfig.js`:

```javascript
const authority = B2C_AUTHORITY;
const clientId = B2C_CLIENT_ID;

const authorityHost = new URL(authority).hostname;

export const msalConfig = {
  auth: {
    clientId,
    authority,
    redirectUri: "http://localhost:8080/mynordic/callback",
    knownAuthorities: [authorityHost],
  },
  cache: {
    cacheLocation: "sessionStorage",
  },
};

export const loginRequest = {
  scopes: ["openid", "profile"],
};
```

`knownAuthorities` tells MSAL to trust the B2C/CIAM issuer instead of defaulting
to `login.microsoftonline.com`.

## 4 — Create the HTML entry point

Create `index.html`:

```html
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <title>myNordic OAuth Test</title>
  </head>
  <body>
    <h1>myNordic OAuth Test Harness</h1>
    <button id="login">Log in with myNordic</button>
    <button id="logout" hidden>Log out</button>
    <pre id="output">Not authenticated.</pre>
    <script type="module" src="./main.js"></script>
  </body>
</html>
```

## 5 — Implement the flow

Create `main.js`:

```javascript
import { PublicClientApplication } from "@azure/msal-browser";
import { msalConfig, loginRequest } from "./msalConfig.js";

const msal = new PublicClientApplication(msalConfig);

async function init() {
  await msal.initialize();

  // Handle the redirect callback (runs on /callback after Azure redirects back)
  const authResult = await msal.handleRedirectPromise();

  if (authResult) {
    show(authResult);
    return;
  }

  // Check for an existing session
  const accounts = msal.getAllAccounts();
  if (accounts.length > 0) {
    show({ account: accounts[0], note: "restored from session cache" });
  }
}

function show(result) {
  document.getElementById("output").textContent = JSON.stringify(
    result,
    null,
    2,
  );
  document.getElementById("login").hidden = true;
  document.getElementById("logout").hidden = false;
}

document.getElementById("login").addEventListener("click", () => {
  msal.loginRedirect(loginRequest);
});

document.getElementById("logout").addEventListener("click", () => {
  msal.logoutRedirect();
});

init();
```

## 6 — Add a Vite config

Create `vite.config.js`:

```javascript
import { defineConfig } from "vite";

export default defineConfig({
  define: {
    B2C_AUTHORITY: JSON.stringify(process.env.B2C_AUTHORITY),
    B2C_CLIENT_ID: JSON.stringify(process.env.B2C_CLIENT_ID),
  },
  server: {
    port: 8080,
  },
});
```

`define` replaces the identifiers `B2C_AUTHORITY` and `B2C_CLIENT_ID`
with their string values at compile time. The port must match the redirect URI
registered in Azure.

## 7 — Run the test

```bash
npx vite
```

Open `http://localhost:8080`, click **Log in with myNordic**, authenticate in
the Azure-hosted login page, and verify that:

1. You are redirected back to `http://localhost:8080/mynordic/callback`.
2. `handleRedirectPromise` resolves with an `AuthenticationResult` containing
   `idToken`, `accessToken`, `account`, and `idTokenClaims`.
3. The `oid` and `name` claims are present in `idTokenClaims`.
4. (Optional) If your app needs the user's email from Microsoft Graph, verify
   the access token works:

```javascript
const res = await fetch(
  `https://graph.microsoft.com/v1.0/users/${authResult.account.localAccountId}?$select=mail`,
  { headers: { Authorization: `Bearer ${authResult.accessToken}` } },
);
console.log(await res.json()); // { mail: "user@example.com", ... }
```

This requires the `User.Read` scope and appropriate API permissions on the app
registration.

## Running the test application locally

Make sure the environment variables from step 2 are exported in your shell
(e.g. via `direnv allow` if you use an `.envrc`), then install dependencies:

```bash
npm install
npx playwright install chromium
```

### Dev server (manual testing)

```bash
npm run dev
```

Open `http://localhost:8080` in a browser to interact with the OAuth flow
manually.

### E2E tests (Playwright)

```bash
npm test
```

This starts Vite automatically, runs the Playwright test suite against it, and
shuts it down when finished. The test suite reads `TEST_USER_EMAIL` and
`TEST_USER_PASSWORD` from the environment.

