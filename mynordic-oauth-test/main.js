import {
  InteractionRequiredAuthError,
  PublicClientApplication,
} from "@azure/msal-browser";
import { loginRequest, msalConfig, ssoSilentRequest } from "./msalConfig.js";

const msal = new PublicClientApplication(msalConfig);

const SSO_QUERY_KEYS = ["login_hint", "loginHint", "sid"];
const SSO_HINTS_STORAGE_KEY = "nrfcloud_sso_hints";
const SSO_METHOD_STORAGE_KEY = "nrfcloud_sso_method";

/** login_hint from myNordic redirect (e.g. ?login_hint=user@example.com). */
function ssoHintsFromLocation() {
  const params = new URLSearchParams(window.location.search);
  const loginHint =
    params.get("login_hint")?.trim() || params.get("loginHint")?.trim() || null;
  const sid = params.get("sid")?.trim() || null;
  return { loginHint, sid };
}

function persistSsoHints(hints) {
  if (hints.loginHint || hints.sid) {
    sessionStorage.setItem(SSO_HINTS_STORAGE_KEY, JSON.stringify(hints));
  }
}

/** Keep login_hint across CIAM redirects (query params are stripped on error responses). */
function loadSsoHints() {
  const fromUrl = ssoHintsFromLocation();
  if (fromUrl.loginHint || fromUrl.sid) {
    persistSsoHints(fromUrl);
    return fromUrl;
  }
  try {
    const stored = sessionStorage.getItem(SSO_HINTS_STORAGE_KEY);
    return stored ? JSON.parse(stored) : { loginHint: null, sid: null };
  } catch {
    return { loginHint: null, sid: null };
  }
}

function clearSsoQueryParams() {
  const url = new URL(window.location.href);
  let changed = false;
  for (const key of SSO_QUERY_KEYS) {
    if (url.searchParams.has(key)) {
      url.searchParams.delete(key);
      changed = true;
    }
  }
  if (changed) {
    window.history.replaceState({}, document.title, url.pathname + url.search);
  }
}

function buildSilentSsoRequest({ loginHint, sid }) {
  const request = { ...ssoSilentRequest, ...loginRequest, loginHint };
  if (sid) request.sid = sid;
  return request;
}

function buildInteractiveRequest({ loginHint, sid }) {
  const request = { ...loginRequest };
  if (loginHint) request.loginHint = loginHint;
  if (sid) request.sid = sid;
  return request;
}

/** MSAL AuthError fields are often non-enumerable — plain console.log(err) looks empty. */
function msalErrorDetails(err) {
  if (err == null) return { note: "no error object" };
  return {
    type: err.constructor?.name ?? typeof err,
    name: err.name,
    message: err.message,
    errorCode: err.errorCode,
    errorMessage: err.errorMessage,
    subError: err.subError,
    correlationId: err.correlationId,
    stack: err.stack,
  };
}

function formatMsalError(err) {
  const d = msalErrorDetails(err);
  const parts = [d.errorCode, d.errorMessage, d.subError, d.message].filter(Boolean);
  return parts.length > 0 ? parts.join(" — ") : "unknown error";
}

/** Multi-line text for alert / confirm (objects stringify as [object Object] otherwise). */
function formatMsalErrorAlert(err) {
  const d = msalErrorDetails(err);
  return [
    `type: ${d.type ?? "unknown"}`,
    d.errorCode ? `errorCode: ${d.errorCode}` : null,
    d.subError ? `subError: ${d.subError}` : null,
    d.errorMessage ? `errorMessage: ${d.errorMessage}` : null,
    d.message ? `message: ${d.message}` : null,
    d.correlationId ? `correlationId: ${d.correlationId}` : null,
  ]
    .filter(Boolean)
    .join("\n");
}

function logMsalError(label, err) {
  console.error(label, msalErrorDetails(err));
}

/** Set true to alert/debugger before loginRedirect fallbacks. */
const DEBUG_SSO_PAUSE = false;

async function pauseBeforeLoginRedirect(err, fallbackLabel) {
  const details = formatMsalErrorAlert(err);
  const summary = formatMsalError(err);

  setStatus(`ssoSilent failed — paused before ${fallbackLabel}\n${summary}`);
  alert(`ssoSilent failed\n\n${details}`);
  debugger; // DevTools open → execution stops here

  return confirm(
    `Continue to ${fallbackLabel}?\n\n${summary}\n\nCancel to stay on this page.`,
  );
}

function isCallbackPage() {
  const path = window.location.pathname.replace(/\/$/, "");
  return path === "/mynordic/from-portal";
}

function authErrorFromUrl() {
  const search = new URLSearchParams(window.location.search);
  const hash = new URLSearchParams(window.location.hash.replace(/^#/, ""));
  const error = search.get("error") ?? hash.get("error");
  if (!error) return null;
  return {
    error,
    description: search.get("error_description") ?? hash.get("error_description"),
  };
}

/** Silent flows cannot show consent UI — user must sign in interactively once per app. */
function requiresInteractiveSignIn(err) {
  return err instanceof InteractionRequiredAuthError;
}

function urlRequiresInteractiveSignIn() {
  const urlError = authErrorFromUrl();
  if (!urlError) return false;
  return ["consent_required", "interaction_required", "login_required"].includes(
    urlError.error,
  );
}

async function startInteractiveSignIn(hints = ssoHintsFromLocation()) {
  clearAuthErrorFromUrl();
  setStatus("Opening interactive sign-in for nrfCloud…");
  await msal.loginRedirect(
    hints.loginHint ? buildInteractiveRequest(hints) : loginRequest,
  );
}

function clearAuthErrorFromUrl() {
  const url = new URL(window.location.href);
  for (const key of ["error", "error_description", "error_uri", "state"]) {
    url.searchParams.delete(key);
  }
  const hashParams = new URLSearchParams(url.hash.replace(/^#/, ""));
  let hashChanged = false;
  for (const key of ["error", "error_description", "error_uri", "state"]) {
    if (hashParams.has(key)) {
      hashParams.delete(key);
      hashChanged = true;
    }
  }
  if (hashChanged) {
    url.hash = hashParams.toString() ? `#${hashParams}` : "";
  }
  window.history.replaceState({}, document.title, url.pathname + url.search + url.hash);
}

/**
 * Top-level redirect reuses the CIAM session cookie (works when iframe ssoSilent cannot).
 * @see https://learn.microsoft.com/en-us/entra/identity-platform/msal-js-sso#third-party-cookies
 */
async function attemptPromptNoneRedirect({ loginHint, sid }) {
  persistSsoHints({ loginHint, sid });
  sessionStorage.setItem(SSO_METHOD_STORAGE_KEY, "prompt=none");
  setStatus("Continuing myNordic session…");
  await msal.loginRedirect({
    ...buildInteractiveRequest({ loginHint, sid }),
    prompt: "none",
  });
}

/**
 * Cross-app SSO (strict order):
 *   1. ssoSilent
 *   2. loginRedirect + prompt=none (if step 1 fails)
 *   3. interactive loginRedirect (if step 2 fails — handled in init after redirect)
 * @see https://learn.microsoft.com/en-us/entra/identity-platform/msal-js-sso#sso-between-different-apps
 */
async function attemptCrossAppSignIn({ loginHint, sid }) {
  persistSsoHints({ loginHint, sid });

  try {
    sessionStorage.setItem(SSO_METHOD_STORAGE_KEY, "ssoSilent");
    const result = await msal.ssoSilent(buildSilentSsoRequest({ loginHint, sid }));
    clearSsoQueryParams();
    show(result, "silent SSO (ssoSilent)");
    return true;
  } catch (err) {
    logMsalError("ssoSilent failed, trying prompt=none", err);
    if (DEBUG_SSO_PAUSE && !(await pauseBeforeLoginRedirect(err, "prompt=none"))) {
      return false;
    }
  }

  await attemptPromptNoneRedirect({ loginHint, sid });
  return true;
}

async function init() {
  await msal.initialize();

  // Complete loginRedirect fallback (PKCE is generated in this app).
  let authResult = null;
  try {
    authResult = await msal.handleRedirectPromise();
  } catch (err) {
    logMsalError("handleRedirectPromise failed", err);
    if (isCallbackPage() && requiresInteractiveSignIn(err)) {
      await startInteractiveSignIn(loadSsoHints());
      return;
    }
    showError(err);
    return;
  }

  if (authResult) {
    clearSsoQueryParams();
    clearAuthErrorFromUrl();
    const method = sessionStorage.getItem(SSO_METHOD_STORAGE_KEY);
    sessionStorage.removeItem(SSO_METHOD_STORAGE_KEY);
    const note =
      method === "ssoSilent"
        ? "silent SSO (ssoSilent)"
        : method === "prompt=none"
          ? "silent SSO (prompt=none)"
          : "sign-in";
    show(authResult, note);
    return;
  }

  // Step 3: loginRedirect (prompt=none) failed — interactive sign-in only (no second ssoSilent).
  if (isCallbackPage() && urlRequiresInteractiveSignIn()) {
    const hints = loadSsoHints();
    logMsalError("loginRedirect (prompt=none) failed", {
      errorCode: authErrorFromUrl()?.error,
      errorMessage: authErrorFromUrl()?.description,
    });
    clearAuthErrorFromUrl();
    await startInteractiveSignIn(hints);
    return;
  }

  const accounts = msal.getAllAccounts();
  if (accounts.length > 0) {
    show({ account: accounts[0] }, "restored from MSAL cache");
    return;
  }

  // myNordic redirects here with ?login_hint= after CIAM login.
  if (isCallbackPage()) {
    const hints = loadSsoHints();
    if (!hints.loginHint) {
      await startInteractiveSignIn(hints);
      return;
    }
    await attemptCrossAppSignIn(hints);
    return;
  }

  setStatus("Not authenticated. Open from myNordic or sign in.");
}

function show(result, note) {
  const payload = note ? { ...result, note } : result;
  document.getElementById("output").textContent = JSON.stringify(payload, null, 2);
  document.getElementById("login").hidden = true;
  document.getElementById("logout").hidden = false;
}

function showError(err) {
  setStatus(`Authentication failed: ${err.message ?? err}`);
  document.getElementById("login").hidden = false;
  document.getElementById("logout").hidden = true;
}

function setStatus(message) {
  document.getElementById("output").textContent = message;
}

document.getElementById("login").addEventListener("click", () => {
  msal.loginRedirect(loginRequest);
});

/**
 * CIAM shows an empty "Pick an account" screen when logout has no account/logout_hint
 * (common after myNordic portal SSO). See README → Logout after portal SSO.
 */
function buildLogoutRequest() {
  const account = msal.getAllAccounts()[0];
  const { loginHint } = loadSsoHints();
  const logoutHint =
    account?.idTokenClaims?.login_hint ?? loginHint ?? account?.username ?? undefined;

  const request = {
    postLogoutRedirectUri: msalConfig.auth.postLogoutRedirectUri,
  };
  if (account) request.account = account;
  if (logoutHint) request.logoutHint = logoutHint;
  return request;
}

document.getElementById("logout").addEventListener("click", async () => {
  const request = buildLogoutRequest();
  if (!request.account && !request.logoutHint) {
    console.warn(
      "Logout without account or logoutHint — CIAM may show an empty account picker.",
    );
  }
  await msal.logoutRedirect(request);
});

init();
