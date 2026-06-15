import { PublicClientApplication } from "@azure/msal-browser";
import { msalConfig, loginRequest } from "./msalConfig.js";

const msal = new PublicClientApplication(msalConfig);

async function init() {
  await msal.initialize();

  const authResult = await msal.handleRedirectPromise();

  if (authResult) {
    show(authResult);
    return;
  }

  const accounts = msal.getAllAccounts();
  if (accounts.length > 0) {
    show({ account: accounts[0], note: "restored from session cache" });
  }
}

const portalSignUpUrl = new URL(
  "https://mynordic.nordicsemi.com/en/sign-up/create-account",
);
portalSignUpUrl.searchParams.set("redirect_uri", msalConfig.auth.redirectUri);

function show(result) {
  document.getElementById("output").textContent = JSON.stringify(
    result,
    null,
    2,
  );
  document.getElementById("login").hidden = true;
  document.getElementById("portal-login").hidden = true;
  document.getElementById("logout").hidden = false;
}

document.getElementById("login").addEventListener("click", () => {
  msal.loginRedirect(loginRequest);
});

document.getElementById("portal-login").addEventListener("click", () => {
  window.location.href = portalSignUpUrl.toString();
});

document.getElementById("logout").addEventListener("click", async () => {
  const account = msal.getAllAccounts()[0];
  const logoutHint =
    account?.idTokenClaims?.login_hint ?? account?.username ?? undefined;

  const request = {
    postLogoutRedirectUri: msalConfig.auth.postLogoutRedirectUri,
  };
  if (account) request.account = account;
  if (logoutHint) request.logoutHint = logoutHint;

  await msal.logoutRedirect(request);
});

init();
