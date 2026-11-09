// Token acquisition for the task pane.
// Primary path: Nested App Authentication (NAA) via MSAL, which uses the Outlook host's signed-in account.
// Fallback for clients without NAA: an Office dialog that runs a regular MSAL redirect flow (see dialog.js).
import { createNestablePublicClientApplication } from "@azure/msal-browser";
import { SCOPES } from "./config.js";

let pca = null;
let useNaa = false;
let cachedToken = null; // { accessToken, expiresOn } for the dialog fallback

export async function initAuth(config) {
  useNaa = Office.context.requirements.isSetSupported("NestedAppAuth", "1.1");
  if (useNaa) {
    pca = await createNestablePublicClientApplication({
      auth: { clientId: config.clientId, authority: config.authority },
      system: { loggerOptions: { piiLoggingEnabled: false } },
    });
  }
}

export async function getAccessToken() {
  if (useNaa) {
    const request = { scopes: SCOPES };
    try {
      const result = await pca.acquireTokenSilent(request);
      return result.accessToken;
    } catch {
      // Silent acquisition fails on first use or when consent is missing; ask interactively.
      try {
        const result = await pca.acquireTokenPopup(request);
        return result.accessToken;
      } catch (popupError) {
        console.warn("NAA popup failed, falling back to the Office dialog.", popupError);
      }
    }
  }
  return getTokenViaDialog();
}

function getTokenViaDialog() {
  if (cachedToken && cachedToken.expiresOn - Date.now() > 5 * 60 * 1000) {
    return Promise.resolve(cachedToken.accessToken);
  }
  const dialogUrl = new URL("dialog.html", window.location.href).toString();

  return new Promise((resolve, reject) => {
    Office.context.ui.displayDialogAsync(dialogUrl, { height: 60, width: 30 }, (asyncResult) => {
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        reject(new Error("dialog-blocked"));
        return;
      }
      const dialog = asyncResult.value;

      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        dialog.close();
        let message;
        try {
          message = JSON.parse(arg.message);
        } catch {
          reject(new Error("Invalid response from the sign-in dialog."));
          return;
        }
        if (message.status === "success") {
          cachedToken = { accessToken: message.accessToken, expiresOn: message.expiresOn };
          resolve(message.accessToken);
        } else {
          reject(new Error(message.error || "Sign-in was cancelled."));
        }
      });

      dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
        reject(new Error("Sign-in was cancelled."));
      });
    });
  });
}
