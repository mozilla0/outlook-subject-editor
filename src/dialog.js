// Sign-in dialog for Outlook clients without Nested App Authentication.
// Runs a standard MSAL redirect flow inside an Office dialog and passes the token back to the task pane.
import { PublicClientApplication } from "@azure/msal-browser";
import { loadConfig, SCOPES } from "./config.js";

function reply(payload) {
  Office.context.ui.messageParent(JSON.stringify(payload));
}

Office.onReady(async () => {
  try {
    const config = await loadConfig();
    const redirectUri = new URL("dialog.html", window.location.href).toString().split("?")[0];
    const pca = new PublicClientApplication({
      auth: { clientId: config.clientId, authority: config.authority, redirectUri },
      cache: { cacheLocation: "localStorage" },
    });
    await pca.initialize();

    const redirectResult = await pca.handleRedirectPromise();
    if (redirectResult) {
      reply({ status: "success", accessToken: redirectResult.accessToken, expiresOn: redirectResult.expiresOn?.getTime() ?? 0 });
      return;
    }

    const account = pca.getAllAccounts()[0];
    if (account) {
      try {
        const silent = await pca.acquireTokenSilent({ scopes: SCOPES, account });
        reply({ status: "success", accessToken: silent.accessToken, expiresOn: silent.expiresOn?.getTime() ?? 0 });
        return;
      } catch {
        /* fall through to the interactive redirect */
      }
    }

    await pca.acquireTokenRedirect({ scopes: SCOPES, account });
  } catch (error) {
    reply({ status: "error", error: error.message });
  }
});
