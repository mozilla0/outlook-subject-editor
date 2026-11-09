// Loads the tenant-specific runtime configuration written by scripts/Deploy-SubjectEditor.ps1.
const EMPTY_GUID = "00000000-0000-0000-0000-000000000000";

export async function loadConfig() {
  const url = new URL("config.json", window.location.href);
  const response = await fetch(url, { cache: "no-store" });
  if (!response.ok) {
    throw new Error("config-missing");
  }
  const config = await response.json();
  if (!config.clientId || config.clientId === EMPTY_GUID) {
    throw new Error("config-missing");
  }
  return {
    clientId: config.clientId,
    // Single-tenant app: always authenticate against the customer's own tenant.
    authority: `https://login.microsoftonline.com/${config.tenantId || "organizations"}`,
  };
}

// Delegated Microsoft Graph permissions. Mail.ReadWrite.Shared covers shared and delegated mailboxes.
export const SCOPES = ["Mail.ReadWrite", "Mail.ReadWrite.Shared"];
