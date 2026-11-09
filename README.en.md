# Subject Editor for Outlook

[Deutsche Version](README.md)

Outlook add-in that lets users change the subject of a **received** e-mail – for example to add a ticket number, clean up "RE: FW: RE:" chains or make a message easier to find.

The change is made through **Microsoft Graph** and stored in the mailbox, so it is visible in every Outlook client, in search and in rules. EWS is not used.

## Features

- "Change subject" pane in the Outlook reading view (web, new Outlook, classic Outlook for Windows, Mac)
- Pinnable pane: selecting another message loads its subject automatically
- Shared and delegated mailboxes
- Sign-in via **Nested App Authentication (NAA)** – no separate login, Outlook issues the token; an Office dialog serves as fallback for older clients
- UI in German (Swiss spelling) and English, light and dark
- Per-tenant setup with one PowerShell script (app registration, consent, manifest, optional deployment)

## How it works

```
Outlook ──Office.js──▶ Task pane (static HTML/JS)
                          │  MSAL (NAA) ─▶ Entra ID (single-tenant app)
                          └─ PATCH /me/messages/{id} ─▶ Microsoft Graph
```

Graph only allows changing the `subject` property on drafts. The add-in therefore writes the MAPI property `PidTagSubject` (`String 0x0037`) through `singleValueExtendedProperties`, then reads the stored subject back and displays it.

> Depending on the client, Outlook refreshes the reading pane only after the message is reopened or the list is refreshed. The change itself is saved immediately.

## Requirements

| What | Why |
| --- | --- |
| Exchange Online (Microsoft 365) | Mailboxes, Graph |
| Static HTTPS hosting | Serving `dist/` (e.g. Azure Static Web Apps, Azure Storage static website, GitHub Pages, IIS, nginx) |
| Node.js 20 or later | Build |
| PowerShell 7.2 or later | Setup script |
| Modules `Microsoft.Graph.Applications`, `Microsoft.Graph.Identity.SignIns` | App registration and consent |
| Module `ExchangeOnlineManagement` (optional) | Deployment with `-DeployToExchange` |
| Roles *Privileged Role Administrator* (for consent) and *Exchange Administrator* | Setup |

## Setup

```bash
npm install
npm run build
```

```powershell
./scripts/Deploy-SubjectEditor.ps1 -TenantId contoso.onmicrosoft.com -BaseUrl https://addins.contoso.com/subject-editor
```

The script creates or updates the single-tenant app registration with the redirect URIs `brk-multihub://<host>` (NAA) and `<BaseUrl>/dialog.html` (fallback), grants admin consent for `User.Read`, `Mail.ReadWrite`, `Mail.ReadWrite.Shared`, `openid`, `profile` and `offline_access`, and writes `dist/config.json` and `dist/manifest.xml`. It is idempotent.

Upload the content of `dist/` to `BaseUrl`, then deploy the add-in either with `-DeployToExchange` or by uploading `dist/manifest.xml` in the Microsoft 365 admin center under **Settings › Integrated apps**.

For testing without deployment, sideload the manifest in Outlook on the web via <https://aka.ms/olksideload>.

## Update

Bump the version in `package.json`, run `npm run build`, run the script again and upload `dist/`.

## Removal

```powershell
./scripts/Remove-SubjectEditor.ps1 -TenantId contoso.onmicrosoft.com -RemoveAppRegistration
```

## Permissions and privacy

Delegated permissions only – the add-in can change only messages the signed-in user has write access to. There is no backend; the browser talks to Microsoft Graph directly and nothing is stored or logged.

## License

[MIT](LICENSE) © Matthias Hummler
