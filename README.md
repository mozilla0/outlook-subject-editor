# Subject Editor für Outlook

[English version](README.en.md)

Outlook-Add-In, mit dem Benutzer den Betreff einer **empfangenen** E-Mail nachträglich ändern können – etwa um eine Ticketnummer zu ergänzen, «AW: WG: AW:»-Ketten aufzuräumen oder eine Nachricht besser auffindbar zu machen.

Die Änderung läuft über **Microsoft Graph** und wird direkt im Postfach gespeichert. Damit ist sie in allen Outlook-Clients, in der Suche und in Regeln sichtbar. EWS wird nicht verwendet.

## Funktionen

- Bereich «Betreff ändern» in der Leseansicht von Outlook (Web, neues Outlook, klassisches Outlook für Windows, Mac)
- Anheftbarer Bereich: beim Wechsel zu einer anderen Nachricht wird der Betreff automatisch geladen
- Freigegebene und stellvertretend geöffnete Postfächer
- Anmeldung per **Nested App Authentication (NAA)** – kein separates Login, Outlook stellt das Token aus; Fallback über einen Office-Dialog für ältere Clients
- Oberfläche auf Deutsch (Schweizer Rechtschreibung) und Englisch, hell und dunkel
- Bereitstellung pro Tenant mit einem PowerShell-Skript (App-Registrierung, Zustimmung, Manifest, optional Verteilung)

## Funktionsweise

```
Outlook ──Office.js──▶ Task Pane (statisches HTML/JS)
                          │  MSAL (NAA) ─▶ Entra ID (Single-Tenant-App)
                          └─ PATCH /me/messages/{id} ─▶ Microsoft Graph
```

Graph erlaubt das Ändern der Eigenschaft `subject` nur bei Entwürfen. Das Add-In schreibt deshalb die MAPI-Eigenschaft `PidTagSubject` (`String 0x0037`) über `singleValueExtendedProperties`. Danach liest es den gespeicherten Betreff zurück und zeigt ihn an.

> Outlook aktualisiert die Anzeige im Lesebereich je nach Client erst, wenn die Nachricht neu geöffnet oder die Liste aktualisiert wird. Gespeichert ist die Änderung sofort.

## Voraussetzungen

| Was | Wozu |
| --- | --- |
| Exchange Online (Microsoft 365) | Postfächer, Graph |
| Statisches HTTPS-Hosting | Ausliefern von `dist/` (z. B. Azure Static Web Apps, Azure Storage Static Website, GitHub Pages, IIS, nginx) |
| Node.js 20 oder neuer | Build |
| PowerShell 7.2 oder neuer | Bereitstellungsskript |
| Module `Microsoft.Graph.Applications`, `Microsoft.Graph.Identity.SignIns` | App-Registrierung und Zustimmung |
| Modul `ExchangeOnlineManagement` (optional) | Verteilung mit `-DeployToExchange` |
| Rollen *Privileged Role Administrator* (für die Zustimmung) und *Exchange Administrator* | Einrichtung |

```powershell
Install-Module Microsoft.Graph.Applications, Microsoft.Graph.Identity.SignIns, ExchangeOnlineManagement -Scope CurrentUser
```

## Einrichtung

### 1. Build

```bash
npm install
npm run build
```

Das Ergebnis liegt in `dist/`.

### 2. Tenant einrichten

```powershell
./scripts/Deploy-SubjectEditor.ps1 `
    -TenantId contoso.onmicrosoft.com `
    -BaseUrl  https://addins.contoso.ch/subject-editor
```

Das Skript

1. erstellt oder aktualisiert die App-Registrierung «Subject Editor» (Single-Tenant, Plattform *Single-page application*) mit den Redirect-URIs `brk-multihub://addins.contoso.ch` (NAA) und `…/dialog.html` (Fallback),
2. erteilt die Administratorzustimmung für `User.Read`, `Mail.ReadWrite`, `Mail.ReadWrite.Shared`, `openid`, `profile` und `offline_access`,
3. schreibt `dist/config.json` (Client- und Tenant-ID) und `dist/manifest.xml` (mit der Hosting-URL).

Es ist idempotent und kann jederzeit erneut ausgeführt werden.

| Parameter | Bedeutung |
| --- | --- |
| `-DeployToExchange` | Verteilt das Add-In organisationsweit über Exchange Online. Eine bestehende Version wird ersetzt. |
| `-ProviderName` | Anbietername im Manifest (Standard: «Subject Editor») |
| `-SupportUrl` | Support-Link im Manifest (Standard: Ursprung der `BaseUrl`) |
| `-ClientId` + `-SkipAppRegistration` | Bestehende App-Registrierung verwenden, Entra ID nicht anfassen |
| `-WhatIf` | Zeigt nur an, was geändert würde |

### 3. Hochladen

Den gesamten Inhalt von `dist/` nach `BaseUrl` hochladen. `config.json` sollte ohne Caching ausgeliefert werden.

### 4. Verteilen

Entweder mit `-DeployToExchange` (siehe oben) oder im Microsoft 365 Admin Center unter **Einstellungen › Integrierte Apps › Benutzerdefinierte Apps hochladen** die Datei `dist/manifest.xml` hochladen und Benutzer oder Gruppen zuweisen. Bis das Add-In bei allen erscheint, kann es einige Stunden dauern.

Zum Testen ohne Verteilung: In Outlook im Web unter <https://aka.ms/olksideload> → *Meine Add-Ins* → *Benutzerdefiniertes Add-In hinzufügen* → *Aus Datei* die `manifest.xml` hochladen.

## Update

1. Version in `package.json` erhöhen (wird zur Manifest-Version `x.y.z.0`).
2. `npm run build`
3. `Deploy-SubjectEditor.ps1` erneut ausführen (mit `-DeployToExchange`, falls so verteilt).
4. `dist/` hochladen.

Reine Code-Änderungen sind aktiv, sobald die Dateien auf dem Webserver liegen. Eine neue Manifest-Version braucht es nur, wenn sich Manifest-Inhalte ändern.

## Entfernen

```powershell
./scripts/Remove-SubjectEditor.ps1 -TenantId contoso.onmicrosoft.com -RemoveAppRegistration
```

## Berechtigungen und Datenschutz

- Nur **delegierte** Berechtigungen: Das Add-In kann ausschliesslich Nachrichten ändern, auf die der angemeldete Benutzer selbst Schreibzugriff hat. Bei freigegebenen Postfächern braucht es Vollzugriff.
- Es gibt kein Backend. Der Browser des Benutzers spricht direkt mit Microsoft Graph; der Webserver liefert nur statische Dateien aus.
- Es werden keine Inhalte gespeichert oder protokolliert.

## Fehlerbehebung

| Meldung | Ursache und Lösung |
| --- | --- |
| «Das Add-In ist noch nicht eingerichtet» | `config.json` fehlt auf dem Webserver oder enthält keine Client-ID. Skript ausführen und `dist/` erneut hochladen. |
| AADSTS50011 (Redirect URI mismatch) | Die `BaseUrl` beim Skriptaufruf stimmt nicht mit der tatsächlichen Hosting-URL überein. |
| AADSTS65001 (Zustimmung fehlt) | Skript ohne `-SkipAppRegistration` erneut ausführen oder Zustimmung in Entra ID erteilen. |
| «Keine Berechtigung, diese Nachricht zu ändern» | Bei freigegebenen Postfächern fehlt Vollzugriff. |
| Anmeldefenster öffnet sich nicht | Client ohne NAA und blockierte Pop-ups. Pop-ups für Outlook zulassen. |

## Entwicklung

```bash
npm run watch      # baut bei Änderungen neu
```

Für lokales Testen `dist/` über HTTPS ausliefern (z. B. `npx http-server dist -S -C cert.pem -K key.pem -p 3000`) und das Skript mit `-BaseUrl https://localhost:3000` ausführen.

```
src/
  taskpane.html/.css/.js   Oberfläche und Ablauf
  auth.js                  Token über NAA, Fallback über Office-Dialog
  dialog.html/.js          Anmeldedialog für Clients ohne NAA
  graph.js                 Graph-Aufrufe (Betreff lesen/schreiben)
  config.js                Laufzeitkonfiguration und Scopes
  i18n.js                  Texte Deutsch/Englisch
public/assets/             Icons
manifest.template.xml      Manifest-Vorlage mit Platzhaltern
scripts/                   Bereitstellung und Entfernen
```

## Lizenz

[MIT](LICENSE) © Matthias Hummler
