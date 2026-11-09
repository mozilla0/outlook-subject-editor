#Requires -Version 7.2
<#
.SYNOPSIS
    Richtet den Subject Editor für einen Microsoft-365-Tenant ein.

.DESCRIPTION
    1. Erstellt oder aktualisiert die App-Registrierung in Entra ID (Single-Tenant, SPA,
       Redirect-URIs für Nested App Authentication und den Dialog-Fallback).
    2. Erteilt die Administratorzustimmung für die delegierten Graph-Berechtigungen
       (User.Read, Mail.ReadWrite, Mail.ReadWrite.Shared, openid, profile, offline_access).
    3. Schreibt config.json und manifest.xml für die angegebene Hosting-URL in den Ausgabeordner.
    4. Optional (-DeployToExchange): verteilt das Add-In organisationsweit über Exchange Online.
       Ist es bereits installiert, wird es durch die neue Manifest-Version ersetzt.

    Das Skript ist idempotent und kann für Updates erneut ausgeführt werden.

.PARAMETER TenantId
    Tenant-ID oder primäre Domain des Ziel-Tenants (z. B. contoso.onmicrosoft.com).

.PARAMETER BaseUrl
    HTTPS-URL, unter der der Inhalt von dist/ gehostet wird, z. B. https://addins.contoso.ch/subject-editor

.PARAMETER ClientId
    Bestehende App-Registrierung verwenden. Zusammen mit -SkipAppRegistration wird Entra ID nicht verändert.

.EXAMPLE
    ./scripts/Deploy-SubjectEditor.ps1 -TenantId contoso.onmicrosoft.com -BaseUrl https://addins.contoso.ch/subject-editor

.EXAMPLE
    ./scripts/Deploy-SubjectEditor.ps1 -TenantId contoso.onmicrosoft.com -BaseUrl https://addins.contoso.ch/subject-editor -DeployToExchange
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string] $TenantId,

    [Parameter(Mandatory)]
    [ValidatePattern('^https://')]
    [string] $BaseUrl,

    [string] $AppDisplayName = 'Subject Editor',

    [string] $ProviderName = 'Subject Editor',

    [string] $SupportUrl,

    # Fixe Add-In-ID aus dem Repository. Nur ändern, wenn eine zweite, parallele Installation nötig ist.
    [ValidatePattern('^[0-9a-fA-F-]{36}$')]
    [string] $AddInId = 'e0d00464-913a-49dc-8dad-9ee6a2ab6ce0',

    [string] $OutputPath,

    [string] $ClientId,

    [switch] $SkipAppRegistration,

    [switch] $DeployToExchange
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot   = Split-Path -Parent $PSScriptRoot
$OutputPath = if ($OutputPath) { $OutputPath } else { Join-Path $repoRoot 'dist' }
$BaseUrl    = $BaseUrl.TrimEnd('/')
$baseUri    = [Uri] $BaseUrl
$origin     = $baseUri.GetLeftPart([UriPartial]::Authority)
$SupportUrl = if ($SupportUrl) { $SupportUrl } else { $origin }

$graphAppId = '00000003-0000-0000-c000-000000000046'
$scopeNames = @('User.Read', 'Mail.ReadWrite', 'Mail.ReadWrite.Shared', 'openid', 'profile', 'offline_access')

function Write-Step([string] $Text) { Write-Host "`n==> $Text" -ForegroundColor Cyan }

function Assert-Module([string] $Name) {
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        throw "Das PowerShell-Modul '$Name' fehlt. Installation: Install-Module $Name -Scope CurrentUser"
    }
}

function Get-AddInVersion {
    $package = Get-Content (Join-Path $repoRoot 'package.json') -Raw | ConvertFrom-Json
    $parts = @($package.version -split '[.-]' | Select-Object -First 3)
    while ($parts.Count -lt 3) { $parts += '0' }
    return ($parts + '0') -join '.'
}

# ---------------------------------------------------------------------------
# 1 + 2: App-Registrierung und Administratorzustimmung
# ---------------------------------------------------------------------------
if (-not $SkipAppRegistration) {
    Assert-Module 'Microsoft.Graph.Applications'
    Assert-Module 'Microsoft.Graph.Identity.SignIns'

    Write-Step "Verbindung zu Microsoft Graph ($TenantId)"
    Connect-MgGraph -TenantId $TenantId -NoWelcome -Scopes @(
        'Application.ReadWrite.All',
        'DelegatedPermissionGrant.ReadWrite.All'
    )
    $tenantGuid = (Get-MgContext).TenantId

    $graphSp = Get-MgServicePrincipal -Filter "appId eq '$graphAppId'"
    $resourceAccess = foreach ($name in $scopeNames) {
        $scope = $graphSp.Oauth2PermissionScopes | Where-Object Value -EQ $name
        if (-not $scope) { throw "Graph-Berechtigung '$name' wurde nicht gefunden." }
        @{ Id = $scope.Id; Type = 'Scope' }
    }

    # brk-multihub:// ist die Redirect-URI für Nested App Authentication (Outlook-Host als Broker).
    # dialog.html ist der Fallback für Outlook-Versionen ohne NAA.
    $redirectUris = @("brk-multihub://$($baseUri.Authority)", "$BaseUrl/dialog.html")

    $appSettings = @{
        SignInAudience         = 'AzureADMyOrg'
        Spa                    = @{ RedirectUris = $redirectUris }
        RequiredResourceAccess = @(@{ ResourceAppId = $graphAppId; ResourceAccess = @($resourceAccess) })
    }

    Write-Step 'App-Registrierung'
    $app = if ($ClientId) {
        Get-MgApplication -Filter "appId eq '$ClientId'"
    } else {
        Get-MgApplication -Filter "displayName eq '$($AppDisplayName.Replace("'", "''"))'" | Select-Object -First 1
    }

    if ($app) {
        if ($PSCmdlet.ShouldProcess($app.DisplayName, 'App-Registrierung aktualisieren')) {
            Update-MgApplication -ApplicationId $app.Id -BodyParameter $appSettings
        }
        Write-Host "Aktualisiert: $($app.DisplayName) ($($app.AppId))"
    } elseif ($PSCmdlet.ShouldProcess($AppDisplayName, 'App-Registrierung erstellen')) {
        $app = New-MgApplication -BodyParameter ($appSettings + @{ DisplayName = $AppDisplayName })
        Write-Host "Erstellt: $AppDisplayName ($($app.AppId))"
    }

    if ($app) {
        $ClientId = $app.AppId

        $sp = Get-MgServicePrincipal -Filter "appId eq '$ClientId'"
        if (-not $sp) {
            # Neu erstellte Apps sind manchmal erst nach ein paar Sekunden sichtbar.
            for ($i = 0; $i -lt 6 -and -not $sp; $i++) {
                try { $sp = New-MgServicePrincipal -AppId $ClientId } catch { Start-Sleep -Seconds 5 }
            }
            if (-not $sp) { throw 'Der Service Principal konnte nicht erstellt werden.' }
        }

        Write-Step 'Administratorzustimmung (alle Benutzer)'
        $scopeString = $scopeNames -join ' '
        $grant = Get-MgOauth2PermissionGrant -Filter "clientId eq '$($sp.Id)'" -All |
            Where-Object { $_.ConsentType -eq 'AllPrincipals' -and $_.ResourceId -eq $graphSp.Id }

        if ($grant) {
            if ($PSCmdlet.ShouldProcess('Graph-Berechtigungen', 'Zustimmung aktualisieren')) {
                Update-MgOauth2PermissionGrant -OAuth2PermissionGrantId $grant.Id -Scope $scopeString
            }
        } elseif ($PSCmdlet.ShouldProcess('Graph-Berechtigungen', 'Zustimmung erteilen')) {
            New-MgOauth2PermissionGrant -ClientId $sp.Id -ConsentType 'AllPrincipals' -ResourceId $graphSp.Id -Scope $scopeString | Out-Null
        }
        Write-Host "Zugestimmt: $scopeString"
    }
} else {
    if (-not $ClientId) { throw '-SkipAppRegistration braucht -ClientId.' }
    $tenantGuid = $TenantId
}

# ---------------------------------------------------------------------------
# 3: config.json und manifest.xml
# ---------------------------------------------------------------------------
Write-Step "Konfiguration schreiben ($OutputPath)"
if (-not (Test-Path (Join-Path $OutputPath 'taskpane.html'))) {
    Write-Warning "In '$OutputPath' liegt noch kein Build. Vor dem Hochladen 'npm run build' ausführen (und dieses Skript danach erneut)."
}
New-Item -ItemType Directory -Force -Path $OutputPath | Out-Null

$config = [ordered] @{ clientId = $ClientId; tenantId = "$tenantGuid" } | ConvertTo-Json
Set-Content -Path (Join-Path $OutputPath 'config.json') -Value $config -Encoding utf8NoBOM

$tokens = [ordered] @{
    '{{ADDIN_ID}}'      = $AddInId
    '{{VERSION}}'       = Get-AddInVersion
    '{{PROVIDER_NAME}}' = $ProviderName
    '{{SUPPORT_URL}}'   = $SupportUrl
    '{{ORIGIN}}'        = $origin
    '{{BASE_URL}}'      = $BaseUrl
}
$manifest = Get-Content (Join-Path $repoRoot 'manifest.template.xml') -Raw
foreach ($token in $tokens.GetEnumerator()) {
    $manifest = $manifest.Replace($token.Key, [System.Security.SecurityElement]::Escape([string] $token.Value))
}

[xml] $manifest | Out-Null   # wirft eine Ausnahme bei ungültigem XML
$manifestPath = Join-Path $OutputPath 'manifest.xml'
Set-Content -Path $manifestPath -Value $manifest -Encoding utf8NoBOM

Write-Host "config.json   -> Client-ID $ClientId"
Write-Host "manifest.xml  -> Version $(Get-AddInVersion), $BaseUrl"

# ---------------------------------------------------------------------------
# 4: Verteilung über Exchange Online (optional)
# ---------------------------------------------------------------------------
if ($DeployToExchange) {
    Assert-Module 'ExchangeOnlineManagement'

    Write-Step 'Verteilung über Exchange Online'
    Connect-ExchangeOnline -ShowBanner:$false
    $bytes = [System.IO.File]::ReadAllBytes($manifestPath)

    $existing = Get-App -OrganizationApp | Where-Object AppId -EQ $AddInId
    if ($existing -and $PSCmdlet.ShouldProcess($existing.DisplayName, 'Bestehende Version entfernen')) {
        Remove-App -Identity $AddInId -OrganizationApp -Confirm:$false
    }
    if ($PSCmdlet.ShouldProcess('Organisation', 'Add-In installieren')) {
        New-App -OrganizationApp -FileData $bytes -DefaultStateForUser Enabled -ProvidedTo Everyone | Out-Null
        Write-Host 'Installiert. Bis das Add-In bei allen Benutzern erscheint, kann es einige Stunden dauern.'
    }
}

Write-Step 'Fertig'
Write-Host "Inhalt von '$OutputPath' nach $BaseUrl hochladen."
if (-not $DeployToExchange) {
    Write-Host "Danach manifest.xml im Microsoft 365 Admin Center unter «Integrierte Apps» hochladen oder das Skript mit -DeployToExchange ausführen."
}
