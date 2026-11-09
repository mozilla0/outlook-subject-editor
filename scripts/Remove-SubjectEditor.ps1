#Requires -Version 7.2
<#
.SYNOPSIS
    Entfernt den Subject Editor aus einem Tenant.

.DESCRIPTION
    Entfernt das organisationsweit verteilte Add-In aus Exchange Online und löscht auf Wunsch
    (-RemoveAppRegistration) auch die App-Registrierung samt Administratorzustimmung.

.EXAMPLE
    ./scripts/Remove-SubjectEditor.ps1 -TenantId contoso.onmicrosoft.com -RemoveAppRegistration
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory)]
    [string] $TenantId,

    [string] $AddInId = 'e0d00464-913a-49dc-8dad-9ee6a2ab6ce0',

    [string] $AppDisplayName = 'Subject Editor',

    [switch] $RemoveAppRegistration
)

$ErrorActionPreference = 'Stop'

Connect-ExchangeOnline -ShowBanner:$false
$existing = Get-App -OrganizationApp | Where-Object AppId -EQ $AddInId
if ($existing) {
    if ($PSCmdlet.ShouldProcess($existing.DisplayName, 'Add-In aus der Organisation entfernen')) {
        Remove-App -Identity $AddInId -OrganizationApp -Confirm:$false
        Write-Host 'Add-In entfernt.'
    }
} else {
    Write-Host 'Kein organisationsweit installiertes Add-In gefunden.'
}

if ($RemoveAppRegistration) {
    Connect-MgGraph -TenantId $TenantId -NoWelcome -Scopes 'Application.ReadWrite.All'
    $app = Get-MgApplication -Filter "displayName eq '$($AppDisplayName.Replace("'", "''"))'" | Select-Object -First 1
    if ($app -and $PSCmdlet.ShouldProcess($app.DisplayName, 'App-Registrierung löschen')) {
        # Löscht auch den Service Principal und damit die erteilte Zustimmung.
        Remove-MgApplication -ApplicationId $app.Id
        Write-Host "App-Registrierung $($app.AppId) gelöscht."
    }
}
