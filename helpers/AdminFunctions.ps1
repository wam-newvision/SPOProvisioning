# --------------------------------------------------------------------
# Include Helper Functions
# --------------------------------------------------------------------
. (Join-Path $PSScriptRoot 'LoggingFunctions.ps1')

# --------------------------------------------------------------------
# Module laden (PnP.PowerShell, Graph-Module, ...)
# --------------------------------------------------------------------
function LoadPnPPSModule {
    param (
        [string]$PnPVersion,  # $PnPVersion = "3.1.0"
        [string]$FktPath  # $FktPath = "C:\Functions\dms-provisioning\Modules"
    )
        
    #if (-not $PnPVersion) {$PnPVersion = "3.1.0"} #$PnPVersion = "1.12.0"
    #if (-not $FktPath) {$FktPath = "C:\Functions\dms-provisioning\CreateTeamSiteW"} 

    $sep = [IO.Path]::PathSeparator
    $env:PSModulePath = "${FktPath}${sep}$env:PSModulePath"
    
    $PnPPath = Join-Path $FktPath "PnP.PowerShell\$PnPVersion\PnP.PowerShell.psd1"

    Log "📦 Import PnP.PowerShell from local module: $PnPVersion"
    Log "📦 Path: $PnPPath"
    
    Import-Module $PnPPath -DisableNameChecking -Global -ErrorAction Stop

#    if ($local) {
#       Log "📦 Import PnP.PowerShell from local module..."
#        Import-Module (Join-Path $FktPath 'modules\PnP.PowerShell\3.1.0\PnP.PowerShell.psd1') -DisableNameChecking -Global -ErrorAction Stop
#    } else {
#        Log "📦 Import PnP.PowerShell from requirements.psd1..."
#        Import-Module "PnP.PowerShell" -DisableNameChecking -Global -ErrorAction Stop
#        #Import-Module (Join-Path $FktPath 'requirements.psd1') -DisableNameChecking -Global
#    }

    Log "📦 PowerShell Version:"
    $PSVersionTable.PSVersion
    Log "📦 PnP.PowerShell Versionen verfügbar:"
    Get-Module -Name PnP.PowerShell -ListAvailable | Select-Object Version, ModuleBase
    Log "📦 PnP.PowerShell Version verwendet:"
    Get-Module -Name PnP.PowerShell | Select-Object Version, ModuleBase
    Log "📦 PnP.PowerShell import done"

}
# --------------------------------------------------------------------

function LoadGraphModule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ModuleName,     # z.B. 'Microsoft.Graph.Teams'
        [Parameter(Mandatory)][string]$FktPath,        # Pfad auf den 'Modules'-Ordner
        [Parameter(Mandatory)][string]$GraphVersion    # z.B. '2.29.1'
    )

    $ErrorActionPreference = 'Stop'

    # Erwartete Struktur: <FktPath>\<ModuleName>\<GraphVersion>\<ModuleName>.psd1
    $psd1 = Join-Path $FktPath (Join-Path $ModuleName (Join-Path $GraphVersion "$ModuleName.psd1"))

    if (-not (Test-Path $psd1)) {
        throw "❌ Modulmanifest nicht gefunden: $psd1  (erwartet unter '$FktPath\$ModuleName\$GraphVersion\')"
    }

    Import-Module $psd1 -Force -ErrorAction Stop
    Log "✅ Modul '$ModuleName' geladen aus $($psd1 | Split-Path -Parent)"
}
# --------------------------------------------------------------------
function Connect-MSGraph {
    param(
        [Parameter(Mandatory=$true)][string]$TenantId,   # Kunden-Tenant ID oder Domain
        [string]$ClientId,
        [string]$PfxPath,
        [securestring]$PfxPassword
    )

    if (-not $ClientId)     { $ClientId = $env:ClientId }
    if (-not $PfxPath)      { $PfxPath  = $env:PfxPath }
    if (-not $PfxPassword)  { $PfxPassword = (ConvertTo-SecureString $env:PfxPassword -AsPlainText -Force) }

    if (-not (Test-Path $PfxPath)) {
        throw "❌ Zertifikat nicht gefunden unter $PfxPath"
    }

    # --------------------------
    # Microsoft Graph Login
    # --------------------------
    Log "🔐 Verbinde zu Microsoft Graph im Tenant $TenantId ..."
    $cert = Get-PfxCertificate -FilePath $PfxPath
    Connect-MgGraph -ClientId $ClientId `
                    -TenantId $TenantId `
                    -Certificate $cert `
                    -NoWelcome
    $ctx = Get-MgContext
    if (-not $ctx -or $ctx.TenantId -ne $TenantId) {
        throw "❌ Graph-Login fehlgeschlagen oder falscher Tenant (aktueller: $($ctx.TenantId))"
    }
    Log "✅ Graph verbunden."
}
# --------------------------------------------------------------------
function Connect-PnP {
    param(
        [Parameter(Mandatory=$true)][string]$TenantId,   # Kunden-Tenant ID oder Domain
        [Parameter(Mandatory=$true)][string]$SPOUrl,     # z.B. https://contoso-admin.sharepoint.com
        [string]$ClientId,
        [string]$PfxPath,
        [securestring]$PfxPassword
    )

    Log "Try to connect to $SPOUrl ..."

    if (-not $ClientId)     { $ClientId = $env:ClientId }
    if (-not $PfxPath)      { $PfxPath  = $env:PfxPath }
    if (-not $PfxPassword)  { $PfxPassword = (ConvertTo-SecureString $env:PfxPassword -AsPlainText -Force) }

    if (-not (Test-Path $PfxPath)) {
        throw "❌ Zertifikat nicht gefunden unter $PfxPath"
    }

    # --------------------------
    # 1) PnP.PowerShell Login
    # --------------------------
    Connect-PnPOnline -Tenant $tenantId `
        -ClientId $ClientId `
        -CertificatePath $PfxPath `
        -CertificatePassword $PfxPassword `
        -Url $SPOUrl -ErrorAction Stop
    
    Log "✅ Connected to $SPOUrl"
}