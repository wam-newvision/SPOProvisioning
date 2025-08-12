# --------------------------------------------------------------------
# Include Helper Functions
# --------------------------------------------------------------------
. (Join-Path $PSScriptRoot 'LoggingFunctions.ps1')
# --------------------------------------------------------------------

function Connect-PnP {
    param(
        [string]$tenantId,  
        [string]$SPOUrl,   # "https://<tenant>.sharepoint.com/sites/NEUESITE" or "https://<tenant>-admin.sharepoint.com" 
        [string]$ClientId,
        [string]$PfxPath,
        [securestring]$PfxPassword
    )

    Log "Try to connect to $SPOUrl ..."

    if (-not $ClientId) {$ClientId = $env:ClientId}
    if (-not $PfxPath) {$PfxPath = $env:PfxPath}
    if (-not $PfxPassword) {$PfxPassword = (ConvertTo-SecureString $env:PfxPassword -AsPlainText -Force)}
    
    if (-not (Test-Path $PfxPath)) {
        Log "❌ File not found at $PfxPath"
        throw "Certificate not found"
    }

    Connect-PnPOnline -Tenant $tenantId `
        -ClientId $ClientId `
        -CertificatePath $PfxPath `
        -CertificatePassword $PfxPassword `
        -Url $SPOUrl -ErrorAction Stop
    
    Log "✅ Connected to $SPOUrl"

}