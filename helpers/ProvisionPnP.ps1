# --------------------------------------------------------------------
# Include Helper Functions
# --------------------------------------------------------------------
. (Join-Path $PSScriptRoot 'LoggingFunctions.ps1')
. (Join-Path $PSScriptRoot 'SPOAdminFunctions.ps1')

# --------------------------------------------------------------------
# Module laden (PnP.PowerShell)
# --------------------------------------------------------------------
function LoadPnPPSModule {
    param (
        [string]$PnPVersion,  # $PnPVersion = "3.1.0"
        [string]$FktPath  # $FktPath = "C:\Functions\dms-provisioning\CreateTeamSiteW"
    )
        
    #if (-not $PnPVersion) {$PnPVersion = "3.1.0"} #$PnPVersion = "1.12.0"
    #if (-not $FktPath) {$FktPath = "C:\Functions\dms-provisioning\CreateTeamSiteW"} 

    $modBase = Join-Path $FktPath 'modules'
    $sep = [IO.Path]::PathSeparator
    $env:PSModulePath = "${modBase}${sep}$env:PSModulePath"
    
    $PnPPath = Join-Path $modBase "PnP.PowerShell\$PnPVersion\PnP.PowerShell.psd1"

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

function Enable-SiteCollectionFeatures {
    param(
        [string]$featureId,  # Feature ID, z.B. "b50e3104-6812-424f-a011-cc90e6327318"
        [string]$scope = "Site"  # Standardmäßig "Site"
    )
    
    Log "Aktiviere Site-Collection Feature: $featureId (Scope: $scope)"
    
    try {
        Enable-PnPFeature -Identity $featureId -Scope $scope -ErrorAction Stop
        Log "✅ Feature $featureId erfolgreich aktiviert!"
    } catch {
        Log "❌ Fehler beim Aktivieren des Features $featureId - $($_.Exception.Message)"
    }
}
# --------------------------------------------------------------------
function ProvisionPnPSite {
    param(
        [string]$siteUrl,   # "https://<tenant>.sharepoint.com/sites/NEUESITE"
        [string]$templateFolder,
        [string]$template
    )
    
    $templatePath = Join-Path $templateFolder $template  # Pfad zu XML-Datei
    Log "Provisioniere SPO Site $siteUrl mit $templatePath"

    $result = Invoke-PnPSiteTemplate -Path "$templatePath"
    Log "Invoke-PnPSiteTemplate Result $($result | Out-String)"
}
# --------------------------------------------------------------------
