# --------------------------------------------------------------------
# Include Helper Functions
# --------------------------------------------------------------------
. (Join-Path $PSScriptRoot 'LoggingFunctions.ps1')
. (Join-Path $PSScriptRoot 'SPOAdminFunctions.ps1')

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
