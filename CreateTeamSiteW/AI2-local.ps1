# ========================================================================
#  PowerShell: Copilot Studio Teams App als Tab automatisiert hinzufügen
#  Mit Helper-Funktion Add-GraphTeamsTab
# ========================================================================

param(
    [string]$Site,                              # Teams Name (wie in Teams angezeigt)
    [string]$channelName    = "General",        # Kanalname (Standard: "Allgemein")
    [string]$tabDisplayName = "ProjectAgent",   # Tab-Name
    [string]$AgentVersion   = "107"             # Agent (=zipfile) Version
)

# --------------------------------------------------------------------
# Include Helper Functions
# --------------------------------------------------------------------
. (Join-Path $PSScriptRoot 'helpers\SPOAdminFunctions.ps1')
. (Join-Path $PSScriptRoot 'helpers\PSHelpers.ps1')
. (Join-Path $PSScriptRoot 'helpers\LoggingFunctions.ps1')

# --------------------------------------------------------------------
# Eingaben zuweisen
# --------------------------------------------------------------------
#$FktPath      = "C:\Functions\dms-provisioning\CreateTeamSiteW"
$FktPath      = $PSScriptRoot
$zipPath      = Join-Path $PSScriptRoot "AIAgents\zip\$tabDisplayName$AgentVersion.zip" # Exportiertes Copilot Studio ZIP
Log "Zip-Datei: " $zipPath

$ClientId     = "5a19516e-dc54-4d2f-aebc-f1b679a69457"
#$clientSecret = $env:AZURE_CLIENT_SECRET

$tenantId     = "mwpnewvision.onmicrosoft.com"
$siteTitle    = $Site
#$hubName      = "ProjektHub"

$PfxPath      = Join-Path $FktPath 'certs\mwpnewvision.pfx'
$PfxPwd       = "MyP@ssword!" # Setze hier dein PFX-Passwort
$PfxPassword  = (ConvertTo-SecureString $PfxPwd -AsPlainText -Force)

# --------------------------------------------------------------------
# Alias / URLs
# --------------------------------------------------------------------
$base     = $tenantId.Split('.')[0]
$alias    = ($siteTitle -replace '\s+', '')
$siteUrl  = "https://${base}.sharepoint.com/sites/$alias"
$adminUrl = "https://${base}-admin.sharepoint.com"
Log "🔗 SiteUrl = $siteUrl"

# ------------------------------------------------------------------------
# Hauptskript: Anmeldung, App-Bereitstellung, Tab-Erstellung
# ------------------------------------------------------------------------
$env:DEBUG = 'true'

# 1. Prüfe, ob ZIP existiert
if (!(Test-Path $zipPath)) {
    throw "Teams ZIP $zipPath nicht gefunden!"
}

# 2. manifest.json extrahieren
Add-Type -AssemblyName System.IO.Compression.FileSystem
$manifestJson = $null

$manifestEntry = [System.IO.Compression.ZipFile]::OpenRead($zipPath).Entries | Where-Object { $_.Name -eq "manifest.json" }
if ($manifestEntry) {
    $reader = New-Object IO.StreamReader $manifestEntry.Open()
    $manifestJson = $reader.ReadToEnd()
    $reader.Close()
}

if (-not $manifestJson) {
    throw "manifest.json nicht in $zipPath gefunden!"
}

# 3. AppId und AppName aus manifest.json lesen
$manifest = $manifestJson | ConvertFrom-Json
$appManifestId = $manifest.id
if (-not $appManifestId) { $appManifestId = $manifest.appId }
if (-not $appManifestId) { throw "Keine AppId im Manifest gefunden!" }

# AppName automatisch bestimmen (String oder Objekt!)
if ($manifest.name -is [string]) {
    $appName = $manifest.name
} elseif ($manifest.name.short) {
    $appName = $manifest.name.short
} elseif ($manifest.name.full) {
    $appName = $manifest.name.full
} else {
    throw "Kein App-Name im Manifest gefunden!"
}

Log "Manifest AppId: $appManifestId"
Log "Manifest AppName: $appName"

# 1. Microsoft Graph laden & anmelden
Log "Import-Module Microsoft.Graph..."
Import-Module Microsoft.Graph
#Import-Module "C:\Users\wam\Documents\PowerShell\Modules\Microsoft.Graph\2.29.1\Microsoft.Graph.psd1"

Log "Connect-MgGraph..."
Connect-MgGraph -NoWelcome -Scopes `
    "User.Read",
    "Group.ReadWrite.All",
    "Channel.ReadBasic.All",
    "ChannelSettings.ReadWrite.All",
    "AppCatalog.ReadWrite.All",
    "TeamsTab.ReadWrite.All",
    "TeamsAppInstallation.ReadWriteForTeam"

    <#
$token = (Get-MgContext).AccessToken
Log "Token: $token"
$headers = @{
  "Authorization" = "Bearer $token"
  "Content-Type"  = "application/json"
}
Log "Headers: $($headers | ConvertTo-Json -Depth 10)"
#>

# 4. ALLE Apps im App-Katalog mit passendem DisplayName (inkl. Paging)
Log "Frage Teams Apps mit Name '$appName' im App-Katalog ab..."
$baseUri = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps"
$uri = $baseUri
$matchingApps = @()
do {
    try {
        $allApps = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
    } catch {
        Write-Host "Invoke-MgGraphRequest-Fehler: $($_.Exception.Message)"
        if ($_.Exception.Response) {
            Write-Host "HTTP Response: $($_.Exception.Response.Content.ReadAsStringAsync().Result)"
        }
        throw "Fehler beim Aufruf von $uri"
    }
    $pageMatches = $allApps.value | Where-Object { $_.displayName -eq $appName }
    if ($pageMatches) { $matchingApps += $pageMatches }
    # Prüfen, ob ein NextLink existiert und NICHT $null oder leer ist
    if ($allApps.PSObject.Properties.Name -contains '@odata.nextLink' -and $allApps.'@odata.nextLink') {
        $uri = $allApps.'@odata.nextLink'
    } else {
        $uri = $null
    }
} while ($uri)

if ($matchingApps) {
     Log "Apps mit Namen '$appName' gefunden, prüfe auf ManifestId '$appManifestId' ..."
    # 5. Nur diese im Detail prüfen (auf ManifestId)
    $teamsAppId = $null
    foreach ($app in $matchingApps) {
        $details = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/$($app.id)"
        #Log "DETAILS: " ($details | ConvertTo-Json -Depth 10)
        log $app.displayName "Details AppId: $($details.externalId)"
        if ($details.externalId -eq $appManifestId) {
            $teamsAppId = $app.id
            Log "Gefunden: TeamsAppId = $teamsAppId für Manifest AppId $appManifestId"
            break
        }
    }
}

$noAPP = $false
if ($teamsAppId) {
    Log "Teams-App bereits im Katalog: $teamsAppId"
    $deleteAPP = $false
    if ($deleteAPP) {
        Log "❌ Lösche Teams-App aus dem App-Katalog: $teamsAppId"
        Invoke-MgGraphRequest -Method DELETE `
            -Uri "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/$teamsAppId"
        Log "App wurde entfernt."
        $noAPP = $true
    }
} else {
    Log "⚠️  Keine App-ID in Teams APPs gefunden"
    $noAPP = $true
}

$noAPP = $false
if ($noAPP) {
    Log "⬆️  Lade neue Teams-App hoch..."
    $zipBytes = [System.IO.File]::ReadAllBytes($zipPath)
    $response = Invoke-MgGraphRequest -Method POST `
        -Uri "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps" `
        -Body $zipBytes `
        -ContentType "application/zip"
    $teamsAppId = $response.id
    Log "Neue App bereitgestellt: $teamsAppId"
    throw "!!! CAUTION !!! You have to enable/unblock this APP manually in Teams Admin Center !!!"
}

# 6. Team und Kanal suchen
Log "🔍 Suche Team '$Site'..."
$team = Get-MgTeam -Filter "displayName eq '$Site'"
if (!$team) { ErrorExit "Team '$Site' nicht gefunden!" }
$teamId = $team.Id
$teamName = $team.DisplayName

Log "🔍 Suche Kanal '$channelName'..."
$channel = Get-MgTeamChannel -TeamId $teamId | Where-Object { $_.DisplayName -eq $channelName }
if (!$channel) { ErrorExit "Channel '$channelName' nicht gefunden!" }
$channelId = $channel.Id

# 7. Teams-App im Team installieren, falls noch nicht installiert
Log "🔍 Suche APPs in Team '$channelName'..."
$uri = "https://graph.microsoft.com/v1.0/teams/$teamId/installedApps"
$installedApps = Invoke-MgGraphRequest -Method GET -Uri $uri

$installed = $false
foreach ($app in $installedApps.value) {
    # Logge original (Base64)
    # Log ("InstalledApp: " + ($app | ConvertTo-Json -Depth 10))
    try {
        $decoded = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($app.id))
        #Log ("Decoded InstalledApp id: $decoded")
        if ($decoded -match [regex]::Escape($teamsAppId)) {
            $installed = $true
            Log "Teams-App bereits im Team installiert! ($teamsAppId)"
            break
        }
    } catch {
        # Falls $app.id nicht Base64 ist (selten, aber möglich)
        if ($app.id -match [regex]::Escape($teamsAppId)) {
            $installed = $true
            Log "Teams-App bereits im Team installiert! ($teamsAppId)"
            break
        }
    }
}

if (-not $installed) {
    Log "📦 Installiere App im Team..."
    $installBody = [ordered]@{}
    $installBody."teamsApp@odata.bind" = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps('$teamsAppId')"
    Invoke-MgGraphRequest -Method POST `
        -Uri "https://graph.microsoft.com/v1.0/teams/$teamId/installedApps" `
        -Body ($installBody | ConvertTo-Json -Depth 10)
    Start-Sleep -Seconds 10 # Warte, bis App bereit ist
    Log "App installiert."
}

# ------------------------------------------------------------------------
# 8. Content URL 

# für Copilot Studio Bots meist so:
# $teamsAppId = "com.microsoft.copilot.studio"
# $contentUrl = "https://copilotstudio.microsoft.com/bots/12345678
# $contentUrl = "https://teams.microsoft.com/l/entity/$teamsAppId/" # ggf. anpassen

# Für Webpages:
#$teamsAppId     = "com.microsoft.teamspace.tab.web"
#$contentUrl     = "https://sailing-ninoa.com"

# Für SharePoint Document Libraries:
#$teamsAppId     = "com.microsoft.teamspace.tab.file.staticviewer.sharepoint" # Document Library-Tab
#$contentUrl     = "https://mwpnewvision.sharepoint.com/sites/1578/Shared Documents/Forms/AllItems.aspx" 

# Für OneNote-Notizbücher:
# $teamsAppId     = "com.microsoft.teamspace.tab.onenote" # OneNote-Tab

# Für SharePoint-Seiten:
# Installiere Teams App ins Team (Scope: Team)

$teamsAppId     = "com.microsoft.teamspace.tab.web.site" # SharePoint Pages and Lists-Tab
$contentUrl     = "https://mwpnewvision.sharepoint.com/sites/1578/SitePages/Forms/ByAuthor.aspx"
#$contentUrl     = "https://mwpnewvision.sharepoint.com/sites/projekte/Test1/Forms/AllItems.aspx"

# ------------------------------------------------------------------------
# 9. Helper-Funktion zum Anlegen des Tabs aufrufen

#$TabType               = [TabType]::SharePointPageAndList  # Typ des Tabs (z.B. SharePointPageAndList, WebSite, etc.)
#$TabType               = [TabType]::WebSite  # Typ des Tabs (z.B. SharePointPageAndList, WebSite, etc.)
#$TabDisplayName        = "AI Agent"  # Name des Tabs
#$contentUrl           = "https://newvision.eu/impressum/"
#$WebSiteUrlDisplayName = "NewVision"  # DisplayName der Website, die im Tab angezeigt werden soll

# WebSite Tab hinzufügen
Log "WebSite Tab hinzufügen zu Team: '$teamName' ..."

$PnP = $false
if ($PnP) {
    Log "Starte Teams Tab Einrichtung via PnP..."
    Connect-PnP -Tenant $tenantId -SPOUrl $adminUrl -ClientId $ClientId -PfxPath $PfxPath -PfxPassword $PfxPassword
    $team = Get-PnPTeamsTeam -Identity $alias -ErrorAction Stop

    AddTeamsTab `
        -team $team `
        -TeamsChannel $channel `
        -TabDisplayName $TabDisplayName `
        -WebSiteUrl $contentUrl `
        -TabType WebSite

} else {
    Log "PnP nicht verfügbar, nutze Graph-API zum Hinzufügen des Tabs..."
    #$AppID = "0ae35b36-0fd7-422e-805b-d53af1579093" # Sharepoint Pages and Lists App ID
    $AppID = "bd871fde-dc2d-49af-a140-59693a6a1d33" # Project Manager Agent App ID
    $paramsApp = @{
        'teamsApp@odata.bind' = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/$AppID"
    }
    Log "Binde die TeamsAPP '$TabDisplayName' in das Team '$alias' / Channel '$channelName' ein..."
    New-MgTeamInstalledApp -TeamId $teamId -BodyParameter $paramsApp

    $paramsTab = @{
        displayName = $tabDisplayName
        'teamsApp@odata.bind' = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/$AppID"
        #configuration = @{
        #    contentUrl = $contentUrl
        #    websiteUrl = $contentUrl
        #}
    }
    Log "Erstelle Teams-Register '$TabDisplayName' im Team '$alias' / Channel '$channelName' ..."
    New-MgTeamChannelTab -TeamId $teamId -ChannelId $channelId -BodyParameter $paramsTab

}

<# # Beispielaufruf der Graph-Helper-Funktion Add-GraphTeamsTab
    Add-GraphTeamsTab `
        -TeamId $teamId `
        -ChannelId $channelId `
        -TabDisplayName $tabDisplayName `
        -TeamsAppId $teamsAppId `
        -ContentUrl $contentUrl
#>

# ------------------------------------------------------------------------
# Fertig! Der Copilot Studio Bot ist als Tab im Teams-Channel verfügbar!
# ------------------------------------------------------------------------
