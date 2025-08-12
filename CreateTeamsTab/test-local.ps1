param(
    [string]$payload,
    [string]$payloadFile,
    [string]$TeamId,          # GUID, Name/Alias oder Teams-URL (mit groupId=)
    [string]$TenantId       = "mwpnewvision.onmicrosoft.com",
    [string]$ChannelName    = "",               # leer => primaryChannel
    [string]$TabDisplayName = "ProjectAI",
    [string]$ContentUrl     = "https://teams.sailing-ninoa.com",
    [string]$WebsiteUrl     = "https://teams.sailing-ninoa.com",
    [string]$EntityId       = "home",
    [string]$TeamsAppExternalId = "2a357162-7738-459a-b727-8039af89a684" # Manifest-ID deiner Custom App
)

$ErrorActionPreference = "Stop"
$PSModuleAutoloadingPreference = 'None'  # wir laden gezielt aus wwwroot\Modules

# ------------------------------------------------------------
# Payload laden (Datei hat Vorrang)
# ------------------------------------------------------------
if ($payloadFile -and (Test-Path $payloadFile)) {
    $cfg = Get-Content -Path $payloadFile -Raw | ConvertFrom-Json
} elseif ($payload) {
    $cfg = $payload | ConvertFrom-Json
}
if ($cfg) {
    if ($cfg.TeamId)            { $TeamId = $cfg.TeamId }
    if ($cfg.TenantId)          { $TenantId = $cfg.TenantId }
    if ($cfg.ChannelName)       { $ChannelName = $cfg.ChannelName }
    if ($cfg.TabDisplayName)    { $TabDisplayName = $cfg.TabDisplayName }
    if ($cfg.ContentUrl)        { $ContentUrl = $cfg.ContentUrl }
    if ($cfg.WebsiteUrl)        { $WebsiteUrl = $cfg.WebsiteUrl }
    if ($cfg.EntityId)          { $EntityId = $cfg.EntityId }
    if ($cfg.TeamsAppExternalId){ $TeamsAppExternalId = $cfg.TeamsAppExternalId }
}

# ------------------------------------------------------------
# Helpers & Core
# ------------------------------------------------------------
$functionRoot = Split-Path -Parent $PSScriptRoot       # …\wwwroot
$helpersDir   = Join-Path $functionRoot 'helpers'

. (Join-Path $helpersDir 'LoggingFunctions.ps1')
. (Join-Path $helpersDir 'TeamsTab.Core.ps1')
. (Join-Path $helpersDir 'PSHelpers.ps1')   # enthält LoadGraphModule

# ------------------------------------------------------------
# Graph-Module analog zu PnP-Loader aus wwwroot\Modules laden
# ------------------------------------------------------------
$appRoot    = Split-Path -Parent $PSScriptRoot      # …\wwwroot
$modulesDir = Join-Path $appRoot 'Modules'

# Diese drei Module reichen für Tabs/App-Katalog/Teams/Gruppen
foreach ($m in @('Microsoft.Graph.Authentication','Microsoft.Graph.Teams','Microsoft.Graph.Groups')) {
    LoadGraphModule -ModuleName $m -FktPath $modulesDir
}

# ------------------------------------------------------------
# Graph-Login (Delegated; für App-Only ggf. Zertifikat/ClientCreds)
# ------------------------------------------------------------
Connect-MgGraph -Scopes `
    "TeamsAppInstallation.ReadWriteForTeam", `
    "TeamsTab.ReadWriteForTeam", `
    "Group.Read.All", `
    "AppCatalog.Read.All"

try {
    Log "ℹ️ TeamId Eingabe: '$TeamId'"
    $resolvedTeamId = Resolve-TeamId -TeamRef $TeamId
    Log "ℹ️ TeamId aufgelöst: $resolvedTeamId"

    $chan = Get-ChannelId -ResolvedTeamId $resolvedTeamId -ChannelName $ChannelName
    $channelId = $chan[0]
    $channelNameResolved = $chan[1]

    # App-Katalog: externalId -> Katalog-ID (teamsAppId)
    $catalogAppId = Get-CatalogAppId -ExternalId $TeamsAppExternalId
    Log "ℹ️ Custom-App (externalId=$TeamsAppExternalId) im App-Katalog gefunden. Catalog-ID: '$catalogAppId'"

    # Channel-Infos holen (MembershipType für Installationsziel)
    $channelInfo = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/teams/$resolvedTeamId/channels/$channelId"

    switch ($channelInfo.membershipType) {
        'standard' {
            Log "[InstallCheck] Standard-Channel erkannt – App im Team installieren"
            Get-TeamsAppInstalled -ResolvedTeamId $resolvedTeamId -CatalogAppId $catalogAppId
        }
        'private' {
            Log "[InstallCheck] Privater Channel erkannt – App im Channel installieren"
            Get-TeamsAppInstalledForChannel -ResolvedTeamId $resolvedTeamId -ChannelId $channelId -CatalogAppId $catalogAppId
        }
        'shared' {
            Log "[InstallCheck] Shared Channel erkannt – App im Channel installieren"
            Get-TeamsAppInstalledForChannel -ResolvedTeamId $resolvedTeamId -ChannelId $channelId -CatalogAppId $catalogAppId
        }
        default {
            throw "Unbekannter membershipType: $($channelInfo.membershipType)"
        }
    }

    # 1) Warten bis App wirklich "ready" ist (richtiger Var-Name!)
    Wait-TeamsAppReady -TeamId $resolvedTeamId -CatalogAppId $catalogAppId -TimeoutSeconds 20

    # 2) Tab anlegen
    $tabParams = @{
        TeamId         = $resolvedTeamId
        ChannelId      = $channelId
        TabDisplayName = $TabDisplayName
        TeamsAppId     = $catalogAppId     # Katalog-ID (nicht externalId)
        EntityId       = $EntityId
        ContentUrl     = $ContentUrl
        WebsiteUrl     = $WebsiteUrl
    }
    Add-GraphTeamsTab @tabParams

    # 3) Sichtbarkeit prüfen (richtiger Var-Name & ChannelId verwenden)
    Wait-TeamsTabVisible -TeamId $resolvedTeamId -ChannelId $channelId -TabDisplayName $TabDisplayName -CatalogAppId $catalogAppId -TimeoutSeconds 45 -IntervalSeconds 3

    Log "✅ Tab '$TabDisplayName' im Channel '$channelNameResolved' erstellt."

    # Deep-Link erzeugen
    $entityForLink = $EntityId
    if ($catalogAppId -like 'com.microsoft.teamspace.tab.*') { $entityForLink = $ContentUrl }

    $deeplink = New-TeamsTabDeepLink @{
        AppId      = $catalogAppId
        EntityId   = $entityForLink
        ContentUrl = $ContentUrl
        TabName    = $TabDisplayName
        TeamId     = $resolvedTeamId
        ChannelId  = $channelId
    }

    # Channel-Nachricht posten
    $msgHtml = "🔔 <b>$TabDisplayName</b> wurde installiert. Klick zum Öffnen: <a href=""$deeplink"">$TabDisplayName</a><br/>
    <i>Tipp:</i> Schreibe <b>@NewViBot</b> 'Hallo' – das aktiviert den Agent im Channel."
    $sent = Send-TeamsChannelMessage -TeamId $resolvedTeamId -ChannelId $channelId -Html $msgHtml
    if ($sent.id) { Log "📨 Nachricht erfolgreich gepostet (Message ID: $($sent.id))" } else { Log "⚠️ Nachricht konnte nicht bestätigt werden." }

    if ($env:DEBUG -and $env:DEBUG.ToLower() -eq 'true') {
        try { Start-Process $deeplink | Out-Null } catch { Log "⚠️ Konnte Deep-Link nicht öffnen: $($_.Exception.Message)" }
    }

}
catch {
    Log "❌ Fehler: $($_.Exception.Message)"
    throw
}
