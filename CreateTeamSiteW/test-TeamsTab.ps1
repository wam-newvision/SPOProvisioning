param(
    [Parameter(Mandatory)] [string]$TeamId,          # GUID, Name/Alias oder Teams-URL (mit groupId=)
    [string]$TenantId            = "mwpnewvision.onmicrosoft.com",
    [string]$ChannelName         = "",               # leer => primaryChannel
    [string]$TabDisplayName      = "NewViBot",
    [string]$ContentUrl          = "https://teams.sailing-ninoa.com",
    [string]$WebsiteUrl          = "https://teams.sailing-ninoa.com",
    [string]$EntityId            = "home",
    [string]$TeamsAppExternalId  = "2a357162-7738-459a-b727-8039af89a684" # Manifest-ID deiner Custom App
)

$ErrorActionPreference = "Stop"
$PSModuleAutoloadingPreference = 'None'
$env:DEBUG = 'true'

# Logging + deine Add-GraphTeamsTab-Funktion
# -------- Helpers laden --------
$functionRoot = Split-Path -Parent $PSScriptRoot
$helpersDir   = Join-Path $functionRoot 'Helpers'

. (Join-Path $helpersDir 'LoggingFunctions.ps1')
. (Join-Path $helpersDir 'TeamsTab.Core.Graph.ps1')

# Login (Delegated). Für App-Only ggf. auf Zertifikat + Application Permissions umstellen.
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

    $catalogAppId = Get-CatalogAppId -ExternalId $TeamsAppExternalId
    Log "ℹ️ Custom-App (externalId=$TeamsAppExternalId) im App-Katalog gefunden. Catalog-ID: '$catalogAppId'"

    # Channel-Infos holen
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

    # 1. Warten, bis App vollständig installiert ist
    Wait-TeamsAppReady -TeamId $resolvedTeamId -CatalogAppId $CatalogAppId -TimeoutSeconds 20

    # 2. Tab anlegen – über deine vorhandene Funktion (Katalog-ID binden); Splatting vermeidet Backtick-Fallen
    $tabParams = @{
        TeamId         = $resolvedTeamId
        ChannelId      = $channelId
        TabDisplayName = $TabDisplayName
        TeamsAppId     = $catalogAppId     # <-- Catalog-ID (nicht externalId)
        EntityId       = $EntityId
        ContentUrl     = $ContentUrl
        WebsiteUrl     = $WebsiteUrl
    }
    Add-GraphTeamsTab @tabParams

    # 3. Sichtbarkeit prüfen + Retries
    Wait-TeamsTabVisible -TeamId $resolvedTeamId -ChannelId $ChannelId -TabDisplayName $TabDisplayName -CatalogAppId $CatalogAppId -TimeoutSeconds 45 -IntervalSeconds 3

# 4) Wenn $null → 3 zusätzliche Versuche mit 2s Backoff

    Log "✅ Tab '$TabDisplayName' im Channel '$channelNameResolved' erstellt."

    # EntityId für den Link bestimmen (Website-Tab => URL, Custom-Tab => EntityId)
    $entityForLink = $EntityId
    if ($catalogAppId -like 'com.microsoft.teamspace.tab.*') {
        $entityForLink = $ContentUrl
    }

    # Deep-Link erzeugen (Helper aus TeamsTab.Core.ps1)
    $deeplinkParams = @{
        AppId      = $catalogAppId
        EntityId   = $entityForLink
        ContentUrl = $ContentUrl
        TabName    = $TabDisplayName
        TeamId     = $resolvedTeamId
        ChannelId  = $channelId
    }
    $deeplink = New-TeamsTabDeepLink @deeplinkParams

    # Channel-Nachricht mit Deep Link posten
    $msgHtml = "🔔 <b>$TabDisplayName</b> wurde installiert. Klick zum Öffnen: <a href=""$deeplink"">$TabDisplayName</a><br/>
    <i>Tipp:</i> Schreibe <b>@NewViBot</b> 'Hallo' – das aktiviert den Agent im Channel."

    $sent = Send-TeamsChannelMessage -TeamId $resolvedTeamId -ChannelId $channelId -Html $msgHtml

    if ($sent.id) {
        Log "📨 Nachricht erfolgreich gepostet (Message ID: $($sent.id))"
    } else {
        Log "⚠️ Nachricht konnte nicht bestätigt werden."
    }

<#
    Log "⏳ Tab noch nicht sichtbar – starte Refresh-Trigger (Dummy Website-Tab)."
    # 2) Fallback: Dummy-Website-Tab anlegen & löschen
    Invoke-TeamsTabsRefreshTrigger -TeamId $resolvedTeamId -ChannelId $channelId -Delete $true


    # 1) Warten, bis Tab in /tabs sichtbar ist
    $tabId = Wait-TeamsTabVisible -TeamId $resolvedTeamId -ChannelId $channelId -TabDisplayName $TabDisplayName -TimeoutSeconds 45 -IntervalSeconds 3

    if (-not $tabId) {

        # 3) Noch einmal prüfen
        $tabId = Wait-TeamsTabVisible -TeamId $resolvedTeamId -ChannelId $channelId -TabDisplayName $TabDisplayName -TimeoutSeconds 30 -IntervalSeconds 3
    }

    if ($tabId) {
        Log "✅ Tab ist sichtbar (TabId: $tabId)."
        # (Optional) Nudge per Rename hin/retour – schadet nicht:
        $refreshParams = @{
            TeamId         = $resolvedTeamId
            ChannelId      = $channelId
            TabDisplayName = $TabDisplayName
            TabId          = $tabId
        }
        $null = Invoke-TeamsTabRefresh @refreshParams
    } else {
        Log "⚠️ Tab weiterhin nicht sichtbar. (Serverseitig evtl. vorhanden, UI-Refresh abhängig vom Client.)"
    }    
#>

    # Optional lokal öffnen, wenn DEBUG=true
    if ($env:DEBUG -and $env:DEBUG.ToLower() -eq 'true') {
        try { Start-Process $deeplink | Out-Null } catch { Log "⚠️ Konnte Deep-Link nicht öffnen: $($_.Exception.Message)" }
    }

}
catch {
    Log "❌ Fehler: $($_.Exception.Message)"
    throw
}
