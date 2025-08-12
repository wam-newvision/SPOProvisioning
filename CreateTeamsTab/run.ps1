param(
    [Parameter(Mandatory=$true)][object]$Request,
    $TriggerMetadata
)

# -------- HTTP Payload lesen --------
# Erwartet: JSON-Body mit Feldern wie TeamId, TenantId, ChannelName, ...
try {
    $bodyText = $Request.Body
    if (-not $bodyText) { throw "Leerer Request-Body." }
    $cfg = $bodyText | ConvertFrom-Json
}
catch {
    $msg = "Ungültiger JSON-Body: $($_.Exception.Message)"
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{ StatusCode = 400; Body = $msg })
    return
}

# -------- Variablen befüllen (Defaults ident zu deiner Vorlage) --------
[string]$TeamId              = $cfg.TeamId
[string]$TenantId            = $cfg.TenantId            ?? "mwpnewvision.onmicrosoft.com"
[string]$ChannelName         = $cfg.ChannelName         ?? ""
[string]$TabDisplayName      = $cfg.TabDisplayName      ?? "ProjectAI"
[string]$ContentUrl          = $cfg.ContentUrl          ?? "https://teams.sailing-ninoa.com"
[string]$WebsiteUrl          = $cfg.WebsiteUrl          ?? "https://teams.sailing-ninoa.com"
[string]$EntityId            = $cfg.EntityId            ?? "home"
[string]$TeamsAppExternalId  = $cfg.TeamsAppExternalId  ?? "2a357162-7738-459a-b727-8039af89a684"

$ErrorActionPreference = "Stop"
$PSModuleAutoloadingPreference = 'None'  # wir laden gezielt aus wwwroot\Modules

# -------- Helpers & Core (gemeinsam aus wwwroot\helpers) --------
$functionRoot = Split-Path -Parent $PSScriptRoot       # …\wwwroot
$helpersDir   = Join-Path $functionRoot 'helpers'

. (Join-Path $helpersDir 'LoggingFunctions.ps1')
. (Join-Path $helpersDir 'TeamsTab.Core.ps1')
. (Join-Path $helpersDir 'PSHelpers.ps1')   # enthält LoadGraphModule

# -------- Graph-Module aus \modules laden (PnP-frei) --------
$modulesDir = Join-Path $PSScriptRoot 'modules'
$GraphVersion = "2.29.1"

foreach ($m in @('Microsoft.Graph.Authentication','Microsoft.Graph.Teams','Microsoft.Graph.Groups')) {
    Log "LoadGraphModule -ModuleName $m -FktPath $modulesDir -GraphVersion $GraphVersion"
    LoadGraphModule -ModuleName $m -FktPath $modulesDir -GraphVersion $GraphVersion
}

# -------- Graph-Login --------
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

    # 1) Warten bis App wirklich "ready" ist
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

    # 3) Sichtbarkeit prüfen
    Wait-TeamsTabVisible -TeamId $resolvedTeamId -ChannelId $channelId -TabDisplayName $TabDisplayName -CatalogAppId $catalogAppId -TimeoutSeconds 45 -IntervalSeconds 3

    Log "✅ Tab '$TabDisplayName' im Channel '$channelNameResolved' erstellt."

    $result = @{
        ok = $true
        teamId = $resolvedTeamId
        channelId = $channelId
        tab = $TabDisplayName
    } | ConvertTo-Json -Compress

    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = 200
        Body       = $result
        Headers    = @{ "Content-Type" = "application/json" }
    })
}
catch {
    $err = "❌ Fehler: $($_.Exception.Message)"
    Log $err
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = 500
        Body       = $err
    })
}
