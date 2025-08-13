param(
    [Parameter(Mandatory=$true)][object]$Request,
    $TriggerMetadata
)

trap {
    $e = $_.Exception
    $body = @{
        ok = $false
        error = @{
            message = $e.Message
            type    = $e.GetType().FullName
            inner   = $e.InnerException?.Message
            script  = $($_.InvocationInfo?.ScriptName)
            line    = $($_.InvocationInfo?.ScriptLineNumber)
        }
    } | ConvertTo-Json -Depth 5 -Compress

    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = 500
        Body       = $body
        Headers    = @{ "Content-Type" = "application/json" }
    })
    continue
}

# -------- HTTP Payload lesen (robust) --------
try {
    $raw = $Request.Body

    if ($null -eq $raw) { throw "Leerer Request-Body." }

    switch ($raw.GetType().FullName) {
        'System.String' {
            if ([string]::IsNullOrWhiteSpace($raw)) { throw "Leerer Request-Body." }
            $cfg = $raw | ConvertFrom-Json -ErrorAction Stop
        }
        'System.Byte[]' {
            $text = [System.Text.Encoding]::UTF8.GetString($raw)
            if ([string]::IsNullOrWhiteSpace($text)) { throw "Leerer Request-Body." }
            $cfg = $text | ConvertFrom-Json -ErrorAction Stop
        }
        'System.IO.MemoryStream' {
            $reader = New-Object System.IO.StreamReader($raw, [System.Text.Encoding]::UTF8)
            $text = $reader.ReadToEnd()
            if ([string]::IsNullOrWhiteSpace($text)) { throw "Leerer Request-Body." }
            $cfg = $text | ConvertFrom-Json -ErrorAction Stop
        }
        default {
            # Falls das Runtime den Body bereits zu PSCustomObject deserialisiert hat
            if ($raw -is [pscustomobject]) {
                $cfg = $raw
            } else {
                # Letzter Versuch: stringify und parsen
                $text = $raw | ConvertTo-Json -Depth 50
                $cfg  = $text | ConvertFrom-Json -ErrorAction Stop
            }
        }
    }
}
catch {
    $e = $_.Exception
    $body = @{
        ok = $false
        error = @{
            message = $e.Message
            type    = $e.GetType().FullName
            inner   = $e.InnerException?.Message
            script  = $($_.InvocationInfo?.ScriptName)
            line    = $($_.InvocationInfo?.ScriptLineNumber)
        }
    } | ConvertTo-Json -Depth 5 -Compress

    Log "❌ Fehler: $($e.Message)"
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = 500
        Body       = $body
        Headers    = @{ "Content-Type" = "application/json" }
    })
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
$helpersDir   = Join-Path $functionRoot 'Helpers'

. (Join-Path $helpersDir 'LoggingFunctions.ps1')
. (Join-Path $helpersDir 'AdminFunctions.ps1')   # enthält LoadGraphModule
. (Join-Path $helpersDir 'TeamsTab.Core.ps1')

# ------------- Framework-Helpers ----------------------------
$InformationPreference = 'Continue'
$CurDir                = Get-Location
$certsDir              = Join-Path $functionRoot 'Certs'
Get-ChildItem -Path $certsDir
$modulesDir            = Join-Path $functionRoot 'Modules'
Log "---------------------- Start Logging ---------------------"
Log "PowerShell Version: $($PSVersionTable.PSVersion)"
Log "Current Directory : $CurDir"
Log "FunctionRoot      : $functionRoot"
Log "PSScriptRoot      : $PSScriptRoot"
Log "CertLocation      : $certsDir"
Log "ModulesLocation   : $modulesDir"
Log "----------------------------------------------------------"

# -------- Graph-Module aus \modules laden (PnP-frei) --------
$GraphVersion = "2.29.1"
Log "ℹ️ Loading Modules from Graph-Version: $GraphVersion ..."

foreach ($m in @('Microsoft.Graph.Authentication','Microsoft.Graph.Teams','Microsoft.Graph.Groups')) {
    Log "LoadGraphModule -ModuleName $m -FktPath $modulesDir -GraphVersion $GraphVersion"
    LoadGraphModule -ModuleName $m -FktPath $modulesDir -GraphVersion $GraphVersion
}

# -------- Graph-Login --------
Log "ℹ️ Graph-Login: Tenant: '$TenantId' - ClientId: '$env:ClientId'"
Connect-MSGraph -ClientId $env:ClientId -TenantId $TenantId
<#
#Connect-MgGraph -TenantId $TenantId 
# Aufruf als User 
Connect-MgGraph -Scopes `
    "TeamsAppInstallation.ReadWriteForTeam", `
    "TeamsTab.ReadWriteForTeam", `
    "Group.Read.All", `
    "AppCatalog.Read.All"
#>

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
