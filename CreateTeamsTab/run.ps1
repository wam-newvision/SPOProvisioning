param(
    [Parameter(Mandatory=$true)][object]$Request,
    $TriggerMetadata
)

$script:ResponseAlreadySet = $false
function Send-HttpResponse([int]$StatusCode, $Body) {
    if ($script:ResponseAlreadySet) { return }
    $script:ResponseAlreadySet = $true
    if ($Body -isnot [string]) {
        $Body = ($Body | ConvertTo-Json -Depth 10 -Compress)
    }
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = $StatusCode
        Body       = $Body
        Headers    = @{ "Content-Type" = "application/json" }
    })
}

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
    }
    Send-HttpResponse 500 $body
    continue
}

# -------- Body lesen --------
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
            if ($raw -is [pscustomobject]) { $cfg = $raw }
            else {
                $text = $raw | ConvertTo-Json -Depth 50
                $cfg  = $text | ConvertFrom-Json -ErrorAction Stop
            }
        }
    }
}
catch {
    $err = @{ ok = $false; error = $_.Exception.Message }
    Send-HttpResponse 500 $err
}

# -------- Variablen / Defaults --------
[string]$TeamId              = $cfg.TeamId
# KEINE lokale $TenantId mehr -> als ENV (für Helpers)
$env:TenantId                = $cfg.TenantId            ?? "mwpnewvision.onmicrosoft.com"
[string]$ChannelName         = $cfg.ChannelName         ?? ""
[string]$TabDisplayName      = $cfg.TabDisplayName      ?? "ProjectAI"
[string]$ContentUrl          = $cfg.ContentUrl          ?? "https://teams.sailing-ninoa.com"
[string]$WebsiteUrl          = $cfg.WebsiteUrl          ?? "https://teams.sailing-ninoa.com"
[string]$EntityId            = $cfg.EntityId            ?? "home"
[string]$TeamsAppExternalId  = $cfg.TeamsAppExternalId  ?? "2a357162-7738-459a-b727-8039af89a684"

$ErrorActionPreference = "Stop"
$PSModuleAutoloadingPreference = 'None'

# -------- Helpers laden --------
$functionRoot = Split-Path -Parent $PSScriptRoot
$helpersDir   = Join-Path $functionRoot 'Helpers'

. (Join-Path $helpersDir 'LoggingFunctions.ps1')
. (Join-Path $helpersDir 'AdminFunctions.ps1')   # (falls du dort Nicht-Graph-Utilities hast)
. (Join-Path $helpersDir 'GraphRestHelpers.ps1') # <— REST einbinden
. (Join-Path $helpersDir 'TeamsTab.Core.ps1')

# -------- Optional: Health-Check --------
try {
    $null = Invoke-Graph -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization'
} catch {
    throw "Graph REST Health-Check fehlgeschlagen: $($_.Exception.Message)"
}

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

try {
    Log "ℹ️ TeamId Eingabe: '$TeamId'"
    $resolvedTeamId = Resolve-TeamId -TeamRef $TeamId
    Log "ℹ️ TeamId aufgelöst: $resolvedTeamId"

    $chan = Get-ChannelId -ResolvedTeamId $resolvedTeamId -ChannelName $ChannelName
    $channelId = $chan[0]
    $channelNameResolved = $chan[1]

    $catalogAppId = Get-CatalogAppId -ExternalId $TeamsAppExternalId
    Log "ℹ️ Custom-App (externalId=$TeamsAppExternalId) im App-Katalog gefunden. Catalog-ID: '$catalogAppId'"

    $channelInfo = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/teams/$resolvedTeamId/channels/$channelId"

    switch ($channelInfo.membershipType) {
        'standard' {
            Log "[InstallCheck] Standard-Channel – App im Team installieren"
            Get-TeamsAppInstalled -ResolvedTeamId $resolvedTeamId -CatalogAppId $catalogAppId
        }
        'private' {
            Log "[InstallCheck] Privater Channel – App im Channel installieren"
            Get-TeamsAppInstalledForChannel -ResolvedTeamId $resolvedTeamId -ChannelId $channelId -CatalogAppId $catalogAppId
        }
        'shared' {
            Log "[InstallCheck] Shared Channel – App im Channel installieren"
            Get-TeamsAppInstalledForChannel -ResolvedTeamId $resolvedTeamId -ChannelId $channelId -CatalogAppId $catalogAppId
        }
        default { throw "Unbekannter membershipType: $($channelInfo.membershipType)" }
    }

    Log "Wait-TeamsAppReady -TeamId $resolvedTeamId -CatalogAppId $catalogAppId -TimeoutSeconds 20 ..."
    Wait-TeamsAppReady -TeamId $resolvedTeamId -CatalogAppId $catalogAppId -TimeoutSeconds 20

    $tabParams = @{
        TeamId         = $resolvedTeamId
        ChannelId      = $channelId
        TabDisplayName = $TabDisplayName
        TeamsAppId     = $catalogAppId
        EntityId       = $EntityId
        ContentUrl     = $ContentUrl
        WebsiteUrl     = $WebsiteUrl
    }
    Log "Add-GraphTeamsTab ..."
    Add-GraphTeamsTab @tabParams

    Log "Wait-TeamsTabVisible -TeamId $resolvedTeamId -ChannelId $channelId -TabDisplayName $TabDisplayName -TimeoutSeconds 45 -IntervalSeconds 3"
    Wait-TeamsTabVisible -TeamId $resolvedTeamId -ChannelId $channelId -TabDisplayName $TabDisplayName -TimeoutSeconds 45 -IntervalSeconds 3

    Log "✅ Tab '$TabDisplayName' im Channel '$channelNameResolved' erstellt."

    $result = @{ ok = $true; teamId = $resolvedTeamId; channelId = $channelId; tab = $TabDisplayName }
    Send-HttpResponse 200 $result
}
catch {
    $err = "❌ Fehler: $($_.Exception.Message)"
    Log $err
    Send-HttpResponse 500 @{ ok = $false; error = $err }
}
