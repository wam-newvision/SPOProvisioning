# ===============================
# TeamsTab.Core.ps1 (REST only)
# ===============================
# Hinweis: Diese Datei definiert KEINE Install-Funktionen mehr.
# Verwende die Implementierungen aus GraphRestHelpers.ps1:
#   - Get-TeamsAppInstalled
#   - Get-TeamsAppInstalledForChannel
#   - Invoke-Graph / Invoke-MgGraphRequest
#
# Optionaler Standalone-Import (auskommentiert lassen, wenn run.ps1 bereits lädt):
# $functionRoot = Split-Path -Parent $PSScriptRoot
# $helpersDir   = Join-Path $functionRoot 'Helpers'
# . (Join-Path $helpersDir 'GraphRestHelpers.ps1')

# ---------------------------------------------
# Team-Resolver: URL/GUID/Name -> groupId (Team)
# ---------------------------------------------

# -------- Helpers laden --------
$functionRoot = Split-Path -Parent $PSScriptRoot
$helpersDir   = Join-Path $functionRoot 'Helpers'

. (Join-Path $helpersDir 'GraphRestHelpers.ps1') # <— REST einbinden


function Resolve-TeamId {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$TeamRef)

    if ([string]::IsNullOrWhiteSpace($TeamRef)) { throw "TeamId/Name/Alias wurde nicht übergeben." }
    $TeamRefTrim = $TeamRef.Trim()

    # 0) Teams-URL / Deep Link? -> groupId extrahieren
    $m = [regex]::Match($TeamRefTrim, '(?i)(groupId=|/group/)([0-9a-f-]{36})')
    if ($m.Success) { return $m.Groups[2].Value }

    # 1) GUID? -> direkt zurück
    $g = [ref]([guid]::Empty)
    if ([guid]::TryParse($TeamRefTrim, $g)) { return $TeamRefTrim }

    # 2) Exakt auf displayName / mailNickname
    $needle = $TeamRefTrim.Replace("'", "''")
    $uriEq  = "https://graph.microsoft.com/v1.0/groups?`$filter=resourceProvisioningOptions/Any(x:x eq 'Team') and (displayName eq '$needle' or mailNickname eq '$needle')&`$select=id,displayName,mailNickname"
    $resEq  = Invoke-MgGraphRequest -Method GET -Uri $uriEq
    if ($resEq.value -and $resEq.value.Count -eq 1) { return $resEq.value[0].id }
    if ($resEq.value -and $resEq.value.Count -gt 1) {
        $list = ($resEq.value | ForEach-Object { "$($_.displayName) [$($_.mailNickname)] = $($_.id)" }) -join "; "
        throw "Mehrere Teams gefunden für '$TeamRefTrim': $list"
    }

    # 3) startsWith auf displayName / mailNickname
    $uriSw = "https://graph.microsoft.com/v1.0/groups?`$filter=resourceProvisioningOptions/Any(x:x eq 'Team') and (startsWith(displayName,'$needle') or startsWith(mailNickname,'$needle'))&`$select=id,displayName,mailNickname"
    $resSw = Invoke-MgGraphRequest -Method GET -Uri $uriSw
    if ($resSw.value -and $resSw.value.Count -eq 1) { return $resSw.value[0].id }
    if ($resSw.value -and $resSw.value.Count -gt 1) {
        $best = $resSw.value | Where-Object { $_.displayName -ieq $TeamRefTrim -or $_.mailNickname -ieq $TeamRefTrim } | Select-Object -First 1
        if ($best) { return $best.id }
        $list = ($resSw.value | Select-Object -First 5 | ForEach-Object { "$($_.displayName) [$($_.mailNickname)] = $($_.id)" }) -join "; "
        throw "Mehrere Kandidaten gefunden für '$TeamRefTrim': $list"
    }

    throw "Kein Team gefunden für '$TeamRefTrim'."
}

# ---------------------------------------------
# Channel-Resolver (Standard = "General")
# ---------------------------------------------
function Get-ChannelId {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ResolvedTeamId,
        [string]$ChannelName
    )
    if ([string]::IsNullOrWhiteSpace($ChannelName)) { $ChannelName = 'General' }

    $uri = "https://graph.microsoft.com/v1.0/teams/$ResolvedTeamId/channels?`$select=id,displayName,membershipType"
    $res = Invoke-MgGraphRequest -Method GET -Uri $uri
    if (-not $res.value) { throw "Keine Channels im Team $ResolvedTeamId gefunden." }

    $exact = $res.value | Where-Object { $_.displayName -ieq $ChannelName } | Select-Object -First 1
    if ($exact) { 
        if (Get-Command Log -ErrorAction SilentlyContinue) { Log "[Get-ChannelId] ℹ️ Verwende Kanal: $($exact.displayName) (ID: $($exact.id))" }
        return @($exact.id, $exact.displayName) 
    }

    $start = $res.value | Where-Object { $_.displayName -like "$ChannelName*" } | Select-Object -First 1
    if ($start) { 
        if (Get-Command Log -ErrorAction SilentlyContinue) { Log "[Get-ChannelId] ℹ️ Verwende Kanal (startsWith): $($start.displayName) (ID: $($start.id))" }
        return @($start.id, $start.displayName) 
    }

    $list = ($res.value | Select-Object -First 5 | ForEach-Object { "$($_.displayName) [$($_.id)]" }) -join "; "
    throw "Channel '$ChannelName' nicht gefunden. Kandidaten: $list"
}

# ---------------------------------------------
# Warten bis ein Tab sichtbar ist (Polling)
# ---------------------------------------------
function Wait-TeamsTabVisible {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TeamId,
        [Parameter(Mandatory)][string]$ChannelId,
        [Parameter(Mandatory)][string]$TabDisplayName,
        [int]$TimeoutSeconds = 45,
        [int]$IntervalSeconds = 3
    )

    $end = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        $uri = "https://graph.microsoft.com/v1.0/teams/$TeamId/channels/$ChannelId/tabs?`$select=id,displayName"
        $tabs = Invoke-MgGraphRequest -Method GET -Uri $uri
        if ($tabs.value -and ($tabs.value.displayName -contains $TabDisplayName)) { return $true }
        Start-Sleep -Seconds $IntervalSeconds
    } while ((Get-Date) -lt $end)

    throw "Tab '$TabDisplayName' wurde nicht innerhalb von $TimeoutSeconds Sekunden sichtbar."
}
