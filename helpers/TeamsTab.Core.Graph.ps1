# helpers/TeamsTab.Core.ps1
# Kern-Helfer für Teams-Tab-Provisionierung (PnP-frei)

Set-StrictMode -Version Latest
$env:DEBUG = 'true'

function Get-ChannelId {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ResolvedTeamId,
        [string]$ChannelName
    )

    if ([string]::IsNullOrWhiteSpace($ChannelName)) {
        $primary = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/teams/$ResolvedTeamId/primaryChannel"
        if (-not $primary -or -not $primary.id) { throw "primaryChannel konnte nicht ermittelt werden." }
        Log "ℹ️ Verwende Standard-Kanal: $($primary.displayName) (ID: $($primary.id))"
        return $primary.id, $primary.displayName
    } else {
        $channels = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/teams/$ResolvedTeamId/channels"
        $ch = $channels.value | Where-Object { $_.displayName -ieq $ChannelName } | Select-Object -First 1
        if (-not $ch) { throw "Channel '$ChannelName' nicht gefunden in Team $ResolvedTeamId." }
        Log "ℹ️ Verwende Kanal: $($ch.displayName) (ID: $($ch.id))"
        return $ch.id, $ch.displayName
    }
}

function Get-CatalogAppId {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ExternalId)

    $apps = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps?`$filter=externalId eq '$ExternalId'&`$select=id,displayName,externalId,distributionMethod"
    if (-not $apps.value -or -not $apps.value[0].id) { throw "Custom-App (externalId=$ExternalId) nicht im App-Katalog gefunden oder blockiert." }
    if ($apps.value[0].distributionMethod -ne 'organization') { Log "⚠️ Hinweis: distributionMethod = $($apps.value[0].distributionMethod)" }
    return $apps.value[0].id
}

function Get-TeamsAppInstalled {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ResolvedTeamId,
        [Parameter(Mandatory)][string]$CatalogAppId
    )

    $installed = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/teams/$ResolvedTeamId/installedApps?`$expand=teamsApp&`$select=id,teamsApp"
    if (@($installed.value | Where-Object { $_.teamsApp.id -eq $CatalogAppId }).Count -gt 0) {
        Log "[Ensure-TeamsAppInstalled] ℹ️ App bereits im Team installiert."
        return
    }

    $body = @{ "teamsApp@odata.bind" = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/$CatalogAppId" }
    $max = 5
    for ($i=1; $i -le $max; $i++) {
        try {
            Log "[Ensure-TeamsAppInstalled] ℹ️ Custom-App installieren in Team $ResolvedTeamId (Versuch $i/$max)…"
            Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/teams/$ResolvedTeamId/installedApps" -Body $body

        } catch {
            $msg = $_.Exception.Message
            # 409 = bereits installiert -> als Erfolg behandeln
            if ($msg -match '(?i)\b409\b' -or $msg -match 'AppEntitlementAlreadyExists' -or $msg -match 'Conflict') {
                Log "[Ensure-TeamsAppInstalled] ✅ 409/Conflict: App ist bereits installiert."
                break
            }
            # sporadische Backend-Fehler -> retry
            if ($msg -match 'BulkMembershipS2SRequest' -or $msg -match 'Skype backend' -or $msg -match 'BadRequest') {
                Start-Sleep -Seconds (2 * $i)
                continue
            }
            throw
        }

        # Poll bis sichtbar
        for ($j=1; $j -le 10; $j++) {
            Start-Sleep -Seconds 2
            $installed = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/teams/$ResolvedTeamId/installedApps?`$expand=teamsApp&`$select=id,teamsApp"
            if (@($installed.value | Where-Object { $_.teamsApp.id -eq $CatalogAppId }).Count -gt 0) {
                Log "[Ensure-TeamsAppInstalled] ✅ App im Team installiert."
                return
            }
        }
    }

    # finaler Check – falls POST 409 war, sollte sie jetzt sichtbar sein:
    $installed = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/teams/$ResolvedTeamId/installedApps?`$expand=teamsApp&`$select=id,teamsApp"
    if (@($installed.value | Where-Object { $_.teamsApp.id -eq $CatalogAppId }).Count -gt 0) {
        Log "[Ensure-TeamsAppInstalled] ✅ App im Team installiert (nach 409)."
        return
    }

    throw "App ließ sich nicht verifizieren (CatalogId=$CatalogAppId)."
}

function Get-TeamsAppInstalledForChannel {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ResolvedTeamId,
        [Parameter(Mandatory)][string]$ChannelId,
        [Parameter(Mandatory)][string]$CatalogAppId
    )

    $list = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/teams/$ResolvedTeamId/channels/$ChannelId/installedApps?`$expand=teamsApp&`$select=id,teamsApp"
    if (@($list.value | Where-Object { $_.teamsApp.id -eq $CatalogAppId }).Count -gt 0) {
        Log "[Ensure-TeamsAppInstalledForChannel] ℹ️ App bereits im Channel installiert."
        return
    }

    $body = @{ "teamsApp@odata.bind" = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/$CatalogAppId" }
    try {
        Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/teams/$ResolvedTeamId/channels/$ChannelId/installedApps" -Body $body
    } catch {
        $msg = $_.Exception.Message
        if ($msg -match '(?i)\b409\b' -or $msg -match 'AppEntitlementAlreadyExists' -or $msg -match 'Conflict') {
            Log "[Ensure-TeamsAppInstalledForChannel] ✅ 409/Conflict: App ist bereits im Channel installiert."
        } else { throw }
    }
}

# ------------------------------------------------------------------------
# Teams Tab über Microsoft Graph API anlegen (z. B. Copilot Bot)
# ------------------------------------------------------------------------
function Add-GraphTeamsTab {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]  [string]$TeamId,
        [Parameter(Mandatory=$true)]  [string]$ChannelId,
        [Parameter(Mandatory=$true)]  [string]$TabDisplayName,
        [Parameter(Mandatory=$true)]  [string]$TeamsAppId,      # Catalog-ID ODER well-known Id (z.B. com.microsoft.teamspace.tab.web)
        [string]$EntityId = "copilot",
        [string]$ContentUrl,                                     # optional -> wird ggf. aus WebsiteUrl genommen
        [string]$WebsiteUrl                                      # optional -> wird ggf. aus ContentUrl genommen
    )

    # 1) Ziel-URL bestimmen (wir sorgen dafür, dass mind. eine URL vorhanden ist)
    $isWebsiteApp = $TeamsAppId -like 'com.microsoft.teamspace.tab.*'

    Log "ContentUrl: $ContentUrl"

    if ($isWebsiteApp) {
        # Für Website-Tab dürfen ContentUrl/WebsiteUrl identisch sein – mind. eine muss gesetzt sein
        if (-not $ContentUrl -and -not $WebsiteUrl) {
            throw "Für Website-Tabs (TeamsAppId=$TeamsAppId) muss mindestens eine URL angegeben werden: -ContentUrl oder -WebsiteUrl."
        }
        if (-not $ContentUrl) { $ContentUrl = $WebsiteUrl }
        if (-not $WebsiteUrl) { $WebsiteUrl = $ContentUrl }
        # Empfehlung von Microsoft: entityId = URL
        $entity = $ContentUrl
    }
    else {
        # Custom App (Catalog-ID): hier erwarten viele Apps eine ContentUrl; falls nicht gesetzt, auf WebsiteUrl zurückfallen
        if (-not $ContentUrl -and -not $WebsiteUrl) {
            throw "Für Custom Tabs (TeamsAppId=$TeamsAppId) muss mindestens eine URL angegeben werden: -ContentUrl oder -WebsiteUrl."
        }
        if (-not $ContentUrl) { $ContentUrl = $WebsiteUrl }
        if (-not $WebsiteUrl) { $WebsiteUrl = $ContentUrl }
        $entity = $EntityId
    }

    Log "📎 Lege Tab '$TabDisplayName' im Channel $ChannelId (Team: $TeamId) an..."

    # 2) Korrektes OData-Binding bauen (ohne Klammern/Quotes um die ID)
    $bindUrl =
        if ($TeamsAppId -like 'com.microsoft.*') {
            "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/$TeamsAppId"
        } else {
            # wir nehmen an: GUID = Catalog-ID
            "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/$TeamsAppId"
        }

    $tabConfig = @{
        displayName = $TabDisplayName
        "teamsApp@odata.bind" = $bindUrl
        configuration = @{
            entityId   = $entity
            contentUrl = $ContentUrl
            websiteUrl = $WebsiteUrl
        }
    }

    $GraphURI = "https://graph.microsoft.com/v1.0/teams/$TeamId/channels/$ChannelId/tabs"
    Log "Graph URI: $GraphURI"

    try {
        # Invoke-MgGraphRequest akzeptiert Hashtables als -Body; kein explizites ConvertTo-Json nötig
        $result = Invoke-MgGraphRequest -Method POST -Uri $GraphURI -Body $tabConfig
        Log "✅ Tab '$TabDisplayName' erfolgreich angelegt!"
        return $result
    } catch {
        ErrorExit "❌ Fehler beim Erstellen des Tabs: $_"
    }
}

function Wait-TeamsAppReady {
    param (
        [string]$TeamId,
        [string]$CatalogAppId,
        [int]$TimeoutSeconds = 20
    )

    Log "⏳ Warte auf vollständige App-Installation..."
    $stopWatch = [Diagnostics.Stopwatch]::StartNew()

    while ($stopWatch.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        $app = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/teams/$TeamId/installedApps?`$expand=teamsApp,teamsAppDefinition" `
            | Select-Object -ExpandProperty value `
            | Where-Object { $_.teamsApp.id -eq $CatalogAppId }

        if ($app -and $app.teamsAppDefinition -and $app.teamsAppDefinition.id) {
            Log "✅ App-Definition gefunden: $($app.teamsAppDefinition.id)"
            return $true
        }

        Start-Sleep -Seconds 2
    }

    Log "⚠️ App war nach $TimeoutSeconds Sekunden nicht vollständig geladen."
    return $false
}

# --- Hilfsfunktionen lokal, falls nicht global verfügbar ---
function Test-IsGuid { param([string]$Value) return [Guid]::TryParse($Value, [ref]([Guid]::Empty)) }

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

# ------------------------------------------------------------------------
# Ersetzt den bisherigen Paginator
function Invoke-GraphPagedGet {
    param([Parameter(Mandatory)][string]$Uri)

    $items = @()
    do {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $Uri -ErrorAction Stop

        # Nur sammeln, wenn ein echtes "value":[...] vorhanden ist
        if ($resp -and $resp.PSObject.Properties.Name -contains 'value' -and $resp.value) {
            # Stelle sicher, dass wir ein Array bekommen
            if ($resp.value -is [System.Array]) {
                $items += $resp.value
            } else {
                $items += @($resp.value)
            }
        } else {
            # Kein "value" => nichts hinzufügen (Tabs-Endpunkt sollte immer collections liefern)
            Log "Invoke-GraphPagedGet: Response ohne 'value' ignoriert. Keys: $(@($resp.PSObject.Properties.Name) -join ', ')"
        }

        $next = $null
        if ($resp -and $resp.PSObject.Properties.Name -contains '@odata.nextLink') {
            $next = $resp.'@odata.nextLink'
        }
        $Uri = $next
    } while ($Uri)

    return $items
}

# ------------------------------------------------------------------------
function Wait-TeamsTabVisible {
    <#
      Wartet bis ein Tab mit diesem Namen (und optional passender App) in /tabs auftaucht.
      Gibt die TabId zurück oder $null nach Timeout.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TeamId,            # GUID, Anzeigename oder Deep-Link (groupId=...)
        [Parameter(Mandatory)][string]$ChannelId,         # GUID, 19:...@thread.tacv2 ODER Channel-Name
        [Parameter(Mandatory)][string]$TabDisplayName,    # Tab-Anzeigename
        [string]$CatalogAppId,                            # optional: App-Katalog-ID für exakten Match
        [int]$TimeoutSeconds = 45,
        [int]$IntervalSeconds = 3
    )

    # Lokaler Logger: nutzt vorhandenes Log-Cmdlet, fällt sonst auf Write-Host zurück
    function _Log { param([string]$Message)
        if (Get-Command -Name Log -ErrorAction SilentlyContinue) { Log $Message } else { Write-Host $Message }
    }

    $resolvedTeamId = Resolve-TeamId -TeamRef $TeamId
    $targetName     = $TabDisplayName.Trim()
    $deadline       = (Get-Date).AddSeconds($TimeoutSeconds)
    $maxAttempts    = [Math]::Max(1, [Math]::Ceiling($TimeoutSeconds / [Math]::Max(1, $IntervalSeconds)))

    # --- Channel-IDs ermitteln (GUID direkt; "19:" oder Name -> Kanäle auflisten & bestimmen/alle prüfen) ---
    $channelIds = @()
    $asGuid = [ref]([guid]::Empty)
    if ([guid]::TryParse(($ChannelId.Trim()), $asGuid)) {
        $channelIds = @($ChannelId.Trim())
    } else {
        $allCh = @( Get-MgTeamChannel -TeamId $resolvedTeamId -All )
        if ($ChannelId -match '^(?i)19:') {
            # Mapping 19: -> GUID nicht zuverlässig verfügbar => alle Kanäle prüfen
            $channelIds = @($allCh | ForEach-Object { $_.Id })
        } else {
            $exact = $allCh | Where-Object { $_.displayName -ieq $ChannelId.Trim() }
            if ($exact) { $channelIds = @($exact[0].Id) }
            if (-not $channelIds -or (@($channelIds).Count -eq 0)) {
                $starts = $allCh | Where-Object { $_.displayName -like ($ChannelId.Trim() + '*') }
                if ($starts) { $channelIds = @($starts[0].Id) }
            }
            if (-not $channelIds -or (@($channelIds).Count -eq 0)) {
                # letzte Rettung: alle Kanäle prüfen
                $channelIds = @($allCh | ForEach-Object { $_.Id })
            }
        }
    }

    $attempt = 0
    do {
        $attempt++
        $totalCandidates = 0

        foreach ($cid in $channelIds) {
            # 1) SDK zuerst
            $tabs = @()
            try {
                $tabs = @( Get-MgTeamChannelTab -TeamId $resolvedTeamId -ChannelId $cid -All )
            } catch { $tabs = @() }

            # 2) Fallback: v1.0
            if ((@($tabs) | Measure-Object).Count -eq 0) {
                try {
                    $u1 = "https://graph.microsoft.com/v1.0/teams/$resolvedTeamId/channels/$cid/tabs"
                    $r1 = Invoke-MgGraphRequest -Method GET -Uri $u1 -ErrorAction Stop
                    if ($r1 -and $r1.PSObject.Properties.Name -contains 'value' -and $r1.value) {
                        $tabs = @($r1.value)
                    }
                } catch { }
            }
            # 3) Fallback: beta
            if ((@($tabs) | Measure-Object).Count -eq 0) {
                try {
                    $u2 = "https://graph.microsoft.com/beta/teams/$resolvedTeamId/channels/$cid/tabs?`$expand=teamsApp"
                    $r2 = Invoke-MgGraphRequest -Method GET -Uri $u2 -ErrorAction Stop
                    if ($r2 -and $r2.PSObject.Properties.Name -contains 'value' -and $r2.value) {
                        $tabs = @($r2.value)
                    }
                } catch { }
            }

            # Nur valide Tab-Objekte berücksichtigen
            $candidateTabs = $tabs | Where-Object {
                $_ -and ($_.PSObject.Properties.Name -contains 'displayName') -and
                -not [string]::IsNullOrWhiteSpace([string]$_.displayName)
            }
            $totalCandidates += (@($candidateTabs) | Measure-Object).Count

            # Match auf Name (case-insensitive)
            $tabByName = $candidateTabs | Where-Object {
                ($_.displayName.ToString().Trim() -ieq $targetName)
            } | Select-Object -First 1

            if ($tabByName) {
                # AppId prüfen – aber NICHT blockieren, nur informieren
                $foundAppId = $null
                if ($tabByName.PSObject.Properties.Name -contains 'teamsAppId') {
                    $foundAppId = $tabByName.teamsAppId
                } elseif ($tabByName.PSObject.Properties.Name -contains 'teamsApp' -and $tabByName.teamsApp) {
                    $foundAppId = $tabByName.teamsApp.id
                }
                if ($CatalogAppId -and $foundAppId -and ($foundAppId -ne $CatalogAppId)) {
                    _Log "[Wait-TeamsTabVisible] ⚠️ Name match, aber AppId abweichend: gefunden=$($foundAppId) erwartet=$($CatalogAppId) (Channel=$($cid))."
                }

                if ($tabByName.PSObject.Properties.Name -contains 'id' -and $tabByName.id) {
                    _Log "[Wait-TeamsTabVisible] ✅ Tab '$targetName' gefunden (Channel=$($cid), Versuch ${attempt}/${maxAttempts}): $($tabByName.id)"
                    return $tabByName.id
                }
            }
        }

        _Log "[Wait-TeamsTabVisible] ⏳ Versuch ${attempt}/${maxAttempts}: '$targetName' noch nicht in Graph gelistet. Kandidaten gesamt: $($totalCandidates) über $((@($channelIds) | Measure-Object).Count) Channel(s)."
        Start-Sleep -Seconds $IntervalSeconds
    } while ((Get-Date) -lt $deadline)

    _Log "[Wait-TeamsTabVisible] ❌ Tab '$targetName' nach $TimeoutSeconds s nicht sichtbar."
    return $null
}

# ------------------------------------------------------------------------
function New-TeamsTabDeepLink {
    <#
      .SYNOPSIS
        Baut einen Deep Link zu einem (neu angelegten) Teams-Tab.

      .PARAMETER AppId
        Catalog-App-ID (GUID) ODER well-known AppId wie 'com.microsoft.teamspace.tab.web'.

      .PARAMETER EntityId
        Website-Tab: üblicherweise die URL; Custom-Tab: dein Entity-Key (z. B. "home").
        (Hinweis: Deine Add-GraphTeamsTab setzt bei Website-Tabs entityId = ContentUrl.)

      .PARAMETER ContentUrl
        Die Content-URL des Tabs (wird im Link als webUrl verwendet).

      .PARAMETER TabName
        Anzeigename des Tabs (Label im Link).

      .PARAMETER TeamId
        GUID des Teams.

      .PARAMETER ChannelId
        ID des Channels (z. B. '19:...@thread.tacv2').

      .OUTPUTS
        [string] – klickbarer Deep Link zu genau diesem Tab.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AppId,
        [Parameter(Mandatory)][string]$EntityId,
        [Parameter(Mandatory)][string]$ContentUrl,
        [Parameter(Mandatory)][string]$TabName,
        [Parameter(Mandatory)][string]$TeamId,
        [Parameter(Mandatory)][string]$ChannelId
    )

    $base = "https://teams.microsoft.com/l/entity/{0}/{1}" -f `
            [uri]::EscapeDataString($AppId), [uri]::EscapeDataString($EntityId)

    $ctx  = @{ channelId = $ChannelId } | ConvertTo-Json -Compress

    $deeplink = "{0}?webUrl={1}&label={2}&context={3}&groupId={4}" -f `
                $base,
                [uri]::EscapeDataString($ContentUrl),
                [uri]::EscapeDataString($TabName),
                [uri]::EscapeDataString($ctx),
                $TeamId

    return $deeplink
}

function Send-TeamsChannelMessage {
    <#
      .SYNOPSIS
        Sendet eine HTML-Nachricht in einen Teams-Channel.

      .PARAMETER TeamId
        GUID des Teams.

      .PARAMETER ChannelId
        Channel-ID (z. B. '19:...@thread.tacv2').

      .PARAMETER Html
        HTML-Inhalt der Nachricht (contentType = "html").

      .OUTPUTS
        Antwortobjekt der Graph-API (inkl. message id).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TeamId,
        [Parameter(Mandatory)][string]$ChannelId,
        [Parameter(Mandatory)][string]$Html
    )

    # Hinweis: Für diesen Call brauchst du (delegated) u.a. "ChannelMessage.Send".
    $body = @{
        body = @{
            contentType = "html"
            content     = $Html
        }
    }

    $uri = "https://graph.microsoft.com/v1.0/teams/$TeamId/channels/$ChannelId/messages"
    return Invoke-MgGraphRequest -Method POST -Uri $uri -Body $body
}

function Get-TeamsTabId {
    <#
      .SYNOPSIS
        Liefert die Tab-ID in einem Channel anhand des Anzeigenamens (case-insensitive).

      .OUTPUTS
        [string] Tab-Id oder $null
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TeamId,
        [Parameter(Mandatory)][string]$ChannelId,
        [Parameter(Mandatory)][string]$TabDisplayName
    )

    $tabs = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/teams/$TeamId/channels/$ChannelId/tabs"
    $tab  = $tabs.value | Where-Object { $_.displayName -ieq $TabDisplayName } | Select-Object -First 1
    if ($tab) { return $tab.id }
    return $null
}

function Invoke-TeamsTabRefresh {
    <#
      .SYNOPSIS
        „Nudged“ einen konfigurierbaren Teams-Tab, indem der Anzeigename kurz geändert und zurück gesetzt wird.

      .DESCRIPTION
        - Holt (falls nötig) die Tab-ID.
        - PATCH → displayName = "$TabDisplayName " (mit Space)
        - PATCH → displayName = $TabDisplayName
        - Fängt typische Fehler ab:
          * 400 für statische Tabs (wird übersprungen)
          * 409/412 (ETag/Concurrency) → kurzer Retry
          * 429 (Throttling) → Backoff + Retry

      .OUTPUTS
        $true bei Erfolg, $false wenn übersprungen (z. B. static tab) oder Tab nicht gefunden.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TeamId,
        [Parameter(Mandatory)][string]$ChannelId,
        [Parameter(Mandatory)][string]$TabDisplayName,
        [string]$TabId
    )

    # Tab-ID ggf. ermitteln
    if (-not $TabId) {
        $TabId = Get-TeamsTabId -TeamId $TeamId -ChannelId $ChannelId -TabDisplayName $TabDisplayName
        if (-not $TabId) {
            Log "[Invoke-TeamsTabRefresh] ⚠️ Tab '$TabDisplayName' nicht gefunden – kein Refresh möglich."
            return $false
        }
    }

    $uri = "https://graph.microsoft.com/v1.0/teams/$TeamId/channels/$ChannelId/tabs/$TabId"
    $max = 3

    # kleine lokale Helper-Funktion für PATCH mit Retry
    function _DoPatch([hashtable]$body) {
        for ($i=1; $i -le $max; $i++) {
            try {
                return Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $body
            } catch {
                $msg = $_.Exception.Message

                # static tabs können nicht gePATCHt werden
                if ($msg -match '(?i)static tab' -or $msg -match '(?i)cannot update.*static') {
                    Log "[Invoke-TeamsTabRefresh] ℹ️ Static Tab erkannt – kein Refresh möglich."
                    return $null
                }

                # 429 Throttling → Backoff
                if ($msg -match '\b429\b' -or $msg -match '(?i)Too Many Requests') {
                    Start-Sleep -Seconds (2 * $i)
                    continue
                }

                # 409/412 Concurrency/Precondition → kurzer Retry
                if ($msg -match '\b409\b' -or $msg -match '\b412\b' -or $msg -match '(?i)(precondition|conflict)') {
                    Start-Sleep -Seconds (1 * $i)
                    continue
                }

                throw
            }
        }
        throw "PATCH ist nach $max Versuchen fehlgeschlagen."
    }

    # 1) kurz umbenennen
    $r1 = _DoPatch @{ displayName = "$TabDisplayName " }
    if ($null -eq $r1) { return $false }  # static → skipped

    Start-Sleep -Milliseconds 300

    # 2) wieder zurück
    $r2 = _DoPatch @{ displayName = $TabDisplayName }
    if ($null -eq $r2) { return $false }

    Log "[Invoke-TeamsTabRefresh] ✅ Tab '$TabDisplayName' wurde „genudged“."
    return $true
}

function Invoke-TeamsTabsRefreshTrigger {
    <#
      Legt ein temporäres Website-Tab an und löscht es sofort wieder.
      Nutzt die First-Party Website-App (kein Katalog nötig).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TeamId,
        [Parameter(Mandatory)][string]$ChannelId,
        [bool]$Delete = $true  # ob das Tab erstellt oder gelöscht werden soll
    )

    $display = "_RefreshTrigger"
    $bindUrl = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.web"
    $cfg = @{
        displayName = $display
        "teamsApp@odata.bind" = $bindUrl
        configuration = @{
            entityId   = "https://teams.microsoft.com"   # Empfehlung: URL als entityId
            contentUrl = "https://teams.microsoft.com"
            websiteUrl = "https://teams.microsoft.com"
        }
    }

    # Create
    $uriTabs = "https://graph.microsoft.com/v1.0/teams/$TeamId/channels/$ChannelId/tabs"

    Log "📎 Erstelle temporäres Website-Tab '$display' im Channel $ChannelId (Team: $TeamId)..."
    $created = Invoke-MgGraphRequest -Method POST -Uri $uriTabs -Body $cfg

    if ($Delete) {
        Log "🗑️ Lösche temporäres Website-Tab '$display' im Channel $ChannelId (Team: $TeamId)..."
        # Delete (best effort)
        try {
            $tid = $created.id
            if ($tid) {
                Invoke-MgGraphRequest -Method DELETE -Uri "$uriTabs/$tid"
            }
        } catch { }
    }       

    return $true
}

# -------------------------------------------------------------------------------
# OLD: Teams Tab per PnP anlegen (zB AI Agent oder Websuche oder Sharepoint Site)
# -------------------------------------------------------------------------------
<#
function Add-PnPTeamsTab {
    param (
        [Parameter(Mandatory=$true)]
        $team,  # PnP Teams Team Object

        [Parameter(Mandatory=$false)]
        $TeamsChannel  = $null,  # PnP Teams Channel Object (optional, wenn nicht angegeben, wird der Default Channel verwendet)

        [Parameter(Mandatory=$true)]
        $TabDisplayName,  # Name des Tabs

        [Parameter(Mandatory=$true)]
        $TabType,  # Typ des Tabs (z.B. SharePointPageAndList, WebSite, etc.)

        [Parameter(Mandatory=$true)]
        $WebSiteUrl,  # URL der Website, die im Tab angezeigt werden soll

        [Parameter(Mandatory=$false)]
        $WebSiteUrlDisplayName  # DisplayName der Website, die im Tab angezeigt werden soll
    )

    if($TeamsChannel) {
        $channels = $true
    } else {
        Log "Find Default Channel for Team '$($team.DisplayName)' ..."
        $maxTries = 20
        $waitSeconds = 3
        $channels = $null
        for ($i=1; $i -le $maxTries; $i++) {
            try {
                $channels = Get-PnPTeamsChannel -Team $team.GroupId -ErrorAction Stop
                if ($channels -and $channels.Count -gt 0) {
                    Log "✅ Es wurden $($channels.Count) Kanäle gefunden (nach $i Versuch(en))"
                    break
                }
            } catch {
                Log "⌛ Warte auf Channel-Verfügbarkeit (Versuch $i)..."
                Start-Sleep -Seconds $waitSeconds
            }
        }
    }

    if (-not $channels -or $channels.Count -eq 0) {
        throw "❌ Es konnten keine Channels für das Team $team gefunden werden!"
    }

    if($channels) {
        if ($TeamsChannel) {
            # Wenn ein spezifischer Channel angegeben wurde, nutze diesen
            $channel = $TeamsChannel
            Log "📢 Using specified Channel: $($channel.DisplayName) (ID: $($channel.Id))"
        } else {
            $channel = $channels | Select-Object -First 1
            Log "ℹ️ Using default Channel: $($channel.DisplayName) (ID: $($channel.Id))"
        }

        Log "Add AI-Tab to Team '$($team.DisplayName)' in Channel '$($channel.DisplayName)' ..."
        Add-PnPTeamsTab `
            -Team $team `
            -Channel $channel.Id `
            -DisplayName $TabDisplayName `
            -Type SharePointPageAndList `
            -WebSiteUrl $WebSiteUrl

            #-Type Website `
            #-ContentUrl "https://google.com"

    }
}
#>

# ------------------------------------------------------------------------