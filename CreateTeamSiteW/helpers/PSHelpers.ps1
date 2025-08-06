# --------------------------------------------------------------------
# Include Helper Functions
# --------------------------------------------------------------------
. (Join-Path $PSScriptRoot 'LoggingFunctions.ps1')

# --------------------------------------------------------------------
# Evaluiere Request Parameter der Azure Function App
# --------------------------------------------------------------------
function EvaluateRequestParameters {
    param (
        [Parameter(Mandatory=$true)]
        $Request,

        [Parameter(Mandatory=$true)]
        [string[]]$RequiredParams,

        [string[]]$BooleanParams = @(),

        [object[]]$OptionalParams = @()
    )

    $params = @{}

    Log "Starting to evaluate request parameters..."

    # Query-Parameter sammeln
    #if ($Request -and $Request.Query) {
    #    foreach ($key in $Request.Query.Keys) {
    #        $params[$key] = $Request.Query[$key]
    #    }
    #}

    # Body-Parameter sammeln (ohne doppeltes ConvertFrom-Json, siehe vorherige Antworten!)
    $body = $Request.Body
    if ($body -is [string]) { $body = $body | ConvertFrom-Json }
    if ($body) {
        foreach ($key in $body.Keys) {
            $params[$key] = $body[$key]
            Log "Parameter: '$key' = '$($params[$key])'"
        }
    }

    # Pflichtfelder prüfen
    foreach ($p in $RequiredParams) {
        if (-not $params.ContainsKey($p) -or $null -eq $params[$p] -or $params[$p] -eq "") {
            Send-Resp 400 @{ error = "Missing required field: $p" }
            throw "Missing required field: $p"
        }
    }

    # Boolean-Felder prüfen
    foreach ($bp in $BooleanParams) {
        if ($params.ContainsKey($bp)) {
            $val = $params[$bp]
            if (($val -ne $true) -and ($val -ne $false)) {
                Send-Resp 400 @{ error = "$bp must be Boolean true or false, not '$val'" }
                throw "$bp must be Boolean true or false, not '$val'"
            }
        } else {
            # Wenn nicht angegeben, dann auf true setzen
            $params[$bp] = $true
        }
    }

    # Optional: Optionale Felder auf Default setzen, wenn nicht übergeben
    foreach ($op in $OptionalParams) {
        $name = $op.Name
        $default = $op.Default
        Log "Process optional parameter: '$name' ..."
        if (-not $params.ContainsKey($name) -or $params[$name] -eq "" -or $null -eq $params[$name]) {
            $params[$name] = $default
            Log "Use default Value: '$name' = '$($params[$name])'"
        }
    }

    return $params
}

# --------------------------------------------------------------------
# Teste Schema der Ordnerstruktur auf Korrektheit
# --------------------------------------------------------------------
function Test-Schema {
    param (
        $structure, # Kann Objekt, Array oder String sein (KEIN [string] erzwingen!)
        [string]$schema = @'
{
  "$schema": "http://json-schema.org/draft-07/schema#",
  "type": "array",
  "items": {
    "type": "object",
    "required": [ "name" ],
    "properties": {
      "name": { "type": "string" },
      "children": {
        "type": "array",
        "items": {
          "$ref": "#/items"
        }
      }
    },
    "additionalProperties": false
  }
}
'@
    )

    Log "Starting to evaluate Structure JSON..."

    # Wenn es ein String ist, zuerst als JSON interpretieren
    if ($structure -is [string]) {
        try {
            $structure = $structure | ConvertFrom-Json
        } catch {
            Send-Resp 400 @{ error = "Structure is not valid JSON." }
            throw "Structure is not valid JSON."
        }
    }

    # Immer als Array serialisieren, selbst wenn Einzelobjekt (für Schema!)
    if ($structure -isnot [System.Collections.IEnumerable] -or $structure -is [string]) {
        # Falls doch kein Array, mache es zu einem!
        $structure = @($structure)
    }

    $structureJson = $structure | ConvertTo-Json -Depth 10

    Log "structureJson = $structureJson"

    if (-not (Test-Json -Json $structureJson -Schema $schema)) {
        Send-Resp 400 @{ error = "Invalid structure format. Must follow folder schema with 'name' and optional 'children'." }
        throw "Invalid structure format. Must follow folder schema with 'name' and optional 'children'."
    } else {
        Log "Structure JSON is valid."
    }
}

# ----------------------------------------------------------------
# OLD: Teams Tab anlegen (zB AI Agent oder Websuche oder Sharepoint Site)
# ----------------------------------------------------------------
function AddTeamsTab {
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

# ------------------------------------------------------------------------
# Helper: Teams Tab über Microsoft Graph API anlegen (z. B. Copilot Bot)
# ------------------------------------------------------------------------
function Add-GraphTeamsTab {
    param(
        [Parameter(Mandatory=$true)]   [string]$TeamId,
        [Parameter(Mandatory=$true)]   [string]$ChannelId,
        [Parameter(Mandatory=$true)]   [string]$TabDisplayName,
        [Parameter(Mandatory=$true)]   [string]$TeamsAppId,    # App-ID der zugehörigen Teams-App
        [Parameter(Mandatory=$true)]   [string]$ContentUrl,    # Content URL wie im App-Manifest
        [Parameter(Mandatory=$false)]  [string]$EntityId = "copilot", # i. d. R. „copilot“ für Copilot Studio Bots
        [Parameter(Mandatory=$false)]  [string]$WebsiteUrl = $null    # Optional: zusätzlicher Website-Link
    )

    Log "📎 Lege Tab '$TabDisplayName' im Channel $ChannelId (Team: $TeamId) an..."

    $tabConfig = @{
        displayName = $TabDisplayName
        "teamsApp@odata.bind" = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps('$TeamsAppId')"
        configuration = @{}
    }

    # Automatische Konfiguration je nach AppId
    if ($TeamsAppId -eq "com.microsoft.teamspace.tab.web") {
        # Website-Tab: Alle drei Felder müssen gesetzt sein!
        $tabConfig.configuration.contentUrl  = $ContentUrl
        $tabConfig.configuration.websiteUrl  = $ContentUrl
        $tabConfig.configuration.entityId    = $ContentUrl  # Microsoft empfiehlt: URL als entityId für Website-Tabs!
    } else {
        # Bot/Copilot etc.: Nur das Nötigste
        $tabConfig.configuration.entityId    = $EntityId
        $tabConfig.configuration.contentUrl  = $ContentUrl
        if ($WebsiteUrl) { $tabConfig.configuration.websiteUrl = $WebsiteUrl }
    }

    try {
        $result = Invoke-MgGraphRequest -Method POST `
            -Uri "https://graph.microsoft.com/v1.0/teams/$TeamId/channels/$ChannelId/tabs" `
            -Body ($tabConfig | ConvertTo-Json -Depth 10)
        Log "✅ Tab '$TabDisplayName' erfolgreich angelegt!"
        return $result
    } catch {
        ErrorExit "❌ Fehler beim Erstellen des Tabs: $_"
    }
}

# ------------------------------------------------------------------------