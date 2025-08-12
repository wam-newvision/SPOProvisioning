# --------------------------------------------------------------------
# Include Helper Functions
# --------------------------------------------------------------------
$PSModuleAutoloadingPreference = 'None'
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
# --------------------------------------------------------------------
function LoadGraphModule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ModuleName,     # z.B. 'Microsoft.Graph.Teams'
        [Parameter(Mandatory)][string]$FktPath,        # Pfad auf den 'modules'-Ordner
        [Parameter(Mandatory)][string]$GraphVersion    # z.B. '2.29.1'
    )

    $ErrorActionPreference = 'Stop'

    # Erwartete Struktur: <FktPath>\<ModuleName>\<GraphVersion>\<ModuleName>.psd1
    $psd1 = Join-Path $FktPath (Join-Path $ModuleName (Join-Path $GraphVersion "$ModuleName.psd1"))

    if (-not (Test-Path $psd1)) {
        throw "❌ Modulmanifest nicht gefunden: $psd1  (erwartet unter '$FktPath\$ModuleName\$GraphVersion\')"
    }

    Import-Module $psd1 -Force -ErrorAction Stop
    Log "✅ Modul '$ModuleName' geladen aus $($psd1 | Split-Path -Parent)"
}
# --------------------------------------------------------------------