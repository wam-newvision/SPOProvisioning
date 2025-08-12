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
    param(
        [Parameter(Mandatory)][string]$ModuleName,
        [Parameter(Mandatory)][string]$ModulesRoot
    )
    if (-not (Test-Path $ModulesRoot)) {
        throw "Modules-Root nicht gefunden: $ModulesRoot"
    }
    $moduleFolder = Join-Path $ModulesRoot $ModuleName
    if (-not (Test-Path $moduleFolder)) {
        throw "Graph-Modul '$ModuleName' nicht gefunden unter $moduleFolder"
    }
    # Höchste Versions-Unterordner wählen (falls vorhanden)
    $versionDir = Get-ChildItem -Path $moduleFolder -Directory -ErrorAction SilentlyContinue |
                    Sort-Object Name -Descending | Select-Object -First 1
    $moduleBase = if ($versionDir) { $versionDir.FullName } else { $moduleFolder }
    $psd1 = Join-Path $moduleBase "$ModuleName.psd1"
    if (-not (Test-Path $psd1)) {
        throw "PSD1 für '$ModuleName' fehlt: $psd1"
    }
    Import-Module $psd1 -Force -ErrorAction Stop
}
# --------------------------------------------------------------------

function CallTeamsTab {
    param (
        [Parameter(Mandatory = $true)]
        [string]$tabScript,
        [Parameter(Mandatory = $true)]
        [string]$payloadFile
    )

    Log "[CallTeamsTab] Starting in isolated runspace ..."

    if (-not (Test-Path $tabScript))   { throw "Tab-Skript nicht gefunden: $tabScript" }
    if (-not (Test-Path $payloadFile)) { throw "Payload-Datei nicht gefunden: $payloadFile" }

    # AppRoot bestimmen (…\wwwroot)
    $scriptDir = Split-Path -Parent $tabScript
    $appRoot   = Split-Path -Parent $scriptDir   # wwwroot
    $modulesRoot = Join-Path $appRoot 'modules'

    # Isolierten Runspace erstellen
    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $iss.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new(
        'ErrorActionPreference', [System.Management.Automation.ActionPreference]::Stop, ''
    ))
    $iss.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new(
        'PSModuleAutoloadingPreference', 'None', ''
    ))

    $runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($iss)
    $runspace.ApartmentState = 'STA'
    $runspace.ThreadOptions  = 'ReuseThread'
    $runspace.Open()

    try {
        $ps = [powershell]::Create()
        $ps.Runspace = $runspace

        # Im Runspace: Loader-Funktion definieren, Module aus wwwroot\Modules importieren, dann TeamsTab.ps1 starten
        $ps.AddScript({
            param($tabScriptPath, $payloadPath, $modulesRootPath)

            # Graph-Module gezielt laden (keine PnP-Module!)
            foreach ($m in @('Microsoft.Graph.Authentication','Microsoft.Graph.Teams','Microsoft.Graph.Groups')) {
                LoadGraphModule -ModuleName $m -ModulesRoot $modulesRootPath
            }

            # TeamsTab.ps1 mit Payload-Datei ausführen
            & $tabScriptPath -payloadFile $payloadPath
        }).AddArgument($tabScript).AddArgument($payloadFile).AddArgument($modulesRoot) | Out-Null

        # Ausführen mit Timeout
        $async = $ps.BeginInvoke()
        $timeoutMs = 120000
        if (-not $async.AsyncWaitHandle.WaitOne($timeoutMs)) {
            try { $ps.Stop() } catch {}
            throw "Tab-Skript Timeout nach $($timeoutMs/1000)s."
        }

        $result = $ps.EndInvoke($async)
        if ($ps.HadErrors) {
            $errs = $ps.Streams.Error | ForEach-Object { $_.ToString() } -join "`n"
            throw "Tab-Skript im Runspace fehlgeschlagen: $errs"
        }

        if ($result) { Log ($result -join "`n") }
    }
    finally {
        $runspace.Close()
        $runspace.Dispose()
        Remove-Item -Path $payloadFile -ErrorAction SilentlyContinue
    }
}
# --------------------------------------------------------------------