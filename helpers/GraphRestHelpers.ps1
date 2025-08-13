# ================================================
# Graph REST Helpers (App-Only per Zertifikat)
# ================================================

# Base64Url helper
function Convert-ToBase64Url([byte[]]$bytes) {
    [Convert]::ToBase64String($bytes).TrimEnd('=').Replace('+','-').Replace('/','_')
}

# Robust das RSA-PrivateKey-Objekt aus einem X509Certificate2 holen
function Get-RsaPrivateKeySafe {
    param([Parameter(Mandatory)][System.Security.Cryptography.X509Certificates.X509Certificate2]$Cert)

    # 1) Moderne Extension-Methode (normaler Weg)
    try {
        $rsa = $Cert.GetRSAPrivateKey()
        if ($rsa) { return $rsa }
    } catch {}

    # 2) Statischer Fallback (wenn Extension nicht gebunden ist)
    try {
        $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Cert)
        if ($rsa) { return $rsa }
    } catch {}

    throw "❌ Im geladenen Zertifikat ist kein RSA-PrivateKey verfügbar (oder nicht zugreifbar)."
}

# Signiert eine Client-Assertion (JWT) mit dem PFX-PrivateKey
function New-ClientAssertionJwt {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][System.Security.Cryptography.X509Certificates.X509Certificate2]$Cert
    )
    $now = [DateTimeOffset]::UtcNow
    $aud = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"

    $header  = @{
        alg = "RS256"
        typ = "JWT"
        # x5t = Base64Url(SHA1-Fingerprint-Bytes)
        x5t = (Convert-ToBase64Url -bytes $Cert.GetCertHash())
    }
    $payload = @{
        iss = $ClientId
        sub = $ClientId
        aud = $aud
        jti = [guid]::NewGuid().ToString()
        nbf = $now.AddSeconds(-30).ToUnixTimeSeconds()
        exp = $now.AddMinutes(9).ToUnixTimeSeconds()
    }

    $encHeader  = Convert-ToBase64Url -bytes ([Text.Encoding]::UTF8.GetBytes(($header  | ConvertTo-Json -Compress)))
    $encPayload = Convert-ToBase64Url -bytes ([Text.Encoding]::UTF8.GetBytes(($payload | ConvertTo-Json -Compress)))
    $toSign     = [Text.Encoding]::UTF8.GetBytes("$encHeader.$encPayload")

    $rsa  = Get-RsaPrivateKeySafe -Cert $Cert
    $sig  = $rsa.SignData($toSign, [Security.Cryptography.HashAlgorithmName]::SHA256, [Security.Cryptography.RSASignaturePadding]::Pkcs1)
    $encSig = Convert-ToBase64Url -bytes $sig

    return "$encHeader.$encPayload.$encSig"
}

# Token-Cache (pro Function-Ausführung)
$script:GraphAccessToken  = $null
$script:GraphTokenExpires = Get-Date 0

# Holt ein App-Only Access Token via Client Assertions (PFX)
function Get-GraphAccessToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$PfxPath,
        [Parameter(Mandatory)][securestring]$PfxPassword
    )

    # Cache nutzen (erneuern ~2 Minuten vor Ablauf)
    if ($script:GraphAccessToken -and (Get-Date) -lt $script:GraphTokenExpires.AddMinutes(-2)) {
        return $script:GraphAccessToken
    }

    # PFX laden (Functions-taugliche Flags)
    if (-not (Test-Path $PfxPath)) {
        throw "❌ PFX-Datei nicht gefunden: $PfxPath"
    }
    $flags = [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::MachineKeySet `
        -bor [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable `
        -bor [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::EphemeralKeySet

    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($PfxPath, $PfxPassword, $flags)

    if (-not $cert.HasPrivateKey) {
        throw "❌ PFX wurde ohne Private Key geladen (HasPrivateKey = False)."
    }

    # Sicherstellen, dass es ein RSA-Key ist (RS256)
    try {
        $null = Get-RsaPrivateKeySafe -Cert $cert
    } catch {
        throw "❌ Zertifikat enthält keinen RSA-PrivateKey für RS256 (evtl. ECDSA?). Bitte ein RSA-PFX verwenden. Details: $($_.Exception.Message)"
    }

    $clientAssertion = New-ClientAssertionJwt -TenantId $TenantId -ClientId $ClientId -Cert $cert

    $tokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $form = @{
        client_id             = $ClientId
        scope                 = "https://graph.microsoft.com/.default"
        grant_type            = "client_credentials"
        client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
        client_assertion      = $clientAssertion
    }

    $resp = Invoke-RestMethod -Method POST -Uri $tokenEndpoint -Body $form -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
    $script:GraphAccessToken  = $resp.access_token
    $script:GraphTokenExpires = (Get-Date).AddSeconds([int]$resp.expires_in)
    return $script:GraphAccessToken
}

# Generischer REST-Caller gegen Graph
function Invoke-Graph {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet('GET','POST','PATCH','PUT','DELETE')]$Method,
        [Parameter(Mandatory)][string]$Uri,
        $Body = $null,
        [int]$TimeoutSec = 120
    )

    # Werte aus ENV (in run.ps1 gesetzt)
    $tenantId = $env:TenantId
    $clientId = $env:ClientId
    $pfxPath  = $env:PfxPath
    $pfxPwd   = (ConvertTo-SecureString $env:PfxPassword -AsPlainText -Force)

    $token = Get-GraphAccessToken -TenantId $tenantId -ClientId $clientId -PfxPath $pfxPath -PfxPassword $pfxPwd

    $headers = @{ Authorization = "Bearer $token" }
    $prms = @{
        Method      = $Method
        Uri         = $Uri
        Headers     = $headers
        TimeoutSec  = $TimeoutSec
        ErrorAction = 'Stop'
    }
    if ($null -ne $Body) {
        $prms.ContentType = 'application/json'
        $prms.Body        = ($Body | ConvertTo-Json -Depth 20 -Compress)
    }

    try {
        return Invoke-RestMethod @prms
    }
    catch {
        # Exception-Message
        if (Get-Command Log -ErrorAction SilentlyContinue) {
            Log ("Graph-Request-Warning: {0} {1} -> {2}" -f $Method, $Uri, $_.Exception.Message)
        }

        # Response-Body auslesen (JSON von Graph)
        try {
            $resp = $_.Exception.Response
            if ($resp -and $resp.GetResponseStream) {
                $stream = $resp.GetResponseStream()
                $reader = New-Object System.IO.StreamReader($stream)
                $reader.BaseStream.Position = 0
                $reader.DiscardBufferedData()
                $respBody = $reader.ReadToEnd()
                if ($respBody) {
                    if (Get-Command Log -ErrorAction SilentlyContinue) { Log "❌ Graph-Response: $respBody" }
                }
            }
        } catch {}

        throw
    }
}

# Kompatibilität: gleicher Name/Signatur wie das Graph-Cmdlet
function Invoke-MgGraphRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet('GET','POST','PATCH','PUT','DELETE')]$Method,
        [Parameter(Mandatory)][string]$Uri,
        $Body
    )
    return Invoke-Graph -Method $Method -Uri $Uri -Body $Body
}

# --------- TEAMS-SPEZIFISCHE REST-HELPER ---------

function Get-CatalogAppId {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ExternalId)
    $uri = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps?`$filter=externalId eq '$ExternalId'&`$select=id,externalId,displayName"
    $res = Invoke-Graph -Method GET -Uri $uri
    if (-not $res.value -or $res.value.Count -eq 0) { throw "App mit ExternalId '$ExternalId' nicht gefunden." }
    return $res.value[0]
}

function Get-AllPages {
    param([Parameter(Mandatory)][string]$FirstUri)

    $items = @()
    $uri   = $FirstUri

    while ($uri) {
        $page = Invoke-Graph -Method GET -Uri $uri
        if ($null -eq $page) { break }

        if ($page.PSObject.Properties.Name -contains 'value') {
            if ($page.value) { $items += $page.value }
        } else {
            break
        }

        if ($page.PSObject.Properties.Name -contains '@odata.nextLink') {
            $uri = $page.'@odata.nextLink'
        } else {
            $uri = $null
        }
    }
    ,$items
}

function Get-TeamsAppInstalled {
    <#
      Sichert App-Installation auf TEAM-Ebene (Graph v1.0).
      - $expand=teamsAppDefinition($select=id,teamsAppId)
      - Paging, idempotenter POST, Polling bis sichtbar
      Rückgabe: installedAppId (string)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ResolvedTeamId,
        [Parameter(Mandatory)][string]$CatalogAppId,
        [int]$TimeoutSeconds = 60,
        [int]$IntervalSeconds = 3
    )

    if (-not $CatalogAppId) { throw "❌ CatalogAppId wurde nicht übergeben!" }
    if (Get-Command Log -ErrorAction SilentlyContinue) { Log "Ensure app $CatalogAppId on team $ResolvedTeamId" }

    # Katalog-Eintrag prüfen (korrekte URL MIT ID im Pfad) – nicht fatal
    try {
        $check = Invoke-Graph -Method GET -Uri "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/$CatalogAppId?`$select=id,displayName,distributionMethod"
        if ($check -and $check.id -ne $CatalogAppId) { throw "App-Katalog-ID unerwartet (erhalten=$($check.id), erwartet=$CatalogAppId)." }
    } catch {
        if (Get-Command Log -ErrorAction SilentlyContinue) { Log "⚠️ Konnte App-Katalogeintrag nicht prüfen: $($_.Exception.Message)" }
    }

    $listUri  = "https://graph.microsoft.com/v1.0/teams/$ResolvedTeamId/installedApps?`$expand=teamsAppDefinition(`$select=id,teamsAppId)"
    $appsAll  = Get-AllPages -FirstUri $listUri
    $existing = $appsAll | Where-Object { $_.teamsAppDefinition.teamsAppId -eq $CatalogAppId }

    if (-not $existing) {
        $installUri = "https://graph.microsoft.com/v1.0/teams/$ResolvedTeamId/installedApps"
        $body = @{ "teamsApp@odata.bind" = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/$CatalogAppId" }

        try {
            $null = Invoke-Graph -Method POST -Uri $installUri -Body $body
            if (Get-Command Log -ErrorAction SilentlyContinue) { Log "📦 App install triggered at team scope." }
        } catch {
            $msg = $_.Exception.Message
            if ($msg -match 'already' -or $msg -match 'exist' -or $msg -match 'conflict' -or $msg -match '"code"\s*:\s*"Conflict"') {
                if (Get-Command Log -ErrorAction SilentlyContinue) { Log "ℹ️ App offenbar bereits installiert (ignoriere Conflict): $msg" }
            } else { throw }
        }
    } else {
        if (Get-Command Log -ErrorAction SilentlyContinue) { Log "ℹ️ App already installed at team scope" }
    }

    $end = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        $appsAll  = Get-AllPages -FirstUri $listUri
        $existing = $appsAll | Where-Object { $_.teamsAppDefinition.teamsAppId -eq $CatalogAppId }
        if ($existing) { return $existing[0].id }
        Start-Sleep -Seconds $IntervalSeconds
    } while ((Get-Date) -lt $end)

    throw "App $CatalogAppId wurde im Team $ResolvedTeamId nicht innerhalb von $TimeoutSeconds s installiert/gefunden."
}

function Get-TeamsAppInstalledForChannel {
    <#
      Sichert App-Installation auf CHANNEL-Ebene (Graph v1.0; private/shared).
      - $expand=teamsAppDefinition($select=id,teamsAppId)
      - Paging, idempotenter POST, Polling bis sichtbar
      Rückgabe: installedAppId (string)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ResolvedTeamId,
        [Parameter(Mandatory)][string]$ChannelId,
        [Parameter(Mandatory)][string]$CatalogAppId,
        [int]$TimeoutSeconds = 60,
        [int]$IntervalSeconds = 3
    )

    if (-not $CatalogAppId) { throw "❌ CatalogAppId wurde nicht übergeben!" }
    if (Get-Command Log -ErrorAction SilentlyContinue) { Log "Ensure app $CatalogAppId on channel $ChannelId (team $ResolvedTeamId)" }

    # Katalog-Eintrag prüfen – nicht fatal
    try {
        $check = Invoke-Graph -Method GET -Uri "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/$CatalogAppId?`$select=id,displayName,distributionMethod"
        if ($check -and $check.id -ne $CatalogAppId) { throw "App-Katalog-ID unerwartet (erhalten=$($check.id), erwartet=$CatalogAppId)." }
    } catch {
        if (Get-Command Log -ErrorAction SilentlyContinue) { Log "⚠️ Konnte App-Katalogeintrag nicht prüfen: $($_.Exception.Message)" }
    }

    $listUri  = "https://graph.microsoft.com/v1.0/teams/$ResolvedTeamId/channels/$ChannelId/installedApps?`$expand=teamsAppDefinition(`$select=id,teamsAppId)"
    $appsAll  = Get-AllPages -FirstUri $listUri
    $existing = $appsAll | Where-Object { $_.teamsAppDefinition.teamsAppId -eq $CatalogAppId }

    if (-not $existing) {
        $installUri = "https://graph.microsoft.com/v1.0/teams/$ResolvedTeamId/channels/$ChannelId/installedApps"
        $body = @{ "teamsApp@odata.bind" = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/$CatalogAppId" }

        try {
            $null = Invoke-Graph -Method POST -Uri $installUri -Body $body
            if (Get-Command Log -ErrorAction SilentlyContinue) { Log "📦 App install triggered at channel scope." }
        } catch {
            $msg = $_.Exception.Message
            if ($msg -match 'already' -or $msg -match 'exist' -or $msg -match 'conflict' -or $msg -match '"code"\s*:\s*"Conflict"') {
                if (Get-Command Log -ErrorAction SilentlyContinue) { Log "ℹ️ App offenbar bereits im Channel installiert (ignoriere Conflict): $msg" }
            } else { throw }
        }
    } else {
        if (Get-Command Log -ErrorAction SilentlyContinue) { Log "ℹ️ App already installed at channel scope (installedAppId=$($existing[0].id))" }
    }

    $end = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        $appsAll  = Get-AllPages -FirstUri $listUri
        $existing = $appsAll | Where-Object { $_.teamsAppDefinition.teamsAppId -eq $CatalogAppId }
        if ($existing) { return $existing[0].id }
        Start-Sleep -Seconds $IntervalSeconds
    } while ((Get-Date) -lt $end)

    throw "App $CatalogAppId wurde im Channel $ChannelId (Team $ResolvedTeamId) nicht innerhalb von $TimeoutSeconds s installiert/gefunden."
}

function Wait-TeamsAppReady {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TeamId,
        [Parameter(Mandatory)][string]$CatalogAppId,
        [int]$TimeoutSeconds = 30,
        [int]$IntervalSeconds = 2
    )

    if (Get-Command Log -ErrorAction SilentlyContinue) { Log "App ready on team: $Team" }

    $listUri = "https://graph.microsoft.com/v1.0/teams/$TeamId/installedApps?`$expand=teamsAppDefinition(`$select=id,teamsAppId)"
    $end = (Get-Date).AddSeconds($TimeoutSeconds)

    do {
        $appsAll  = Get-AllPages -FirstUri $listUri
        $existing = $appsAll | Where-Object { $_.teamsAppDefinition.teamsAppId -eq $CatalogAppId }
        if ($existing) { return $true }
        Start-Sleep -Seconds $IntervalSeconds
    } while ((Get-Date) -lt $end)

    throw "App $CatalogAppId ist im Team $TeamId nach $TimeoutSeconds s nicht sichtbar."
}

function Add-GraphTeamsTab {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TeamId,
        [Parameter(Mandatory)][string]$ChannelId,
        [Parameter(Mandatory)][string]$TabDisplayName,
        [Parameter(Mandatory)][string]$TeamsAppId,   # = Catalog-App-ID (appCatalogs/teamsApps/{id})
        [Parameter(Mandatory)][string]$EntityId,
        [Parameter(Mandatory)][string]$ContentUrl,
        [string]$WebsiteUrl
    )

    if (Get-Command Log -ErrorAction SilentlyContinue) {
        Log "📌 '$TabDisplayName' -> Team '$TeamId'"
    }

    # 0) Idempotenz: existiert ein Tab mit gleichem displayName bereits?
    $tabsUri = "https://graph.microsoft.com/v1.0/teams/$TeamId/channels/$ChannelId/tabs?`$select=id,displayName"
    $tabs = Invoke-Graph -Method GET -Uri $tabsUri
    if ($tabs.value) {
        $existing = $tabs.value | Where-Object { $_.displayName -ieq $TabDisplayName } | Select-Object -First 1
        if ($existing) {
            if (Get-Command Log -ErrorAction SilentlyContinue) { Log "ℹ️ Tab '$TabDisplayName' existiert bereits (id=$($existing.id)) – überspringe Create." }
            return $existing.id
        }
    }

    # 1) Sanity: App muss im Org-Katalog sein (nicht fatal, nur Hinweis)
    try {
        $check = Invoke-Graph -Method GET -Uri "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/$TeamsAppId?`$select=id,displayName,distributionMethod"
        if ($check -and $check.distributionMethod -and $check.distributionMethod -ne 'organization') {
            if (Get-Command Log -ErrorAction SilentlyContinue) {
                Log "⚠️ App '$($check.displayName)' distributionMethod=$($check.distributionMethod) – erwartete 'organization'."
            }
        }
    } catch {
        if (Get-Command Log -ErrorAction SilentlyContinue) {
            Log "⚠️ Konnte App-Katalogeintrag nicht prüfen: $($_.Exception.Message)"
        }
        # kein throw – nicht fatal
    }

    # 2) Request-Body korrekt für v1.0
    #    WICHTIG: teamsApp@odata.bind -> appCatalogs/teamsApps/{id}
    $body = @{
        displayName = $TabDisplayName
        "teamsApp@odata.bind" = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/$TeamsAppId"
        configuration = @{
            entityId   = $EntityId
            contentUrl = $ContentUrl
        }
    }
    if ($WebsiteUrl) { $body.configuration.websiteUrl = $WebsiteUrl }

    $createUri = "https://graph.microsoft.com/v1.0/teams/$TeamId/channels/$ChannelId/tabs"

    try {
        $created = Invoke-Graph -Method POST -Uri $createUri -Body $body
        if ($created -and $created.id) {
            if (Get-Command Log -ErrorAction SilentlyContinue) { Log "✅ Tab '$TabDisplayName' erstellt (id=$($created.id))." }
            return $created.id
        } else {
            if (Get-Command Log -ErrorAction SilentlyContinue) { Log "⚠️ Tab erstellt, aber keine ID im Response – prüfe Sichtbarkeit per Poll." }
        }
    } catch {
        # Hier kommt dein 400 – gib den Grund aus dem Response-Body aus (Invoke-Graph loggt bereits)
        $msg = $_.Exception.Message
        # Häufige Ursachen erkennen und freundlich erklären:
        if ($msg -match 'contentUrl' -or $msg -match 'domain' -or $msg -match 'validDomains') {
            throw "❌ Tab-Body/Domain-Problem: Ist die Domain in der App-Manifest 'validDomains' enthalten? contentUrl=$ContentUrl. Original: $msg"
        }
        if ($msg -match 'not allowed' -or $msg -match 'scope') {
            throw "❌ App-Manifest-Scope: Unterstützt die App 'configurableTabs' mit scope 'team'? Original: $msg"
        }
        throw
    }

    # 3) Fallback: Warte kurz bis sichtbar
    try {
        Wait-TeamsTabVisible -TeamId $TeamId -ChannelId $ChannelId -TabDisplayName $TabDisplayName -TimeoutSeconds 45 -IntervalSeconds 3
        return $true
    } catch {
        throw "❌ Tab '$TabDisplayName' wurde nach Create nicht sichtbar: $($_.Exception.Message)"
    }
}
