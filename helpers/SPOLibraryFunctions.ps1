# --------------------------------------------------------------------
# Include Helper Functions
# --------------------------------------------------------------------
. (Join-Path $PSScriptRoot 'LoggingFunctions.ps1')
. (Join-Path $PSScriptRoot 'AdminFunctions.ps1')

# ------------------------------------------------------------------------------------
# Bibliothek / Drive einer Sharepoint Site Collection triggern und Erstellung abwarten
# ------------------------------------------------------------------------------------
function Wait-ForGroupDrive {
    param(
        [Parameter(Mandatory)][guid]$groupId,
        [string]$DriveName,
        [int]$maxTries = 30,
        [int]$delaySeconds = 10
    )

    for ($i = 1; $i -le $maxTries; $i++) {
        try {
            $driveResp = Invoke-PnPGraphMethod -Method GET -Url "https://graph.microsoft.com/v1.0/groups/$groupId/drives"
            foreach ($d in $driveResp.value) {
                if ($d.driveType -eq "documentLibrary") {
                    log "DMS Drive: '$DriveName' / actual Drive: '$($d.name)'"
                    if (($DriveName -ne "") -and ($DriveName -ne $d.name)) { continue }
                    try {
                        $rootItem = Invoke-PnPGraphMethod -Method GET `
                            -Url "https://graph.microsoft.com/v1.0/drives/$($d.id)/root"
                        if ($null -eq $rootItem.id) { continue }

                        # Test: Schreibe einen Dummy-Ordner ins Root!
                        $testFolderName = "___provisioning_probe_" + (Get-Random)
                        $testBodyJson = @"
{
    "name": "$testFolderName",
    "folder": {},
    "@microsoft.graph.conflictBehavior": "fail"
}
"@

                        try {
                            $testResp = Invoke-PnPGraphMethod -Method POST `
                                -Url "https://graph.microsoft.com/v1.0/drives/$($d.id)/root/children" `
                                -Content $testBodyJson -ContentType "application/json"

                            # Wenn wir hier landen: Drive ist SCHREIBBAR!
                            # Dummy-Ordner wieder löschen:
                            $dummyId = $testResp.id
                            if ($dummyId) {
                                Invoke-PnPGraphMethod -Method DELETE `
                                    -Url "https://graph.microsoft.com/v1.0/drives/$($d.id)/items/$dummyId"
                            }

                            Write-Information "✅ DRIVE REALLY READY (Try $i): Schreibtest erfolgreich!"
                            return @{ drive = $d; rootItem = $rootItem }
                        }
                        catch {
                            Write-Information "🔄 Schreibtest noch nicht möglich (Try $i)..."
                            # Noch nicht bereit – weiter warten!
                        }
                    }
                    catch {
                        Write-Information "🔄 Root folder nicht gefunden (Try $i)"
                    }
                }
            }
        }
        catch {
            Write-Warning "⚠️ Drive lookup failed: $($_.Exception.Message) (Try $i)"
        }
        Start-Sleep -Seconds $delaySeconds
    }
    throw "❌ Timeout: Drive konnte nicht schreibbar provisioniert werden nach $maxTries Versuchen"
}

# ------------------------------------------------------------------------------------
# Bibliothek / Drive einer Sharepoint Site Collection triggern und Erstellung abwarten
# ------------------------------------------------------------------------------------
function Add-Folders {
    param (
        [array]$items,
        [string]$parentId,
        [string]$driveId
    )

    foreach ($item in $items) {
        $folderName = $item.name

        # Hole ALLE Kinder und prüfe dann lokal, ob der Ordner bereits existiert
        $childrenUrl = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$parentId/children"
        try {
            $existingChildren = Invoke-PnPGraphMethod -Method GET -Url $childrenUrl

            $existing = $existingChildren.value | Where-Object { 
                $_.name -eq $folderName -and $null -ne $_.folder
            }

            if ($existing) {
                Write-Information "ℹ️ Ordner bereits vorhanden: $folderName"
                $newFolder = $existing
            } else {
                $bodyJson = @"
{
    "name": "$folderName",
    "folder": {},
    "@microsoft.graph.conflictBehavior": "rename"
}
"@

                $newFolder = Invoke-PnPGraphMethod -Method POST `
                    -Url "https://graph.microsoft.com/v1.0/drives/$driveId/items/$parentId/children" `
                    -Content $bodyJson -ContentType "application/json"

                Write-Information "📁 Created: $($newFolder.name)"
            }

            # Rekursiv Kinder anlegen
            if ($item.children) {
                Add-Folders -items $item.children -parentId $newFolder.id -driveId $driveId
            }
        }
        catch {
            Write-Warning "⚠️ Failed to handle folder '$folderName': $($_.Exception.Message)"
        }
    }
}

# ------------------------------------------------------------------------------------

# Idempotent helper: ensure these users are owners of the M365 group
function Set-M365GroupOwners {
    param(
        [Parameter(Mandatory)][string]$GroupId,
        [Parameter(Mandatory)][string[]]$Users
    )

    $maxRetries = 10; $waitSec = 3; $ownersUpn = @()
    for ($i=0; $i -lt $maxRetries; $i++) {
        try {
            $owners = Get-PnPMicrosoft365GroupOwners -Identity $GroupId -ErrorAction Stop
            $ownersUpn = @(
                foreach ($o in $owners) {
                    if ($o.UserPrincipalName) { $o.UserPrincipalName.ToLower() }
                    elseif ($o.Mail) { $o.Mail.ToLower() }
                }
            )
            break
        } catch {
            if ($i -lt ($maxRetries-1)) {
                Log "⏳ Owners noch nicht lesbar – retry $($i+1)/$maxRetries"
                Start-Sleep -Seconds $waitSec
            } else { throw }
        }
    }

    $usersNorm = @($Users | Where-Object { $_ } | ForEach-Object { $_.ToLower() } | Select-Object -Unique)
    $toAdd = @($usersNorm | Where-Object { $ownersUpn -notcontains $_ })

    if ($toAdd.Count -gt 0) {
        Log "👑 Füge fehlende Owner hinzu: $($toAdd -join ', ')"
        try {
            Add-PnPMicrosoft365GroupOwner -Identity $GroupId -Users $toAdd -ErrorAction Stop
        } catch {
            if ($_.Exception.Message -match "already exist.*owners") {
                Log "ℹ️ Owner bereits vorhanden (Race/Case) – weiter."
            } else { throw }
        }
    } else {
        Log "ℹ️ Alle gewünschten Owner sind bereits gesetzt."
    }

    # Endstand Owners loggen
    $finalOwners = Get-PnPMicrosoft365GroupOwners -Identity $GroupId -ErrorAction SilentlyContinue
    $finalList = $finalOwners | ForEach-Object { $_.UserPrincipalName ?? $_.Mail }
    Log "📋 Aktuelle Owners: $($finalList -join ', ')"
}

function Set-M365GroupMembers {
    param(
        [Parameter(Mandatory)][string]$GroupId,
        [Parameter(Mandatory)][string[]]$Users
    )

    $maxRetries = 10; $waitSec = 3; $membersUpn = @()
    for ($i=0; $i -lt $maxRetries; $i++) {
        try {
            $members = Get-PnPMicrosoft365GroupMembers -Identity $GroupId -ErrorAction Stop
            $membersUpn = @(
                foreach ($m in $members) {
                    if ($m.UserPrincipalName) { $m.UserPrincipalName.ToLower() }
                    elseif ($m.Mail) { $m.Mail.ToLower() }
                }
            )
            break
        } catch {
            if ($i -lt ($maxRetries-1)) {
                Log "⏳ Members noch nicht lesbar – retry $($i+1)/$maxRetries"
                Start-Sleep -Seconds $waitSec
            } else { throw }
        }
    }

    $usersNorm = @($Users | Where-Object { $_ } | ForEach-Object { $_.ToLower() } | Select-Object -Unique)
    $toAdd = @($usersNorm | Where-Object { $membersUpn -notcontains $_ })

    if ($toAdd.Count -gt 0) {
        Log "👥 Füge fehlende Members hinzu: $($toAdd -join ', ')"
        try {
            Add-PnPMicrosoft365GroupMember -Identity $GroupId -Users $toAdd -ErrorAction Stop
        } catch {
            if ($_.Exception.Message -match "already exist.*members") {
                Log "ℹ️ Members bereits vorhanden (Race/Case) – weiter."
            } else { throw }
        }
    } else {
        Log "ℹ️ Alle gewünschten Members sind bereits gesetzt."
    }

    # Endstand Members loggen
    $finalMembers = Get-PnPMicrosoft365GroupMembers -Identity $GroupId -ErrorAction SilentlyContinue
    $finalList = $finalMembers | ForEach-Object { $_.UserPrincipalName ?? $_.Mail }
    Log "📋 Aktuelle Members: $($finalList -join ', ')"
}

# ------------------------------------------------------------------------------------
# Document Set (and Content Type) auf der Site Collection aktivieren
# ------------------------------------------------------------------------------------
function Test-PnPFeatureEnabled {
    param(
        [Parameter(Mandatory)]
        [string]$FeatureId,
        [ValidateSet("Site","Web")]
        [string]$Scope = "Site"
    )
    # Get-PnPFeature returns enabled features at the given scope
    # Compare on the string form to avoid Guid type quirks
    $enabled = Get-PnPFeature -Scope $Scope -ErrorAction SilentlyContinue |
        Where-Object { "$($_.DefinitionId)" -ieq $FeatureId }
    return [bool]$enabled
}

function Find-PnPFeature {
    param(
        [Parameter(Mandatory)]
        [string]$FeatureId,
        [ValidateSet("Site","Web")]
        [string]$Scope = "Site",
        [switch]$Force
    )
    if (Test-PnPFeatureEnabled -FeatureId $FeatureId -Scope $Scope) {
        Log "ℹ️ Feature $FeatureId already enabled at scope $Scope, skipping"
        return
    }
    Log "Enable-PnPFeature -Identity $FeatureId -Scope $Scope"
    try {
        Enable-PnPFeature -Identity $FeatureId -Scope $Scope -Force:$Force -ErrorAction Stop
        Log "✅ Feature $FeatureId enabled at scope $Scope"
    }
    catch {
        # If another concurrent thread just enabled it, treat as success
        if (Test-PnPFeatureEnabled -FeatureId $FeatureId -Scope $Scope) {
            Log "ℹ️ Feature $FeatureId became enabled during retry window; continuing"
        } else {
            throw
        }
    }
}

function Enable-DocumentSets {
    # IDs
    $docIdFeature = "b50e3104-6812-424f-a011-cc90e6327318" # Document ID Service (Site Collection)
    $docSetFeature = "3bae86a2-776d-499d-9db8-fa4cdc7884f8" # Document Set (Site Collection)
    $docSetCtPrefix = "0x0120D520"

    # 1) Ensure features
    Find-PnPFeature -FeatureId $docIdFeature -Scope Site -Force
    Find-PnPFeature -FeatureId $docSetFeature -Scope Site -Force

    # 2) Wait until the Document Set content type appears (propagation can lag)
    $maxRetries = 30
    $secWait    = 5
    $retry      = 0
    $docSetCT   = $null

    Log "Prüfe auf Content Type: 'Document Set' (ID beginnt mit $docSetCtPrefix)..."
    do {
        $docSetCT = Get-PnPContentType -ErrorAction SilentlyContinue |
                    Where-Object { $_.Id -like "$docSetCtPrefix*" }
        if ($docSetCT) { break }
        $retry++
        Log "⏳ Warte auf Content Type: 'Document Set'... ($retry/$maxRetries)"
        Start-Sleep -Seconds $secWait
    } while ($retry -lt $maxRetries)

    if (-not $docSetCT) {
        throw "❌ Content Type: 'Document Set' nicht gefunden – bitte prüfen, ob die Features auf der Website­sammlung aktiv sind."
    } else {
        Log "✅ Content Type: 'Document Set' aktiv: $($docSetCT.Name) / $($docSetCT.Id)"
    }
}

# ------------------------------------------------------------------------------------
# Content Types zur Library hinzufügen
# ------------------------------------------------------------------------------------
function Add-ContentType {
    param(
        [string]$libName,       # Name der Bibliothek, z.B. "Documents"
        [string]$docSetName,    # Name des Content Types, z.B. "Email"
        [string]$docSetId       # ID des Content Types, z.B. "0x0120D520"
    )

    Log "Prüfe auf Content Type: '$docSetName' ..."
    $docSetCT = Get-PnPContentType | Where-Object { $_.Name -like "$docSetName*" }
    if ($docSetCT) {
            $docSetName = $docSetCT.Name
            $docSetID = $docSetCT.Id
            Log "Content Type: '$docSetName' der Library '$libName' zuweisen"
            Add-PnPContentTypeToList -List $libName -ContentType $docSetCT.Id
        } else {
            Log "Content Type: '$docSetName' für Library '$libName' nicht zugewiesen."
        }
}

# ------------------------------------------------------------------------------------
function Test-SpoSiteOrAliasExists {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$TenantName,   # z.B. "contoso"
        [Parameter(Mandatory=$true)][string]$Alias,
        [switch]$ReturnFirstUrlOnly
    )

    # Alias normalisieren wie bei New-PnPSite
    $norm = ($Alias.ToLowerInvariant() -replace '[^a-z0-9._-]', '')

    # 1) Direktprüfung: existiert Site bereits unter /sites|/teams ?
    $candidates = @(
        "https://$TenantName.sharepoint.com/sites/$norm",
        "https://$TenantName.sharepoint.com/teams/$norm"
    )

    Log "Prüfe auf existierende Site:"
    Log "$($candidates -join ', ')"

    $found = @()
    foreach ($u in $candidates) {
        try {
            $ts = Get-PnPTenantSite -Url $u -ErrorAction SilentlyContinue
            if ($ts) {
                $found += [pscustomobject]@{
                    Url     = $ts.Url
                    Template= $ts.Template
                    # Tipp: GroupId bekommst du über Get-PnPTenantSite (kein -Includes an Get-PnPWeb nötig)
                    GroupId = $ts.GroupId
                }
                if ($ReturnFirstUrlOnly) { break }
            }
        } catch { }
    }
    if ($found.Count -gt 0) {
        return [pscustomobject]@{
            Exists  = $true
            Reason  = 'SiteExists'
            SiteUrl = $found[0].Url
            GroupId = $found[0].GroupId
            Source  = 'Get-PnPTenantSite'
        }
    }

    # 2) Preflight: ist der Alias (SharePoint-seitig) überhaupt frei?
    try {
        $isFree = Get-PnPIsSiteAliasAvailable -Alias $norm
        if (-not $isFree) {
            # Alias ist belegt -> versuche Vorschlags-URL zu bekommen (REST, kein Graph)
            $root = "https://$TenantName.sharepoint.com"
            $suggest = Invoke-PnPSPRestMethod -Method Get -Url "$root/_api/GroupSiteManager/GetValidSiteUrlFromAlias?alias='$norm'&isTeamSite=true"
            Log "$suggest"
            # Hinweis: Der Endpoint liefert eine "gültige" URL zurück; wenn != gewünschtem Pfad, ist der Alias belegt. :contentReference[oaicite:3]{index=3}
            $suggestedUrl = ($suggest?.GetValidSiteUrlFromAlias) ?? $null

            return [pscustomobject]@{
                Exists  = $true
                Reason  = 'AliasInUse'
                SiteUrl = $suggestedUrl
                GroupId = $null
                Source  = 'Get-PnPIsSiteAliasAvailable/GroupSiteManager'
            }
        }
    } catch {
        Write-Warning "Alias-Preflight nicht möglich: $($_.Exception.Message)"
    }

    # 3) Optionaler Fallback: REST-SiteStatus prüfen (0/1/2/3/4) 4 = URL belegt :contentReference[oaicite:4]{index=4}
    try {
        $root = "https://$TenantName.sharepoint.com"
        $status = Invoke-PnPSPRestMethod -Method Post -Url "$root/_api/SPSiteManager/status" `
                  -ContentType "application/json;odata=nometadata" `
                  -Body (@{ url = "https://$TenantName.sharepoint.com/sites/$norm" } | ConvertTo-Json)
        if ($status.SiteStatus -eq 4) {
            return [pscustomobject]@{
                Exists  = $true
                Reason  = 'SiteExists'
                SiteUrl = $status.SiteUrl
                GroupId = $null
                Source  = 'SPSiteManager/status'
            }
        }
    } catch { }

    # Nichts gefunden -> frei
    return [pscustomobject]@{ Exists = $false; Reason = $null; SiteUrl = $null; GroupId = $null; Source = 'None' }
}

# ------------------------------------------------------------------------------------
function Test-AliasOrSiteExistsGraph {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$Alias,
        [Parameter(Mandatory=$true)][string]$TenantName,   # z.B. "contoso"
        [switch]$ReturnFirstUrlOnly
    )

    # ---- Alias normalisieren wie mailNickname (grob) ----
    $normAlias = $Alias.ToLowerInvariant() -replace '[^a-z0-9._-]', ''
    if ([string]::IsNullOrWhiteSpace($normAlias)) {
        return [pscustomobject]@{
            Exists     = $true
            Reason     = 'InvalidAlias'
            Details    = 'Alias enthält keine gültigen Zeichen'
            SiteUrl    = $null
            ObjectType = $null
            ObjectId   = $null
        }
    }

    # ---- 1) Graph: Alias-Kollision prüfen (Groups/Users/Contacts) ----
    # Voraussetzungen: App-Only oder Delegated mit ausreichenden Rechten.
    # Tipp: -ConsistencyLevel eventual erhöht Filter-Zuverlässigkeit.
    $collision = $null
    try {
        # M365-Gruppen (inkl. DL / mailaktivierte Gruppen)
        $g = Get-MgGroup -Filter "mailNickname eq '$normAlias'" -All -ConsistencyLevel eventual -ErrorAction Stop
        if ($g) { $collision = @{ Type='Group'; Object=$g[0] } }

        if (-not $collision) {
            $u = Get-MgUser -Filter "mailNickname eq '$normAlias'" -All -ConsistencyLevel eventual -ErrorAction Stop
            if ($u) { $collision = @{ Type='User'; Object=$u[0] } }
        }
        if (-not $collision) {
            $c = Get-MgContact -Filter "mailNickname eq '$normAlias'" -All -ErrorAction Stop
            if ($c) { $collision = @{ Type='Contact'; Object=$c[0] } }
        }
    } catch {
        Write-Warning "Graph-Abfrage für Alias-Kollision: $($_.Exception.Message)"
    }

    # Falls eine Gruppe existiert, versuche Site-URL herzuleiten (optional)
    $groupSiteUrl = $null
    if ($collision -and $collision.Type -eq 'Group') {
        try {
            # Der Drive.WebUrl zeigt auf "…/Shared Documents". Basis-Site-URL ableiten:
            $drv = Get-MgGroupDrive -GroupId $collision.Object.Id -ErrorAction SilentlyContinue
            if ($drv.WebUrl) {
                $groupSiteUrl = ($drv.WebUrl -replace '/Shared Documents.*$', '')
            }
        } catch { }
    }

    if ($collision) {
        return [pscustomobject]@{
            Exists     = $true
            Reason     = 'AliasCollision'
            Details    = "mailNickname belegt durch $($collision.Type)"
            SiteUrl    = $groupSiteUrl
            ObjectType = $collision.Type
            ObjectId   = $collision.Object.Id
        }
    }

    # ---- 2) SharePoint: Site-Existenz prüfen ----
    $candidates = @(
        "https://$TenantName.sharepoint.com/sites/$normAlias",
        "https://$TenantName.sharepoint.com/teams/$normAlias"
    )

    $foundSites = @()
    foreach ($url in $candidates) {
        try {
            $s = Get-PnPTenantSite -Url $url -ErrorAction SilentlyContinue
            if ($s) { $foundSites += $url; if ($ReturnFirstUrlOnly) { break } }
        } catch {
            # Ignorieren – nicht existent oder kein Zugriff
        }
    }

    if ($foundSites.Count -gt 0) {
        return [pscustomobject]@{
            Exists     = $true
            Reason     = 'SiteExists'
            Details    = 'Es existiert bereits eine Site unter diesem Pfad'
            SiteUrl    = $foundSites[0]
            ObjectType = 'SharePointSite'
            ObjectId   = $null
        }
    }

    # ---- Kein Konflikt gefunden ----
    return [pscustomobject]@{
        Exists     = $false
        Reason     = $null
        Details    = $null
        SiteUrl    = $null
        ObjectType = $null
        ObjectId   = $null
    }
}
# ------------------------------------------------------------------------------------
# --- Helper: robustes Mapping von SP-TimeZone-Description -> PnP-Enum (inkl. MEZ=4)
function Resolve-PnPTimeZone {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$InputText)

    function Normalize([string]$s){
        $s = $s.Trim().ToUpperInvariant()
        $s = $s -replace 'Ä','AE' -replace 'Ö','OE' -replace 'Ü','UE' -replace 'ß','SS'
        $s = ($s -replace '\s+',' ').Trim()
        return $s
    }

    $norm = Normalize $InputText
    # Schnelle CET/MEZ-Abkürzungen
    if ($norm -match '(^|[^A-Z])(MEZ|CET|CENTRAL EUROPEAN)($|[^A-Z])') {
        return [PnP.Framework.Enums.TimeZone]::UTCPLUS0100_AMSTERDAM_BERLIN_BERN_ROME_STOCKHOLM_VIENNA  # = 4
    }

    # Offset aus "UTC±HH:MM"
    $offsetKey = $null
    if ($norm -match 'UTC\s*([+\-])\s*(\d{1,2})(?::?(\d{2}))?') {
        $sign = ($matches[1] -eq '+') ? 'PLUS' : 'MINUS'
        $hh = '{0:D2}' -f [int]$matches[2]; $mm = '{0:D2}' -f [int]($matches[3] ?? 0)
        $offsetKey = "UTC${sign}${hh}${mm}"
    }

    $names = [Enum]::GetNames([PnP.Framework.Enums.TimeZone])
    $candidates = $offsetKey ? ($names | Where-Object { $_ -like "$offsetKey*" }) : $names

    # Stadt-Hints helfen beim +01:00-Match
    $hints = @('AMSTERDAM','BERLIN','VIENNA','ROME','STOCKHOLM','BERN','BRUSSELS','PARIS','MADRID','PRAGUE','BUDAPEST','WARSAW')
    $best = $null; $bestScore = -1
    foreach ($n in $candidates) {
        $score = 0
        foreach ($h in $hints) { if ($norm -like "*$h*") { $score++ } }
        if ($score -gt $bestScore) { $best = $n; $bestScore = $score }
    }

    if (-not $best -and $offsetKey -eq 'UTCPLUS0100') {
        $best = 'UTCPLUS0100_AMSTERDAM_BERLIN_BERN_ROME_STOCKHOLM_VIENNA'
    }
    if (-not $best) { $best = 'UTC_COORDINATED_UNIVERSAL_TIME' }
    return [PnP.Framework.Enums.TimeZone]::$best
}

function Get-SPOGroupInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SiteUrl
    )

    # --- Web laden (ohne ungültige Includes)
    $web = Get-PnPWeb -Includes Title,Url,WebTemplate,RegionalSettings,RegionalSettings.TimeZone
    $rs  = $web.RegionalSettings

    # --- GroupId sicher holen (CSOM -> REST Fallback)
    $groupId = [guid]::Empty
    try {
        $site = Get-PnPSite
        Get-PnPProperty -ClientObject $site -Property GroupId
        $groupId = $site.GroupId
    } catch {
        try {
            $resp = Invoke-PnPSPRestMethod -Method Get -Url "$($web.Url)/_api/site?`$select=GroupId"
            if ($resp.GroupId) { $groupId = [guid]$resp.GroupId }
        } catch { }
    }

    $Privacy = ""
    # --- Type herleiten
    if ($web.WebTemplate -eq 'GROUP' -and $groupId -ne [guid]::Empty) {
        $Type = 'TeamSite'
        $grp = Get-PnPMicrosoft365Group -Identity $site.GroupId
        $Privacy = $grp.Visibility   # -> "Public" oder "Private"
        Log "Group Visibility (Private/Public): $Privacy"
    } elseif ($web.WebTemplate -eq 'SITEPAGEPUBLISHING') {
        $Type = 'CommunicationSite'
    } else {
        $Type = 'TeamSiteWithoutMicrosoft365Group'
    }

    # --- Alias ohne Graph: aus der Site-URL (Segment nach /sites/ oder /teams/)
    $Alias = $null
    try {
        if ($groupId -ne [guid]::Empty) {
            $u = [Uri]$web.Url
            $lastSeg = $u.Segments[-1].TrimEnd('/')
            # Falls /sites/<alias>/  -> letztes Segment = alias
            # Falls URL auf Root endet (selten), fallback:
            if (-not $lastSeg -or $lastSeg -eq 'sites/' -or $lastSeg -eq 'teams/') {
                $lastSeg = $u.Segments[-2].TrimEnd('/')
            }
            $Alias = [Uri]::UnescapeDataString($lastSeg)
        }
    } catch { $Alias = $null }

    # --- TimeZone-Enum (wie New-PnPSite)
    $tzEnum = Resolve-PnPTimeZone -InputText $rs.TimeZone.Description

    # --- SortOrder/CalendarType/Lcid
    $Lcid      = $rs.LocaleId
    $SortOrder = $rs.Collation
    $CalType   = $rs.CalendarType

    [pscustomobject]@{
        Lcid         = $Lcid
        Title        = $web.Title
        Type         = $Type
        Alias        = $Alias
        TimeZone     = $tzEnum.ToString()   # New-PnPSite erwartet Enum-Name
        SortOrder    = $SortOrder
        CalendarType = $CalType
        GroupId      = $groupId
        SiteUrl      = $web.Url
        Privacy      = $Privacy
    }
}
# ------------------------------------------------------------------------------------
function Find-ListsUsingField {
    param([Parameter(Mandatory)][string]$FieldInternalName)
    $hits = @()
    foreach ($l in Get-PnPList) {
        # Feld direkt an der Liste?
        $listField = Get-PnPField -List $l -Identity $FieldInternalName -ErrorAction SilentlyContinue
        if ($listField) { $hits += $l; continue }

        # Oder via Content Type in der Liste?
        $cts = Get-PnPContentType -List $l -ErrorAction SilentlyContinue
        foreach ($ct in $cts) {
            try {
                $fl = Get-PnPField -List $l -Identity $FieldInternalName -ErrorAction SilentlyContinue
                if ($fl) { $hits += $l; break }
            } catch {}
        }
    }
    $hits | Select-Object -Unique
}

# ------------------------------------------------------------------------------------
function DetachFieldFromList {
    param([Parameter(Mandatory)]$List,[Parameter(Mandatory)][string]$FieldInternalName)

    # Content Types aktivieren, sonst können wir nicht sauber über CTs arbeiten
    if (-not $List.ContentTypesEnabled) {
        Set-PnPList -Identity $List -ContentTypesEnabled $true | Out-Null
    }

    # 1) Feld-Links aus allen CTs der Liste entfernen (falls vorhanden)
    $cts = Get-PnPContentType -List $List -ErrorAction SilentlyContinue
    foreach ($ct in $cts) {
        try {
            Remove-PnPFieldFromContentType -Field $FieldInternalName -ContentType $ct -ErrorAction SilentlyContinue
        } catch {}
    }

    # 2) Falls das Feld zusätzlich direkt an der Liste hängt: entfernen
    try {
        $lf = Get-PnPField -List $List -Identity $FieldInternalName -ErrorAction SilentlyContinue
        if ($lf) {
            Remove-PnPField -List $List -Identity $FieldInternalName -Force -ErrorAction SilentlyContinue
        }
    } catch {}
}

# ------------------------------------------------------------------------------------
# Holt (oder bindet) den Dokument-CT (0x0101*) an die angegebene Liste, unabhängig von Sprache/Name
function Get-OrAttach-DocumentCT {
    param([Parameter(Mandatory)]$List)

    # Content Types an der Liste aktivieren, falls nötig
    if (-not $List.ContentTypesEnabled) {
        Set-PnPList -Identity $List -ContentTypesEnabled $true | Out-Null
    }

    # 1) Zuerst nach einem bereits LIST-gebundenen Dokument-CT suchen (ID-basiert!)
    $ct = Get-PnPContentType -List $List -ErrorAction SilentlyContinue |
          Where-Object { $_.StringId -like "0x0101*" } |
          Select-Object -First 1
    if ($ct) { return $ct }

    # 2) Sonst Site-CT mit ID 0x0101 holen (Name ist egal/übersetzt)
    $siteDocCt = Get-PnPContentType -Identity "0x0101" -ErrorAction SilentlyContinue
    if ($siteDocCt) {
        Add-PnPContentTypeToList -List $List -ContentType $siteDocCt -ErrorAction SilentlyContinue
        # Danach erneut aus der Liste auflösen (damit wir die LIST-Instanz bekommen)
        $ct = Get-PnPContentType -List $List -ErrorAction SilentlyContinue |
              Where-Object { $_.StringId -like "0x0101*" } |
              Select-Object -First 1
        if ($ct) { return $ct }
    }

    throw "Document-Content-Type (0x0101*) wurde im Site/Web nicht gefunden und konnte der Liste nicht hinzugefügt werden."
}

# ------------------------------------------------------------------------------------
<#
function ReattachFieldToListViaDocumentCT {
    param([Parameter(Mandatory)]$List,[Parameter(Mandatory)][string]$FieldInternalName)

    $docCt = Get-OrAttach-DocumentCT -List $List   # <-- statt Get-PnPContentType -Identity "Document"
    Add-PnPFieldToContentType -ContentType $docCt -Field $FieldInternalName -ErrorAction SilentlyContinue
}
#>

function ReattachFieldToListViaDocumentCT {
    param([Parameter(Mandatory)]$List,[Parameter(Mandatory)][string]$FieldInternalName)

    if (-not $List.ContentTypesEnabled) {
        Set-PnPList -Identity $List -ContentTypesEnabled $true | Out-Null
    }

    # Sicherstellen, dass der Document-CT an der Liste hängt
    #$docCt = Get-PnPContentType -Identity "Document"
    $docCt = Get-OrAttach-DocumentCT -List $List
    Add-PnPContentTypeToList -List $List -ContentType $docCt -ErrorAction SilentlyContinue

    # Feld als FieldRef am Document-CT hinzufügen
    Add-PnPFieldToContentType -ContentType $docCt -Field $FieldInternalName -ErrorAction SilentlyContinue
}

# ------------------------------------------------------------------------------------