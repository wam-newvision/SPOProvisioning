param($Request, $TriggerMetadata)

# -------- Helpers & Core (gemeinsam aus wwwroot\Helpers) --------
$functionRoot = Split-Path -Parent $PSScriptRoot       # …\wwwroot
$helpersDir   = Join-Path $functionRoot 'Helpers'

. (Join-Path $helpersDir 'LoggingFunctions.ps1')
. (Join-Path $helpersDir 'AdminFunctions.ps1')
. (Join-Path $helpersDir 'SPOLibraryFunctions.ps1')
. (Join-Path $helpersDir 'ProvisionPnP.ps1')
. (Join-Path $helpersDir 'PSHelpers.ps1')

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
    # --------------- Eingaben prüfen ------------------------------
Log "Eingaben prüfen und Variablen initialisieren..."
    # Pflichtfelder
    $requiredParams = @(
        "tenantId", # e.g. "contoso.onmicrosoft.com"
        "siteTitle", # e.g. "Contoso Project Site"
        "hubName", # e.g. "Contoso Hub"
        "creators",  # Array of initial creator(s)
        "owners",  # Array of owners
        "members"  # Array of members
    )

    # Boolean Felder (true/false), wenn nicht angegeben, dann true
    $booleanParams = @(
        "TeamsTabAfterProvisioning", # false = NUR Teams Tab anlegen (Provisioning überspringen)
        "enableTabCreation", # Standard: true = Teams Tab anlegen
        "structureInDefaultChannelFolder",
        "enableProvisioning",
        "enableDocumentSets",
        "enableCleanUp"  # Set to true to delete unnecessary library elements
    )

    # Optionale Felder (nur falls nötig)
    $optionalParams = @(
        @{ Name = "structure"; Default = "" },     # Array of folder structure
        @{ Name = "Privacy"; Default = "Private" },     # Privacy setting for the site (Private/Public)
        @{ Name = "DMSdrive"; Default = $null },
        @{ Name = "SetRegion"; Default = "1031" },    # 1031 = Deutsch (Deutschland)
        @{ Name = "SetTimezone"; Default = "" },      # 4 = Mitteleuropäische Zeit (Berlin, Wien, Zürich), nur setzen, wenn Region != 1031
        @{ Name = "TabDisplayName"; Default = "AI Agent" },   # Name des Tabs
        @{ Name = "TeamsTabURL"; Default = "https://mwpnewvision.sharepoint.com/sites/projekte" },      # URL for Teams Tab (e.g. "https://mwpnewvision.sharepoint.com/sites/contoso/SitePages/Forms/ByAuthor.aspx")
        @{ Name = "SetSortOrder"; Default = "" }      # 25 = Deutsch, nur setzen, wenn Region != 1031
    )

    # Definierte xxx...Params Felder auf Vorhandensein prüfen und auswerten
    $params = EvaluateRequestParameters -Request $Request -RequiredParams $requiredParams -BooleanParams $booleanParams -OptionalParams $optionalParams

    # Setze alle Variablen für die Pflicht-Parameter
    foreach ($param in $requiredParams) {
        Set-Variable -Name $param -Value $params[$param]
    }

    # Setze alle Variablen für die booleschen Parameter
    foreach ($param in $booleanParams) {
        Set-Variable -Name $param -Value $params[$param]
    }

    # Setze alle Variablen für die optionalen Parameter
    foreach ($param in $optionalParams) {
        $paramName = $param.Name
        Set-Variable -Name $paramName -Value $params[$paramName]
    }

    # --------------------------------------------------------------------
    # Ordnerstruktur-Schema prüfen
    if ($structure) {
        Log "Check Folder structure schema for Site '$siteTitle'"
        if(Test-Schema -structure $structure) {
            Log "✅ Folder structure schema is valid."
        } else {
            Log "❌ Invalid folder structure schema for Site '$siteTitle'."
            $structure = ""
        }
    } else {
        Log "ℹ️ No folder structure provided for Site '$siteTitle'."
    }

    # --------------------------------------------------------------------
    # Alias / URLs
    # --------------------------------------------------------------------
    $TenantName = $tenantId.Split('.')[0]
    $alias      = ($siteTitle -replace '\s+', '')
    $teamId     = $alias  # Standardmäßig Alias als TeamId verwenden
    $siteUrl    = "https://${TenantName}.sharepoint.com/sites/$alias"
    $adminUrl   = "https://${TenantName}-admin.sharepoint.com"
    
    Log "🔗 SiteUrl = $siteUrl"

    # --------------------------------------------------------------------
    # NUR Teams Tab anlegen (sonst nichts tun)
    # --------------------------------------------------------------------
    # Erlaube Teams Tab Creation nur, wenn Umgebungsvariable gesetzt ist
    if ($env:ALLOW_TEAMS_TAB_CREATION -eq 'true') {
        Log "ℹ️ ALLOW_TEAMS_TAB_CREATION is set, User defined Teams Tab creation = $enableTabCreation"            
    } else {
        $enableTabCreation = $false
        Log "ℹ️ ALLOW_TEAMS_TAB_CREATION is not set, Teams Tab creation = $enableTabCreation"
    }
    
    if (-not $TeamsTabAfterProvisioning -and $enableTabCreation) { 
        Log "---- Call TeamsTab-Function (HTTP) ----"

        $ContentUrl            = "https://teams.sailing-ninoa.com"
        $TeamsAppExternalId    = "2a357162-7738-459a-b727-8039af89a684"

        $tabPayload = @{
            TeamId             = $teamId
            TenantId           = $tenantId
            ChannelName        = ""                 # Standardkanal ""
            TabDisplayName     = $TabDisplayName
            EntityId           = $TabDisplayName    #"AITab"
            ContentUrl         = $ContentUrl
            WebsiteUrl         = $ContentUrl
            TeamsAppExternalId = $TeamsAppExternalId
        } | ConvertTo-Json -Depth 5 -Compress

        # Funktions-URL & Key aus App Settings (lokal: local.settings.json; Azure: Configuration)
        $teamsTabUrl = $env:TEAMS_TAB_FUNC_URL
        $teamsTabKey = $env:TEAMS_TAB_FUNC_KEY

        if ([string]::IsNullOrWhiteSpace($teamsTabUrl)) { throw "TEAMS_TAB_FUNC_URL ist nicht gesetzt." }

        # Für authLevel=function den Key anhängen
        if (-not [string]::IsNullOrWhiteSpace($teamsTabKey)) {
            if ($teamsTabUrl -notmatch '\?') { $teamsTabUrl += "?code=$teamsTabKey" }
            else { $teamsTabUrl += "&code=$teamsTabKey" }
        }

        Log "TeamsTab HTTP aufrufen: $teamsTabUrl"
        try {
            $resp = Invoke-RestMethod -Method POST -Uri $teamsTabUrl `
                -ContentType 'application/json; charset=utf-8' `
                -Body $tabPayload -TimeoutSec 120
            Log "TeamsTab Response: $($resp | ConvertTo-Json -Compress)"
        }
        catch {
            $e = $_.Exception
            $msg = $e.Message
            try {
                $respStream = $e.Response.GetResponseStream()
                $reader = New-Object System.IO.StreamReader($respStream)
                $bodyText = $reader.ReadToEnd()
                if ($bodyText) { $msg = "$msg`n$bodyText" }
            } catch {}
            throw "Aufruf TeamsTab fehlgeschlagen: $msg"
        }

        # --------------------------------------------------------------------
        # gesamte Funktion beenden
        # --------------------------------------------------------------------
        Log "✅ Teams Tab in '$siteTitle' created successfully."
        Send-Resp 200 @{ status = 'success'; siteUrl = $siteUrl }
        return
    }

    # ====================================================================
    # Start PnP PowerShell Modules for SharePoint Provisioning
    # --------------------------------------------------------------------
    # Module laden (PnP.PowerShell)
    # --------------------------------------------------------------------
    #$PnPVersion = "1.12.0"
    $PnPVersion = "3.1.0"

    Log "LoadPnPPSModule -PnPVersion $PnPVersion -FktPath $modulesDir"
    LoadPnPPSModule -PnPVersion $PnPVersion -FktPath $modulesDir

    # ---------------------------------------------------------------
    Log "App-Only Login to SharePoint ADMIN"
    Connect-PnP -Tenant $tenantId -SPOUrl $adminUrl

    # --------------------------------------------------------------------
    # Check if Sharepoint Site (M365-Gruppe) already exists
    # --------------------------------------------------------------------
    Log "Check if Sharepoint Site (M365-Gruppe) '$alias' already exists..."
    $result = Test-SpoSiteOrAliasExists -TenantName $TenantName -Alias $alias -ReturnFirstUrlOnly
    if ($result.Exists) {
        Log "⚠️ Sharepoint Site (M365-Gruppe) '$alias' already exists, use it..."
        $siteUrl = $result.SiteUrl         # URL der Sharepoint-Site for later use

        # Mit der Site verbinden
        Log "App-Only Login to SharePoint SITE $alias"
        Connect-PnP -Tenant $tenantId -SPOUrl $siteUrl

        Log "📎 Get SPOGroupInfo for site '$siteUrl' ..."
        $SPOGroupInfo = Get-SPOGroupInfo -siteUrl $siteUrl
        #Log "SPOGroupInfo: $($SPOGroupInfo | ConvertTo-Json -Compress)"

        $Type = $SPOGroupInfo.Type
        if($Type -ne "TeamSite") {
            Log "⚠️ Site Type is not TeamSite, but $Type. Cannot create a new TeamSite."
            Send-Resp 400 @{ error = "Site Type is not TeamSite, but $Type. Cannot go on with provisioning!" }
            return
        }

        $groupId     = $SPOGroupInfo.GroupId
        $grp         = Get-PnPMicrosoft365Group -Identity $groupId
        $grpName     = $grp.DisplayName

        $SetRegion   = [string]$SPOGroupInfo.Lcid       # Region (z.B. 1031 = Deutsch) for later use

        $alias       = $SPOGroupInfo.Alias              # for later use
        $teamId      = $alias                           # for later use
        $SPOUpdate   = $true                            # for later use

        Log "Sharepoint-Site $siteTitle already exists, update provisioning ..."

        # Users temporär entfernen, dann neu provisionieren, dann wieder hinzufügen...
        $finalMembers = Get-PnPMicrosoft365GroupMember -Identity $GroupId -ErrorAction SilentlyContinue
        $finalMembers = $finalMembers |
                    ForEach-Object { $_.UserPrincipalName ?? $_.Mail } |
                    Where-Object { $_ } |
                    ForEach-Object { $_.Trim().ToLower() }
        Log "📋 Aktuelle Members: $($finalMembers -join ', ')"

        if ($finalMembers.Count -gt 0) {
            Log "Remove all members from group '$grpName' ..."
            Remove-PnPMicrosoft365GroupMember -Identity $groupId -Users $finalMembers
            Log "📋 Actual Members removed but stored."
        } else {
            Log "Gruppe '$grpName' ist bereits leer."
        }

        Log "👑/👥 Set Creator as Owner..."
        if ($creators) {
            Set-M365GroupOwners -GroupId $groupId -Users $creators
        }
        else {
            Log "ℹ️ No owners specified, skipping further steps."
            Send-Resp 400 @{ error = "No owners specified for group '$alias'" }
        }

        $finalOwners  = Get-PnPMicrosoft365GroupOwner -Identity $GroupId -ErrorAction SilentlyContinue
        $finalOwners = $finalOwners |
                    ForEach-Object { $_.UserPrincipalName ?? $_.Mail } |
                    Where-Object { $_ } |
                    ForEach-Object { $_.Trim().ToLower() } |
                    Sort-Object -Unique
        Log "📋 Aktuelle Owners  : $($finalOwners -join ', ')"

        # 3) Vergleichslisten normalisieren
        [string[]]$creatorsNorm = $creators |
                    ForEach-Object { $_.Trim().ToLowerInvariant() } |
                    Sort-Object -Unique
        Log "📋 Creators normiert: $($creatorsNorm -join ', ')"

        $toRemove  = $finalOwners | Where-Object { $creatorsNorm -notcontains $_ }
        Log "📋To Remove: $($toRemove -join ', ')"

        if ($null -eq $toRemove -or $toRemove.Count -eq 0) {
            Log "Gruppe '$grpName' ist bereits leer."
        } else {
            Log "Remove all owners from group '$grpName' ..."
            Remove-PnPMicrosoft365GroupOwner -Identity $groupId -Users $toRemove
            Log "📋 Actual Owners removed but stored."
        }

        if($Privacy -ne $SPOGroupInfo.Privacy) {
            # Mit der Admin-Site verbinden
            Log "App-Only Login to SharePoint ADMIN"
            Connect-PnP -Tenant $tenantId -SPOUrl $adminUrl

            If ($Privacy -eq "Private") {
                Set-PnPMicrosoft365Group -Identity $groupId -IsPrivate:$true  
            } else {
                Set-PnPMicrosoft365Group -Identity $groupId -IsPrivate:$false  
            }
            Log "⚠️ Privacy setting for site '$siteTitle' changed from '$($SPOGroupInfo.Privacy)' to '$Privacy'."

            # Mit der Site verbinden
            Log "App-Only Login to SharePoint SITE $alias"
            Connect-PnP -Tenant $tenantId -SPOUrl $siteUrl

        }

    } else {
        # ------------------------------ Sharepoint-Site anlegen -----------------------------------
        Log "Sharepoint Site (M365-Gruppe) '$alias' does not exists."
        Log "Create Sharepoint-Site $siteTitle for Region: '$SetRegion' ..."
        $Region    = [int]$SetRegion
        if($Privacy -eq "Public") {
            Log "Create Public Sharepoint-Site $siteTitle ..."
            New-PnPSite -Type TeamSite -IsPublic -Title $siteTitle -Alias $alias -Lcid $Region -Wait -ErrorAction Stop
        } else {
            Log "Create Private Sharepoint-Site $siteTitle ..."
            New-PnPSite -Type TeamSite -Title $siteTitle -Alias $alias -Lcid $Region -Wait -ErrorAction Stop
        }   
        $SPOUpdate = $false
        Log "✅ Site created"
    }

    # ------------------- optional: Sharepoint-Site Einstellungen ändern ----------------------
    if($SetTimezone) {
        Log "Set custom TimeZone: $SetTimezone, SortOrder: $SetSortOrder, Region: $SetRegion ..."
        # Mit der neuen Site verbinden
        Log "App-Only Login to SharePoint SITE $siteTitle"
        Connect-PnP -Tenant $tenantId -SPOUrl $siteUrl

        # 3. Regionale Einstellungen setzen (Beispiele)
        $Timezone  = [int]$SetTimezone
        $SortOrder = [int]$SetSortOrder

        Set-PnPRegionalSettings `
            -TimeZone $Timezone `   # 4 = Mitteleuropäische Zeit (Berlin, Wien, Zürich)
            -SortOrder $SortOrder ` # 25 = Deutsch
            -LocaleId $Region `     # 1031 = Deutsch
            -CalendarType Gregorian

        # Mit der ADMIN Site verbinden
        Log "App-Only Login to SharePoint ADMIN"
        Connect-PnP -Tenant $tenantId -SPOUrl $adminUrl
    }

    # --------------------------------------------------------------------
    # 7) Hub-Zuordnung
    # --------------------------------------------------------------------
    $hub = Get-PnPHubSite | Where-Object { $_.Title -eq $hubName }
    Log "Join Hub $hub ..."
    if ($hub) {
        Add-PnPHubSiteAssociation -Site $siteUrl -HubSite $hub.SiteUrl
        Log "✅ Hub joined"
    }
    else {
        Log "⚠️ Hub '$hubName' not found – skipped"
    }

    # --------------------------------------------------------------------
    # 8) Sharepoint Site Collection GroupId ermitteln (Retry 5× à 4 s)
    # --------------------------------------------------------------------
    Log "📎 Find GroupId of a Sharepoint Site Collection..."

    $groupId = $null
    for ($i = 1; $i -le 5 -and -not $groupId; $i++) {
        $gSite = Get-PnPTenantSite -Url $siteUrl -ErrorAction SilentlyContinue
        if ($gSite -and $gSite.GroupId -ne [guid]::Empty) { $groupId = $gSite.GroupId; break }
        Start-Sleep -Seconds 4
    }
    if (-not $groupId) {
        Log "⚠️ GroupId not available – aborting role/Teams/structure step"
        Send-Resp 500 @{ error = "Site '$siteTitle', GroupId not found" }
        return
    }

    Log "📎 GroupId = $groupId"

    # ----------------------------------------------------------------
    # Standard Library der Gruppe triggern und Erstellung abwarten
    # ----------------------------------------------------------------
    Log "Standard Library der Gruppe triggern und Erstellung abwarten..."

    $driveResult = Wait-ForGroupDrive -groupId $groupId -DriveName ""
    $drive       = $driveResult.drive
    $libName     = $drive.name  # Bibliotheksname 

    Log "Standard Library: '$libName' erstellt und verfügbar"
    
    # --------------------------------------------------------------------
    # Creator als Owner hinterlegen
    # --------------------------------------------------------------------
    Log "👑/👥 Set Creator as Owner..."
    if ($creators) {
        Set-M365GroupOwners -GroupId $groupId -Users $creators
    }
    else {
        Log "ℹ️ No owners specified, skipping further steps."
        Send-Resp 400 @{ error = "No owners specified for group '$alias'" }
    }

    # --------------------------------------------------------------------
    # Teams-Team erstellen (falls noch nicht vorhanden)
    # --------------------------------------------------------------------
    Log "ℹ️ Create Teams Team $alias ..."

    $maxTries = 20
    $waitSeconds = 3

    try {
        $team = Get-PnPTeamsTeam -Identity $groupId -ErrorAction SilentlyContinue
        if (-not $team) {
            New-PnPTeamsTeam -GroupId $groupId -ErrorAction Stop
            for ($i=1; $i -le $maxTries; $i++) {
                try {
                    $team = Get-PnPTeamsTeam -Identity $groupId -ErrorAction Stop
                    if ($team) {
                        Log "📢 Teams Team $alias created (after $i Tries)"
                        break
                    }
                } catch {
                    Log "⌛ Wait for Team $alias (Try $i)..."
                    Start-Sleep -Seconds $waitSeconds
                }
            }
        } else {
            Log "ℹ️ Teams Team already exists"
        }                
    } 
    catch {
        Log "⚠️ Teams Team creation failed: $($_.Exception.Message)"
    }
    
    # ==============================================================
    # Ab hier: SITE Provisioning - App-Only Login to SharePoint SITE
    # ==============================================================
    if ($enableProvisioning -or $EnableDocumentSets) {
        Log "App-Only Login to SharePoint SITE"
        Connect-PnP -Tenant $tenantId -SPOUrl $siteUrl

        Log "Start XML Provisioning for Site '$siteTitle' ..."
        $SiteTemplates = Join-Path $PSScriptRoot 'provisioning\01_Site'  # Pfad zu XML-Dateien

        # ----------------------------------------------------------------
        # Wenn eigener DMSdrive angegeben ist, dann DMS-Bibliothek anlegen
        # ----------------------------------------------------------------
        if ($null -ne $DMSdrive) {
            Log "DMS Library '$DMSdrive' erstellen..."

            New-PnPList -Title $DMSdrive -Template DocumentLibrary -Url $DMSdrive
            $driveResult = Wait-ForGroupDrive -groupId $groupId -DriveName $DMSdrive
            $drive       = $driveResult.drive
            $libName     = $drive.name  # Bibliotheksname 

            Log "DMS Library: '$libName' erstellt und verfügbar"
        }
    }
        
    # ---------------------------------------------------------------
    # DMS Bibliothek (entweder Standard oder eigene DMS) - Variablen festlegen:    
    # ---------------------------------------------------------------
    $rootItem    = $driveResult.rootItem
    $rootId      = $rootItem.id  
    $driveId     = $drive.id
    
    # ---------------------------------------------------------------------
    # Document Set Feature auf der Site aktivieren für Projekt-DMS-Features
    # ---------------------------------------------------------------------
    if (-not $EnableDocumentSets) { 
        Log "Aktivieren von Document Set Feature auf $siteTitle skipped."
    } else {
        Log "Document Set Feature auf der Site aktivieren..."
        Enable-DocumentSets

        # TERM SETS XML Site Provisioning
        $XMLSchema = "TermSets_2022.xml"
        Log "📋 Bereitstellen von $XMLSchema"
        ProvisionPnPSite -siteUrl $siteUrl -templateFolder $SiteTemplates -template $XMLSchema
    }
    
    # ------------------------------------------------------------------
    # Provisioning der Site Collection für DMS Funktionen
    # ------------------------------------------------------------------
    if (-not $enableProvisioning) {
        Log "Provisioning is disabled for Site '$siteTitle' – skipping Provisioning steps"
        #Send-Resp 200 @{ status = 'success'; siteUrl = $siteUrl }
        #return
    } else {

        # ------------------------------------------------------------------
        # Site Provisioning mit PnP XML-Vorlagen:
        # Term Sets, Site Columns, Content Types
        # ------------------------------------------------------------------
        Log "Site Provisioning mit PnP XML-Vorlagen..."

        # SITE COLUMNS PREPROCESSED
        $XMLSchema = "SiteCollection_SiteColumns_Preprocessed_2022.xml"
        Log "📋 Bereitstellen von $XMLSchema"
        ProvisionPnPSite -siteUrl $siteUrl -templateFolder $SiteTemplates -template $XMLSchema

        # CONTENT TYPES
        $XMLSchema = "SiteCollection_ContentTypes_2022.xml"
        Log "📋 Bereitstellen von $XMLSchema"
        ProvisionPnPSite -siteUrl $siteUrl -templateFolder $SiteTemplates -template $XMLSchema

        # SITE COLUMNS CALCULATED
        $XMLSchema = "SiteCollection_SiteColumns_Calculated_2022.xml"
        Log "📋 Bereitstellen von $XMLSchema"

        # Wenn SPOUpdate = true, dann müssen die berechneten Felder entfernt und neu angelegt werden
        if($SPOUpdate) {
            # Welche berechneten Site Columns machen Ärger?
            # !!! Berechnete Felder künftig aus der XML Datei lesen !

            $calcFields = @("mvSubject","mvHasAttachments","mvAttachmentIcon")
            Log "🔧 Berechnete Felder: $($calcFields -join ', ')"

            # --- DETACH: alle betroffenen Listen einsammeln und Feldbindungen lösen ---
            Log "Entferne temporär die Felder: $($calcFields -join ', ') aus den betroffenen Listen..."
            $affected = @{}
            foreach ($f in $calcFields) {
                $lists = Find-ListsUsingField -FieldInternalName $f
                $affected[$f] = $lists
                foreach ($l in $lists) {
                    Log "🔧 Entferne temporär '$f' aus Liste '$($l.Title)' ($($l.RootFolder.ServerRelativeUrl))"
                    DetachFieldFromList -List $l -FieldInternalName $f
                }
            }

            # --- UPDATE: jetzt Calculated-Template anwenden ---
            ProvisionPnPSite -siteUrl $siteUrl -templateFolder $SiteTemplates -template $XMLSchema

            # --- REATTACH: Felder wieder an die vormals betroffenen Listen binden ---
            Log "🔧 Binde die Felder: $($calcFields -join ', ') wieder an die betroffenen Listen..."
            foreach ($f in $calcFields) {
                foreach ($l in $affected[$f]) {
                    Log "↩️ Binde '$f' wieder an Liste '$($l.Title)'"
                    ReattachFieldToListViaDocumentCT -List $l -FieldInternalName $f
                }
            }
            Log "✅ Detach→Update→Reattach abgeschlossen."
        } else {
            # --- NEW: Calculated-Template anwenden ---
            ProvisionPnPSite -siteUrl $siteUrl -templateFolder $SiteTemplates -template $XMLSchema
        }

        Log "✅ Site Provisioning abgeschlossen!"

        # --------------------------------------------------------
        # Feature Content Types in der Bibliothek aktivieren
        # --------------------------------------------------------
        $lib = Get-PnPList -Identity $libName
        $libID = $lib.Id
        
        if ($libID) {
            Log "Content Types in der Bibliothek: '$libName', ID: '$libID' aktivieren ..."
            Set-PnPList -Identity $libID -EnableContentTypes $true

            # Vorhandene Content Types der Site der Bibliothek zuweisen
            #$docSetName = "Document Set"
            #Add-ContentType -libName $libName -docSetName $docSetName

            $docSetName = "Email"
            Add-ContentType -libName $libName -docSetName $docSetName

            $docSetName = "MacroView Document"
            Add-ContentType -libName $libName -docSetName $docSetName

            #Log "Optional: Ein erstes Document Set anlegen"
            #Add-PnPDocumentSet -List $libName -Name $docSetName -ContentType $docSetCT.Id
        }
        else {
            Log "❌ Bibliothek: '$libName', ID: '$libID' nicht gefunden!"
            Send-Resp 500 @{ error = "Library '$libName', ID '$libID'  not found" }
            return
        }

        # ---------------------------------------------------------------
        # nicht benötigte Library Elemente merken
        # ---------------------------------------------------------------
        if ($enableCleanUp) {
            $viewFilter = "*All*Do?ument*"  # Views finden, die nicht benötigt werden (de/en)
            $viewItems  = Get-PnPView -List $libName | Where-Object { $_.Title -like $viewFilter } | Select-Object -ExpandProperty Id

            $ctFilter = "Do?ument*"     # Content Types finden, die nicht benötigt werden
            $ctItems  = Get-PnPContentType -List $libName | Where-Object { $_.Name -like $ctFilter } | Select-Object -ExpandProperty Id
        }

        # ---------------------------------------------------------------
        # Library Provisioning mit PnP XML-Vorlagen für Bibliotheken
        # ---------------------------------------------------------------
        $LibEnableProvisioning = $true  # Setze auf $true, um Document Library Provisioning zu starten
        if ($LibEnableProvisioning) { 
            # LIBRARY PROVISIONING (upgrade)
            $LibraryTemplates = Join-Path $PSScriptRoot 'provisioning\02_Library'  # Pfad zu XML-Dateien

            Log "Starte Library Provisioning mit PnP XML-Vorlagen..."
            if( -not $DMSdrive) {
                Log "ℹ️ Custom DMSdrive is not set, using default library '$libName' for provisioning."
                if ($SetRegion -eq "1031") {
                    $XMLSchema = "Dokumente_DE_2022.xml"
                    Log "Region '$SetRegion' : Deutsch (Deutschland) - Verwende $XMLSchema"
                } elseif ($SetRegion -eq "1033") {
                    $XMLSchema = "Documents_EN_2022.xml"
                    Log "Region '$SetRegion' : English - Verwende $XMLSchema"
                } else {
                    Log "Region '$SetRegion' : No XMLSchema available for this region, using English default library schema."
                    $XMLSchema = "Documents_EN_2022.xml"
                    Log "Region '$SetRegion' : English - Verwende $XMLSchema"
                } 
            } else {
                    $XMLSchema = "Library_XX_2022.xml"
                    Log "ℹ️ Custom DMSdrive: '$DMSdrive' set, using library '$DMSdrive' for provisioning."
                    Log "Region '$SetRegion' : Verwende $XMLSchema"
                    Log "$XMLSchema has to be customized for your DMS drive..."
                    #CustomizeLibXML -XMLSchema $XMLSchema -DMSdrive $DMSdrive -SetRegion $SetRegion -SetTimezone $SetTimezone -SetSortOrder $SetSortOrder
                    Log "Customization is not ready yet, skipping Library Provisioning for custom DMS drive '$DMSdrive'."
                }
            
            Log "📋 Bereitstellen von $XMLSchema"
            ProvisionPnPSite -siteUrl $siteUrl -templateFolder $LibraryTemplates -template $XMLSchema

            Log "✅ Library Provisioning abgeschlossen!"
        }

        # ---------------------------------------------------------------
        # nicht benötigte Library Elemente löschen
        # ---------------------------------------------------------------
        if ($enableCleanUp) {
            Log "Entferne nicht benötigte Views: '$viewFilter' in der Bibliothek '$libName' ..."
            if (-not $viewItems) {
                Log "ℹ️ Keine Views gefunden, die gelöscht werden müssen."
            } else {

                $defaultViewid = Get-PnPView -List $libName | Where-Object { $_.DefaultView -eq $true } | Select-Object -ExpandProperty Id

                foreach ($view in $viewItems) {
                    if ($view -eq $defaultViewid) {
                        Log "ℹ️ DefaultView ID: '$defaultViewid'"
                        Log "ℹ️        View ID: '$view' kann nicht gelöscht werden (=Default!)"
                        continue
                    } else {
                        Log "Entferne View ID: '$view' aus der Bibliothek '$libName' ..."
                        Remove-PnPView -List $libName -Identity $view -Force
                    }
                }
            }

            # Content Types löschen, die nicht benötigt werden
            Log "Entferne nicht benötigte Content Types: '$ctFilter' in der Bibliothek '$libName' ..."
            if (-not $ctItems) {
                Log "ℹ️ Keine Content Types gefunden, die gelöscht werden müssen."
            } else {
                $list = Get-PnPList -Identity $libName
                # ContentTypes explizit laden!
                $ctx = Get-PnPContext
                $ctx.Load($list.ContentTypes)
                $ctx.ExecuteQuery()

                $defaultCTid = Get-PnPContentType -List $libName | Select-Object -First 1 -ExpandProperty Id

                foreach ($ct in $ctItems) {
                    if ($ct -eq $defaultCTid) {
                        Log "ℹ️ Default Content Type ID: '$defaultCTid'"
                        Log "ℹ️         Content Type ID: '$ct' kann nicht gelöscht werden (Default!)"
                    } else {
                        Log "Entferne Content Type ID: '$ct' aus der Bibliothek '$libName' ..."
                        $ctobj = $list.ContentTypes | Where-Object { $_.id.StringValue -eq $ct }
                        $ctobj.DeleteObject()
                        $ctx.ExecuteQuery()
                    }
                }
            }
        }
    }

    # ------------------------------------------------------------------------
    # Ordnerstruktur für Projekt anlegen (option: unter General/Allgemein)
    # ------------------------------------------------------------------------
    if (-not $structure) {
        Log "Keine Ordnerstruktur zum Anlegen angegeben. Skipping folder structure creation."
    } else {
        Log "Ordnerstruktur anlegen ..."

        # ==============================================================
        # Ab hier: App-Only Login to SharePoint ADMIN
        # ==============================================================
        Log "App-Only Login to SharePoint ADMIN"
        Connect-PnP -Tenant $tenantId -SPOUrl $adminUrl

        if (($structureInDefaultChannelFolder -eq $true) -and ($null -eq $DMSdrive)) {
            Log "Ordnerstruktur unterhalb von General/Allgemein in Bibliothek $LibName anlegen ..."
            # === Hier das neue Graph-API-Verfahren ===
            $teamId = $groupId

            # Channels holen
            $channelsResp = Invoke-PnPGraphMethod -Method GET `
                -Url "https://graph.microsoft.com/v1.0/teams/$teamId/channels"

            # Ersten Kanal nehmen (Standardkanal/General)
            $generalChannel = $channelsResp.value | Sort-Object createdDateTime | Select-Object -First 1
            $channelFolderName = $generalChannel.displayName

            # Hole alle Ordner im Library-Root
            $childrenUrl = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$rootId/children"
            $existingChildren = Invoke-PnPGraphMethod -Method GET -Url $childrenUrl

            # ================= NEU: Robust Channel-Ordner anlegen =================
            $channelFolder = $existingChildren.value | Where-Object {
                $_.folder -ne $null -and ($_.name.Trim().ToLower() -eq $channelFolderName.Trim().ToLower())
            }

            if (-not $channelFolder) {
                $bodyJson = @"
{
    "name": "$channelFolderName",
    "folder": {},
    "@microsoft.graph.conflictBehavior": "rename"
}
"@

                try {
                    Log "➕ Erstelle neuen Channel-Ordner '$channelFolderName' (POST mit ConflictBehavior)..."
                    $channelFolder = Invoke-PnPGraphMethod -Method POST `
                        -Url "https://graph.microsoft.com/v1.0/drives/$driveId/items/$rootId/children" `
                        -Content $bodyJson -ContentType "application/json"
                    Log "✅ Ordner erfolgreich angelegt: $($channelFolder.name)"
                    Start-Sleep -Seconds 3
                }
                catch {
                    # 409 = "Conflict", "Name already exists" – dann verwende den existierenden Ordner!
                    if ($_.Exception.Message -match '(Conflict: Name already exists|Status code Conflict)') {
                        Log "ℹ️ Ordner '$channelFolderName' existiert bereits – suche existierenden Ordner..."
                        $existingChildren = Invoke-PnPGraphMethod -Method GET -Url $childrenUrl
                        $channelFolder = $existingChildren.value | Where-Object {
                            $_.folder -ne $null -and ($_.name.Trim().ToLower() -eq $channelFolderName.Trim().ToLower())
                        }
                        if ($channelFolder) {
                            Log "✅ Existierender Ordner gefunden: $($channelFolder.name) (ID: $($channelFolder.id))"
                        } else {
                            Log "❌ Fehler: Ordner existiert laut Graph, aber im Listing nicht gefunden!"
                            throw
                        }
                    } else {
                        Log "❌ Unerwarteter Fehler beim Ordner-Anlegen: $($_.Exception.Message)"
                        throw
                    }
                }
            } else {
                Log "ℹ️ Channel-Ordner '$channelFolderName' existiert bereits (ID: $($channelFolder.id)), keine Neuanlage notwendig."
            }

            # Jetzt wie gehabt weiter...
            Add-Folders -items $structure -parentId $channelFolder.id -driveId $driveId
            Log "✅ Folder structure provisioned under default channel folder '$($channelFolder.name)'"        
        } else {
            Log "Ordnerstruktur in Bibliothek $LibName direkt anlegen ..."
            Add-Folders -items $structure -parentId $rootId -driveId $driveId
            Log "✅ Folder structure provisioned in library root"
        }
    }

    # --------------------------------------------------------------------
    # Teams Tab anlegen in eigener Function (HTTP POST Call)
    # --------------------------------------------------------------------
    if ($TeamsTabAfterProvisioning -and $enableTabCreation) { 
        Log "---- TeamsTab-Function callen (HTTP) ----"

        $ContentUrl = "https://teams.sailing-ninoa.com"
        $TeamsAppExternalId  = "2a357162-7738-459a-b727-8039af89a684"

        $tabPayload = @{
            TeamId             = $teamId
            TenantId           = $tenantId
            ChannelName        = ""
            TabDisplayName     = $TabDisplayName
            ContentUrl         = $ContentUrl
            WebsiteUrl         = $ContentUrl
            EntityId           = "AITab"
            TeamsAppExternalId = $TeamsAppExternalId
        } | ConvertTo-Json -Depth 5 -Compress

        # Funktions-URL & Key aus App Settings (lokal: local.settings.json; Azure: Configuration)
        $teamsTabUrl = $env:TEAMS_TAB_FUNC_URL
        $teamsTabKey = $env:TEAMS_TAB_FUNC_KEY

        if ([string]::IsNullOrWhiteSpace($teamsTabUrl)) { throw "TEAMS_TAB_FUNC_URL ist nicht gesetzt." }

        # Für authLevel=function den Key anhängen
        if (-not [string]::IsNullOrWhiteSpace($teamsTabKey)) {
            if ($teamsTabUrl -notmatch '\?') { $teamsTabUrl += "?code=$teamsTabKey" }
            else { $teamsTabUrl += "&code=$teamsTabKey" }
        }

        Log "TeamsTab HTTP aufrufen: $teamsTabUrl"
        try {
            $resp = Invoke-RestMethod -Method POST -Uri $teamsTabUrl `
                -ContentType 'application/json; charset=utf-8' `
                -Body $tabPayload -TimeoutSec 120
            Log "TeamsTab Response: $($resp | ConvertTo-Json -Compress)"
        }
        catch {
            $e = $_.Exception
            $msg = $e.Message
            try {
                $respStream = $e.Response.GetResponseStream()
                $reader = New-Object System.IO.StreamReader($respStream)
                $bodyText = $reader.ReadToEnd()
                if ($bodyText) { $msg = "$msg`n$bodyText" }
            } catch {}
            throw "Aufruf TeamsTab fehlgeschlagen: $msg"
        }
    }

    # --------------------------------------------------------------------
    # Owners / Members hinterlegen
    # --------------------------------------------------------------------
    Log "👑/👥 Set Owners ..."
    if ($owners) {Set-M365GroupOwners -GroupId $groupId -Users $owners}
    if ($finalOwners) {Set-M365GroupOwners -GroupId $groupId -Users $finalOwners}

    Log "👑/👥 Set Members ..."
    if ($members) {Set-M365GroupMembers -GroupId $groupId -Users $members}
    if ($finalMembers) {Set-M365GroupMembers -GroupId $groupId -Users $finalMembers}

    Log "👑/👥 Owners/Members set"

    # --------------------------------------------------------------------
    # Fertig
    # --------------------------------------------------------------------
    Log "✅ Site Collection '$siteTitle' created and provisioned successfully!"
    Send-Resp 200 @{ status = 'success'; siteUrl = $siteUrl }

}
catch {
    Write-Error "❌ $($_.Exception.Message)"
    Send-Resp 500 @{ error = $_.Exception.Message }
}
