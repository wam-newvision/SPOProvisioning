param(
    [string]$Site,  # Site, z.B. "1100"
    [string]$libName = "Dokumente"  # Name der Dokumentenbibliothek
)

# --------------------------------------------------------------------
# Include Helper Functions
# --------------------------------------------------------------------
. (Join-Path $PSScriptRoot 'helpers\SPOAdminFunctions.ps1')
. (Join-Path $PSScriptRoot 'helpers\SPOLibraryFunctions.ps1')
. (Join-Path $PSScriptRoot 'helpers\ProvisionPnP.ps1')
. (Join-Path $PSScriptRoot 'helpers\PSHelpers.ps1')

$env:DEBUG = 'true'
Log "🔧 Test-Skript" $PSScriptRoot "test-local.ps1 startet..."
# --------------------------------------------------------------------
# Eingaben zuweisen
# --------------------------------------------------------------------
#$FktPath      = "C:\Functions\dms-provisioning\CreateTeamSiteW"
$FktPath      = $PSScriptRoot

$ClientId     = "5a19516e-dc54-4d2f-aebc-f1b679a69457"
#$clientSecret = $env:AZURE_CLIENT_SECRET

$tenantId     = "mwpnewvision.onmicrosoft.com"
$siteTitle    = $Site
#$hubName      = "ProjektHub"

$PfxPath      = Join-Path $FktPath 'certs\mwpnewvision.pfx'
$PfxPwd       = "MyP@ssword!" # Setze hier dein PFX-Passwort
$PfxPassword  = (ConvertTo-SecureString $PfxPwd -AsPlainText -Force)

# --------------------------------------------------------------------
# Alias / URLs
# --------------------------------------------------------------------
$base     = $tenantId.Split('.')[0]
$alias    = ($siteTitle -replace '\s+', '')
$siteUrl  = "https://${base}.sharepoint.com/sites/$alias"
$adminUrl = "https://${base}-admin.sharepoint.com"
Log "🔗 SiteUrl = $siteUrl"

# --------------------------------------------------------------------
# Module laden (PnP.PowerShell)
# --------------------------------------------------------------------
$Graphmodules = $true  # Setze auf $true, um Graph Module zu starten und zeigen

if ($Graphmodules) {
    # Installiere das Modul, falls nicht geschehen
    Install-Module Microsoft.Graph -Scope CurrentUser

    # Authentifiziere dich mit dem benötigten Scope
    Connect-MgGraph -Scopes "Group.ReadWrite.All"   #Connect-MgGraph -Scopes "TeamsTab.ReadWriteForTeam.All"

    # Überprüfe die Verbindung (optional)
    # Get-MgUser -UserId me
}

$PnPmodules = $true  # Setze auf $true, um PnP Module zu starten und zeigen

if ($PnPmodules) {
    #$PnPVersion = "1.12.0"
    $PnPVersion = "3.1.0"

    LoadPnPPSModule -PnPVersion $PnPVersion -FktPath $FktPath

    #Import-Module PnP.PowerShell
    #Get-Module -Name PnP.PowerShell

    # App-Only Login to SharePoint Site
    Log "App-Only Login to SharePoint SITE"
    #Connect-PnP -Tenant $tenantId -SPOUrl $siteUrl
    #Connect-PnP -Tenant $tenantId -SPOUrl $siteUrl -ClientId $ClientId -PfxPath $PfxPath -PfxPassword $PfxPassword

    # App-Only Login to SharePoint Admin
    #Connect-PnP -Tenant $tenantId -SPOUrl $adminUrl
    Connect-PnP -Tenant $tenantId -SPOUrl $adminUrl -ClientId $ClientId -PfxPath $PfxPath -PfxPassword $PfxPassword

    # --------------------------------------------------------------------
    # PnP Powershell Module anzeigen
    # --------------------------------------------------------------------
    Get-Module PnP.PowerShell | Format-List Name, Version, ModuleBase
    #(Get-Command Set-PnPList).Parameters.Keys
}

# --------------------------------------------------------------------
# PnP Provisioning XML Template exportieren
# --------------------------------------------------------------------
$PnPexport = $false  # Setze auf $true, um PnP XML zu exportieren

if ($PnPexport) {

    $Category  = "All"
    #$PnPSchema = "V201805"
    #$PnPSchema = "V202103"
    $PnPSchema = "V202209"
    $templateFolder = Join-Path $FktPath 'provisioning'  # Pfad zu XML-Dateien
    $XMLPath = Join-Path $templateFolder "$Site-$Category.xml"  # Pfad zu XML-Dateien
    #$XMLPath = Join-Path "C:\Temp" "$Site-$Category.xml"  # Pfad zu XML-Dateien

    try {
        if ($Category -eq "All") {
            log "Get-PnPSiteTemplate -Out $XMLPath -Schema $PnPSchema"
            Get-PnPSiteTemplate -Out $XMLPath -Schema $PnPSchema
        } else {
            log "Get-PnPSiteTemplate -Out $XMLPath -Schema $PnPSchema -Handlers $Category"
            Get-PnPSiteTemplate -Out $XMLPath -Schema $PnPSchema -Handlers $Category
        }
    } catch {
        Write-Host "Error: $($_.Exception.Message)"
        Write-Host "StackTrace: $($_.Exception.StackTrace)"
        if ($_.Exception.InnerException) {
            Write-Host "Inner Exception: $($_.Exception.InnerException.Message)"
        }
    }

    return
}

# ---------------------------------------------------------------
# Document Set Einrichtung
# ---------------------------------------------------------------
$EnableDocSet = $false  # Setze auf $true, um Document Set Einrichtung zu starten

if ($EnableDocSet) {
    Log "Starte Document Set Einrichtung..."
    Enable-DocumentSets

    # Alternativ über XML: Enable SiteCollection Features
    #$template = "SiteCollectionFeatures.xml"
    #ProvisionPnPSite -siteUrl $siteUrl -templateFolder $templateFolder -template $template
}

# ---------------------------------------------------------------
# Content Types Einrichtung
# ---------------------------------------------------------------
$ContentTypes = $false  # Setze auf $true, um Document Set Einrichtung zu starten

if ($ContentTypes) {
    $lib = Get-PnPList -Identity $libName
    $libID = $lib.Id
    Log "Content Types in der Bibliothek: '$libName', ID: '$libID' aktivieren ..."
    Set-PnPList -Identity $libID -EnableContentTypes $true -ErrorAction Stop 
    #Set-PnPList -Identity $libName -EnableContentTypes $true

    #Log "Content Type: 'Document Set' der Library zuweisen"
    #Add-PnPContentTypeToList -List $libName -ContentType $docSetCT.Id

    $docSetName = "Email"
    Add-ContentType -libName $libName -docSetName $docSetName

    $docSetName = "MacroView Document"
    Add-ContentType -libName $libName -docSetName $docSetName

    Log "DMS Library: '$libName' erstellt und verfügbar"

    return
}

# ---------------------------------------------------------------
# Provisioning Site mit PnP XML-Vorlagen für Term Sets, Site Columns, Content Types und Listen/Bibliotheken
# ---------------------------------------------------------------
$EnableProvisioning = $false  # Setze auf $true, um Document Site Provisioning zu starten
if ($EnableProvisioning) { 
    Log "Starte Site Provisioning Upgrade..."
    $templateFolder = Join-Path $PSScriptRoot 'provisioning\02_Library'  # Pfad zu XML-Dateien

    # SITE COLUMNS PREPROCESSED
    $XMLSchema = "Upgrade.xml"
    Log "📋 Bereitstellen von $XMLSchema"
    ProvisionPnPSite -siteUrl $siteUrl -templateFolder $templateFolder -template $XMLSchema

    # TERM SETS
    Log "📚 Bereitstellen von Term Sets..."
    $template = "TermSets.xml"
    ProvisionPnPSite -siteUrl $siteUrl -templateFolder $templateFolder -template $template

    # SITE COLUMNS
    Log "📋 Bereitstellen von Site Columns..."
    $template = "SiteColumns.xml"
    ProvisionPnPSite -siteUrl $siteUrl -templateFolder $templateFolder -template $template

    # TAXONOMY (Term Store)
    Log "📚 Bereitstellen von XMLs..."
    $template = "SiteColumns_Taxonomy.xml"
    ProvisionPnPSite -siteUrl $siteUrl -templateFolder $templateFolder -template $template

    # CONTENT TYPES (inkl. Project Document Set)
    Log "📦 Bereitstellen von Content Types..."
    $template = "SiteContentTypes.xml"
    ProvisionPnPSite -siteUrl $siteUrl -templateFolder $templateFolder -template $template

    # LISTEN / BIBLIOTHEKEN mit Project Document Set
    Log "📁 Bereitstellen von Listen und Bibliotheken..."
    $template = "SiteLists.xml"
    #ProvisionPnPSite -siteUrl $siteUrl -templateFolder $templateFolder -template $template

    # LISTEN / BIBLIOTHEKEN mit Project Document Set
    Log "📁 Bereitstellen von Views..."
    $template = "SiteLists_Views.xml"
    ProvisionPnPSite -siteUrl $siteUrl -templateFolder $templateFolder -template $template

return
}

# ---------------------------------------------------------------
# nicht benötigte Library Elemente merken
# ---------------------------------------------------------------
$CleanUp = $false  # Setze auf $true, um nicht benötigte Library Elemente zu löschen
if ($CleanUp) {
    # Views finden, die nicht benötigt werden
    $viewFilter = "*All*Do?ument*"
    #Log "Finde Views, die nicht benötigt werden: '$viewFilter' in der Bibliothek '$libName' ..."
    # Gefilterten Views sammeln
    $viewItems = Get-PnPView -List $libName | Where-Object { $_.Title -like $viewFilter } | Select-Object -ExpandProperty Id
    #$viewItems    # Anzeigen der Views, die gelöscht werden sollen

    # Content Types löschen, die nicht benötigt werden
    $ctFilter = "Do?ument*"
    #Log "Finde Content Types, die nicht benötigt werden: '$ctFilter' in der Bibliothek '$libName' ..."
    # IDs der gefilterten Content Types sammeln
    $ctItems = Get-PnPContentType -List $libName | Where-Object { $_.Name -like $ctFilter } | Select-Object -ExpandProperty Id
    #$ctItems  # Anzeigen der Content Types, die gelöscht werden sollen
}

# ---------------------------------------------------------------
# Add Teams Tab
# ---------------------------------------------------------------
$TeamsTab = $true  # Setze auf $true, um nicht benötigte Library Elemente zu löschen
if ($TeamsTab) {

    Log "Starte Teams Tab Einrichtung..."
    $team = Get-PnPTeamsTeam -Identity $alias -ErrorAction Stop

    if (-not $team) {
        Log "❌ Teams Team '$alias' nicht gefunden!"
        return
    } else {
        Log "Teams Team '$($team.DisplayName)' gefunden! (ID: $($team.GroupId))"
    }

    # Teams Channel finden
    $channel = Get-PnPTeamsChannel -Team $team.GroupId | Select-Object -First 1
    if (-not $channel) {
        Log "❌ Kein Channel für Team '$alias' gefunden!"
        return
    }

    Log "📢 Default Channel: $($channel.DisplayName) (ID: $($channel.Id))"

    #$TabType = [TabType]::SharePointPageAndList  # Typ des Tabs (z.B. SharePointPageAndList, WebSite, etc.)
    #$TabType = [TabType]::WebSite  # Typ des Tabs (z.B. SharePointPageAndList, WebSite, etc.)
    $TabDisplayName = "AI Agent"  # Name des Tabs
    #$TeamsTabURL = "https://newvision.eu/impressum/"
    $TeamsTabURL = "https://mwpnewvision.sharepoint.com/sites/$alias/SitePages/Forms/ByAuthor.aspx"
    #$WebSiteUrlDisplayName = "NewVision"  # DisplayName der Website, die im Tab angezeigt werden soll

    # Tab hinzufügen
    Log "WebSite Tab hinzufügen zu Team: '$($team.DisplayName)' ..."

<#
    AddTeamsTab `
        -team $team `
        -TeamsChannel $channel `
        -TabDisplayName $TabDisplayName `
        -WebSiteUrl $TeamsTabURL `
        -TabType WebSite
#>

    $TeamsAppId = "2a357162-7738-459a-b727-8039af89a684"  # App-ID der zugehörigen Teams-App (z.B. Copilot Studio Bot)
    $TeamsTabURL = "https://teams.sailing-ninoa.com"

    Add-GraphTeamsTab `
        -TeamId $Team.GroupId `
        -ChannelId $Channel.Id `
        -TabDisplayName $TabDisplayName `
        -TeamsAppId $TeamsAppId `  # App-ID der zugehörigen Teams-App
        -EntityId "teaminfotab" `  # i. d. R. „copilot“ für Copilot Studio Bots
        -ContentUrl $TeamsTabURL `  # Content URL wie im App-Manifest
        -WebsiteUrl $TeamsTabURL  # Optional: zusätzlicher Website-Link 

}

# ---------------------------------------------------------------
# nicht benötigte Library Elemente löschen
# ---------------------------------------------------------------
#$CleanUp = $false  # Setze auf $true, um nicht benötigte Library Elemente zu löschen
if ($CleanUp) {
    # Views löschen, die nicht benötigt werden
    Log "Entferne nicht benötigte Views: '$viewFilter' in der Bibliothek '$libName' ..."
    if (-not $viewItems) {
        Log "ℹ️ Keine Views gefunden, die gelöscht werden müssen."
    } else {
        $defaultViewid = Get-PnPView -List $libName | Where-Object { $_.DefaultView -eq $true } | Select-Object -ExpandProperty Id
        Log "ℹ️ DefaultView ID: '$defaultViewid'"
        foreach ($view in $viewItems) {
            if ($view -eq $defaultViewid) {
                Log "ℹ️ DefaultView ID: '$view' kann nicht gelöscht werden!"
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
        Log "Default Content Type ID: '$defaultCTid'"

        foreach ($ct in $ctItems) {
            if ($ct -eq $defaultCTid) {
                Log "ℹ️ Default Content Type ID: '$ct' kann nicht gelöscht werden!"
            } else {
                Log "Entferne Content Type ID: '$ct' aus der Bibliothek '$libName' ..."
                #Log "$list" ".ContentTypes.Delete(" $ct.StringValue ")"
                $ctobj = $list.ContentTypes | Where-Object { $_.id.StringValue -eq $ct }
                $ctobj.DeleteObject()
                #$list.Update()
                $ctx.ExecuteQuery()
            }
        }
    }
}


Log "✅ Bereitstellung abgeschlossen!"
