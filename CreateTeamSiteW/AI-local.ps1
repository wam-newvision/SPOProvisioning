param(
    [string]$Site,                      # Site, z.B. "1100"
    [string]$channelName = "Allgemein", # Name der Dokumentenbibliothek
    [string]$tabName     = "AI Agent"   # Wie der Tab heißen soll
)

# --------------------------------------------------------------------
# Include Helper Functions
# --------------------------------------------------------------------
. (Join-Path $PSScriptRoot 'helpers\PSHelpers.ps1')

function Log {
    Write-Host ($args -join " ")
}

Log "🔧 Test-Skript" $PSScriptRoot "AI-local.ps1 startet..."
# --------------------------------------------------------------------
# Eingaben zuweisen
# --------------------------------------------------------------------
#$FktPath      = "C:\Functions\dms-provisioning\CreateTeamSiteW"
#$FktPath      = $PSScriptRoot

# -----------------------------
# VORRAUSSETZUNGEN:
# - Microsoft.Graph PowerShell Modul installiert:
#   Install-Module Microsoft.Graph -Scope CurrentUser
# - App ZIP-Package von Copilot Studio exportiert
# - Teams Admin/Global Admin Rechte für das Hochladen
# -----------------------------

# VARIABLEN ANPASSEN!
$zipPath      = "C:\Functions\teamsApp.zip"            # Pfad zur ZIP aus Copilot Studio
$teamName     = $Site                                  # Name deines Teams (wie angezeigt in Teams)
#$channelName = "Allgemein"                            # Kanalname (z. B. "Allgemein")
#$tabName     = "AI Agent"                             # Wie der Tab heißen soll

# 1. Microsoft Graph PowerShell laden und anmelden
Import-Module Microsoft.Graph
Connect-MgGraph -Scopes "AppCatalog.ReadWrite.All","TeamsTab.ReadWrite.All","Team.ReadBasic.All","Channel.ReadBasic.All"

# 2. Teams App-Package hochladen
Write-Host "Lade App-Package hoch..."
$zipBytes = [System.IO.File]::ReadAllBytes($zipPath)
$response = Invoke-MgGraphRequest -Method POST `
    -Uri "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps" `
    -Body @{ package = $zipBytes } `
    -ContentType "application/zip"
$teamsAppId = $response.id
Write-Host "App wurde hochgeladen. TeamsAppId: $teamsAppId"

# 3. Team-ID und Kanal-ID suchen
Write-Host "Suche Team und Kanal..."
$team = Get-MgTeam -Filter "displayName eq '$teamName'"
if (!$team) { Write-Error "Team nicht gefunden!"; exit }
$teamId = $team.Id

$channel = Get-MgTeamChannel -TeamId $teamId | Where-Object { $_.DisplayName -eq $channelName }
if (!$channel) { Write-Error "Kanal nicht gefunden!"; exit }
$channelId = $channel.Id

Write-Host "Team-Id: $teamId | Channel-Id: $channelId"

# 4. Tab als Register im Kanal einfügen
Write-Host "Erstelle Tab im Kanal..."
$tabBody = [PSCustomObject]@{
    displayName = $tabName
    "teamsApp@odata.bind" = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps('$teamsAppId')"
    configuration = @{
        entityId = "copilot"
    }
}

$response = Invoke-MgGraphRequest -Method POST `
    -Uri "https://graph.microsoft.com/v1.0/teams/$teamId/channels/$channelId/tabs" `
    -Body ($tabBody | ConvertTo-Json -Depth 10)

Write-Host "Tab '$tabName' wurde erfolgreich als Registerkarte hinzugefügt!"
