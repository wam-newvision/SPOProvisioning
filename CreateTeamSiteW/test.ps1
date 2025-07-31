# Graph-Test.ps1 (nur Basisfunktionen!)
Import-Module Microsoft.Graph

# Nur die Scopes, die garantiert existieren!
Connect-MgGraph -Scopes "Team.ReadBasic.All","Channel.ReadBasic.All"

# Teams abrufen
$teams = Get-MgTeam -Top 5
$teams | Select-Object DisplayName,Id

