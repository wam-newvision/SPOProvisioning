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
# Document Set (and Content Type) auf der Site Collection aktivieren
# ------------------------------------------------------------------------------------
function Enable-DocumentSets {
    param(
        #[string]$featureId,  # Feature ID, z.B. "b50e3104-6812-424f-a011-cc90e6327318"
        #[string]$scope = "Site"  # Standardmäßig "Site"
    )
    
    # Document ID Service auf der Site starten
    $featureId = "b50e3104-6812-424f-a011-cc90e6327318"
    Log "Enable-PnPFeature -Identity $featureId -Scope Site <!-- Document ID Service -->"
    Enable-PnPFeature -Identity $featureId -Scope Site -Force

    # Document Set Feature Service auf der Site starten
    $featureId = "3bae86a2-776d-499d-9db8-fa4cdc7884f8"
    Log "Prüfe auf Content Type: 'Document Set'  (ID: 0x0120D520)..."
    $docSetCT = Get-PnPContentType | Where-Object { $_.Id -like "0x0120D520*" }
    if ($docSetCT) {
        Log "ℹ️ Content Type: 'Document Set' already exists, skipping feature activation"
        #break
    } else {
        Log "Enable-PnPFeature -Identity $featureId -Scope Site <!-- Document Set -->"
        Enable-PnPFeature -Identity $featureId -Scope Site -Force
    }

    # Warten und prüfen, ob das Feature angekommen ist (max. 300 Sekunden)
    $maxRetries  = 10 # z.B. 10 Versuche
    $secWait     = 10 # z.B. alle 10 Sekunden
    $retryCount  = 0
    $docSetCT    = $null

    do {
        $docSetCT = Get-PnPContentType | Where-Object { $_.Id -like "0x0120D520*" }
        if ($docSetCT) {
            break
        }
        $retryCount++
        Log "⏳ Warte auf Content Type: 'Document Set'... ($retryCount/$maxRetries)"
        Start-Sleep -Seconds $secWait
    } while ($retryCount -lt $maxRetries)

    if (-not $docSetCT) {
        throw "❌ Content Type: 'Document Set' nicht gefunden! Prüfe ob das Feature korrekt aktiviert ist."
    } else {
        $docSetName = $docSetCT.Name
        $docSetID = $docSetCT.Id
        Log "✅ Content Type: 'Document Set' erfolgreich aktiviert: $docSetName / $docSetID"
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