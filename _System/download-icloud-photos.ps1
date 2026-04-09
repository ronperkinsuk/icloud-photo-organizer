<#
.SYNOPSIS
    iCloud Photos Organizer for Windows

.DESCRIPTION
    This script processes iCloud Photos downloaded via iCloud for Windows across multiple user accounts,
    extracts metadata using ExifTool, and organises media into a structured folder hierarchy.

    Files are grouped by:
        Person → Year → Device Model or WhatsApp

    Script updates:
        - Triggers iCloud stub hydration via a lightweight 1-byte FileStream read to avoid Windows dialogs
        - Uses [System.IO.File]::Copy() instead of Copy-Item to suppress shell dialog prompts
        - Uses a temporary working directory to avoid iCloud file locking issues
        - Extracts EXIF metadata (date taken, device model)
        - Builds a date-to-model map from photos to infer camera model for videos with no EXIF
        - Routes MP4 files to WhatsApp folder, MOV/M4V without model to Unknown Model folder
        - Detects and skips duplicate files per user using SHA256 hashes
        - Separates WhatsApp / non-EXIF media into appropriate folders
        - Outputs photos and videos to separate destination root folders
        - Appends a timestamp suffix to destination filename if a name collision occurs
        - Logs processed files per person to allow re-running individual people from scratch
        - Logs errors, slow downloads and run history to separate log files
        - Tracks progress, ETA, elapsed time and per-user statistics
        - Reverts processed files back to iCloud stub state to free disk space
        - Separates critical state files from disposable run logs
        - Configurable download timeouts and source paths

.NOTES
    Author: Ron Perkins
    Version: 2.3.0
    Created: 2026-03-31
    Updated: 2026-04-09

.REQUIREMENTS
    - Windows 10/11
    - iCloud for Windows (Photos enabled)
    - ExifTool (https://exiftool.org/)
    - PowerShell 5.1+

.LICENSE
    MIT License
    Copyright (c) 2026 Ron Perkins
#>

<#
 Configuration
#>
. "$PSScriptRoot\config.ps1"
$runLog = "$logsDir\run_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').log"

<#
 Startup Checks
#>
if (!(Test-Path $exiftool)) {
    Write-Host "ERROR: exiftool not found at $exiftool"
    exit 1
}
foreach ($src in $sources) {
    if (!(Test-Path $src.Path)) {
        Write-Host "ERROR: Source path not found for $($src.Person): $($src.Path)"
        exit 1
    }
}

<#
 Startup Cleanup
#>
if (!(Test-Path $sysDir))  { New-Item -ItemType Directory -Force -Path $sysDir  | Out-Null }
if (!(Test-Path $logsDir)) { New-Item -ItemType Directory -Force -Path $logsDir | Out-Null }
if (!(Test-Path $dataDir)) { New-Item -ItemType Directory -Force -Path $dataDir | Out-Null }
if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force }
New-Item -ItemType Directory -Force -Path $tempDir | Out-Null

<#
 Initialize Logs
#>
function Log {
    param([string]$msg)
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $msg"
    Write-Host $msg
    Add-Content $runLog $line
}

<#
 Load per-person processed and hash data
#>
$processed = @{}
$hashDB = @{}
foreach ($src in $sources) {
    $person  = $src.Person
    $logFile = Get-PersonLogFile $person
    $hashLog = Get-PersonHashFile $person

    if (!(Test-Path $hashLog)) { "Person,Hash" | Out-File $hashLog }
    if (Test-Path $logFile) { Get-Content $logFile | ForEach-Object { $processed[$_] = $true } }
    if (Test-Path $hashLog) {
        Import-Csv $hashLog | ForEach-Object { $hashDB["$($_.Person)|$($_.Hash)"] = $true }
    }
}

$stats = @{}
$shell = New-Object -ComObject Shell.Application

<#
 Count Total Files for Progress
#>
$totalFiles = 0
foreach ($src in $sources) {
    $totalFiles += (Get-ChildItem -Path $src.Path -Recurse -File | Where-Object {
        $ext = $_.Extension.ToLower()
        ($photoExt -contains $ext) -or ($videoExt -contains $ext)
    }).Count
}
$totalFiles -= $processed.Count
$currentFile = 0
$startTime = Get-Date
$dateModelMap = @{}

<#
 Process Users
#>
foreach ($src in $sources) {
    $sourcePath = $src.Path
    $person     = $src.Person
    $logFile    = Get-PersonLogFile $person
    $hashLog    = Get-PersonHashFile $person

    $stats[$person] = @{ Photos = 0; Videos = 0; WhatsApp = 0; Skipped = 0; Errors = 0 }

    Log ""
    Log "===== Processing $person ====="

    $personFile = 0
    $files = Get-ChildItem -Path $sourcePath -Recurse -File

    foreach ($file in $files) {
        $filePath = $file.FullName
        if ($processed.ContainsKey($filePath)) { continue }

        $ext = $file.Extension.ToLower()
        if (($photoExt -notcontains $ext) -and ($videoExt -notcontains $ext)) { continue }

        $currentFile++
        $personFile++
        #if ($personFile -gt 5) { $currentFile--; break }
        $elapsed = (Get-Date) - $startTime
        $avgSec = if ($elapsed.TotalSeconds -gt 0) { $elapsed.TotalSeconds / $currentFile } else { 0 }
        $remaining = [int]($avgSec * ($totalFiles - $currentFile))
        $elapsedStr = "{0:hh\:mm\:ss}" -f [timespan]::FromSeconds($elapsed.TotalSeconds)
        $etaStr = "{0:hh\:mm\:ss}" -f [timespan]::FromSeconds($remaining)
        $remainingFiles = $totalFiles - $currentFile
        Log "[$currentFile | $remainingFiles] Time: $elapsedStr | ETA: $etaStr | Processing: $person - $($file.Name)"

        $tempFile = Join-Path $tempDir ([System.IO.Path]::GetFileName($filePath))

        try {
            <#
             Wait for iCloud Download
            #>
            $maxWaitSec = if ($videoExt -contains $ext) { $maxWaitSecVideo } else { $maxWaitSecPhoto }
            $waited = 0

            while ($true) {
                $origFile = Get-Item $filePath

                if (-not ($origFile.Attributes -band [System.IO.FileAttributes]::Offline)) {
                    break
                }

                # Trigger hydration (light touch)
                try {
                    $fs = [System.IO.File]::Open($filePath, 'Open', 'Read', 'ReadWrite')
                    $buffer = New-Object byte[] 1
                    try { $fs.Read($buffer, 0, 1) | Out-Null } finally { $fs.Close() }
                } catch {}

                Start-Sleep -Seconds 2
                $waited += 2

                if ($waited -ge $maxWaitSec) {
                    Add-Content $slowLog "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $filePath"
                    throw "iCloud did not download file after $maxWaitSec seconds"
                }
            }

            <#
             Copy to Temp
            #>
            #Copy-Item $filePath -Destination $tempFile -Force -ErrorAction Stop
            try {
                [System.IO.File]::Copy($filePath, $tempFile, $true)
            } catch {
                throw "Temp copy failed: $($_.Exception.Message)"
            }
            if (!(Test-Path $tempFile)) { throw "Temp copy failed" }

            <#
             Read Metadata
            #>
            $output = & $exiftool -DateTimeOriginal -CreateDate -Model -s3 $tempFile
            $lines = $output -split "`n"

            $dateStr = if ($lines.Count -gt 0 -and $lines[0]) { $lines[0].Trim() } else { "" }
            if (-not $dateStr) { $dateStr = if ($lines.Count -gt 1 -and $lines[1]) { $lines[1].Trim() } else { "" } }
            if (-not $dateStr) { $dateTaken = (Get-Item $tempFile).LastWriteTime }
            else {
                try { $dateTaken = [datetime]::ParseExact($dateStr.Trim(), "yyyy:MM:dd HH:mm:ss", $null) }
                catch { $dateTaken = (Get-Item $tempFile).LastWriteTime }
            }
            $year = $dateTaken.Year

            <#
             Fetch Type
            #>
            $typeFolder = if ($videoExt -contains $ext) { "Videos" } else { "Photos" }

            <#
             Update Date to Model Map
            #>
            $modelStr = if ($lines.Count -gt 2 -and $lines[2]) { $lines[2].Trim() } else { "" }
            $mapKey = "$person|$($dateTaken.ToString('yyyy-MM-dd'))"
            if ($typeFolder -eq "Photos" -and $modelStr) {
                if (-not $dateModelMap.ContainsKey($mapKey)) {
                    $dateModelMap[$mapKey] = $modelStr
                }
            }

            <#
             Fetch Model / WhatsApp
            #>
            if (-not $modelStr -and $typeFolder -eq "Videos" -and $ext -ne ".mp4") {
                if ($dateModelMap.ContainsKey($mapKey)) {
                    $modelStr = $dateModelMap[$mapKey]
                }
            }
            $subFolder = if ($ext -eq ".mp4") {
                "WhatsApp"
            } elseif (-not $modelStr) {
                if ($typeFolder -eq "Videos") { "Unknown Model" } else { "WhatsApp" }
            } else {
                $modelStr
            }
            $subFolder = $subFolder -replace '[\\/:*?"<>|]', '_'

            <#
             Hash (detect duplicate)
            #>
            $hash = (Get-FileHash $tempFile -Algorithm SHA256).Hash
            $hashKey = "$person|$hash"
            if ($hashDB.ContainsKey($hashKey)) {
                Log "  Skipped (duplicate)"
                $stats[$person].Skipped++
                $processed[$filePath] = $true
                Add-Content $logFile $filePath
                Remove-Item $tempFile -Force
                continue
            }

            <#
             Destination
            #>
            if ($typeFolder -eq "Videos") {
                $targetFolder = "$destVideo\$person\$year\$subFolder"
            } else {
                $targetFolder = "$destPhoto\$person\$year\$subFolder"
            }

            New-Item -ItemType Directory -Force -Path $targetFolder | Out-Null

            $destFile = Join-Path $targetFolder $file.Name
            if (Test-Path $destFile) {
                $destFile = Join-Path $targetFolder (
                    "{0}_{1}{2}" -f $file.BaseName, (Get-Date -Format "HHmmssffff"), $file.Extension
                )
            }

            Move-Item $tempFile -Destination $destFile -Force

            Add-Content $logFile $filePath
            Add-Content $hashLog "$person,$hash"
            $hashDB[$hashKey] = $true
            $processed[$filePath] = $true

            <#
             Evict back to iCloud stub
            #>
            try {
                $shellFile = $shell.Namespace($file.DirectoryName).ParseName($file.Name)
                if ($shellFile) {
                    $verb = $shellFile.Verbs() | Where-Object { $_.Name -eq "Free up space" }
                    if ($verb) { $verb.DoIt() }
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shellFile) | Out-Null
                    $shellFile = $null
                }
            } catch {
                Log "  Warning: Eviction failed for $($file.Name) - $_"
            }

            <#
             Stats
            #>
            if ($typeFolder -eq "Videos") { $stats[$person].Videos++ }
            elseif ($subFolder -eq "WhatsApp") { $stats[$person].WhatsApp++ }
            else { $stats[$person].Photos++ }

            Log "  Done -> $targetFolder"

            <#
             Exit after file limit so launcher can restart in a fresh process
            #>
            if ($currentFile -ge $processFileLimit) {
                Log ""
                Log "Processed $currentFile files - exiting for restart..."
                Log ""
                exit
            }
        }
        catch {
            $errMsg = "  Error processing file: $filePath - $_"
            Log $errMsg
            Add-Content $errorLog "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $errMsg"
            $stats[$person].Errors++
            if (Test-Path $tempFile) { Remove-Item $tempFile -Force }
        }
    }
}

<#
 Summary
#>
Log ""
Log "===== SUMMARY ====="
foreach ($person in $stats.Keys) {
    Log ""
    Log "${person}:"
    Log "  Photos:   $($stats[$person].Photos)"
    Log "  Videos:   $($stats[$person].Videos)"
    Log "  WhatsApp: $($stats[$person].WhatsApp)"
    Log "  Skipped:  $($stats[$person].Skipped)"
    Log "  Errors:   $($stats[$person].Errors)"
}
$elapsed = (Get-Date) - $startTime
Log ""
Log "Total processed this run: $currentFile"
Log "Elapsed time: $($elapsed.ToString('hh\:mm\:ss'))"
Log ""
Log "All users processed."

<#
 Cleanup
#>
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) | Out-Null
if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force }
