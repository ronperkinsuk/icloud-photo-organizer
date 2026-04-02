<#
.SYNOPSIS
    iCloud Photos Organizer for Windows with resilient processing and deduplication

.DESCRIPTION
    This script processes iCloud Photos downloaded via iCloud for Windows across multiple user accounts,
    extracting metadata using ExifTool and organising media into a structured, de-duplicated folder hierarchy.

    Files are organised as:
        Person → Media Type (Photos/Videos) → Year → Device Model / WhatsApp / Unknown

    Key features:
        - Forces download of iCloud "on-demand" (stub) files with configurable timeouts
        - Uses a temporary working directory to avoid iCloud locking and partial file issues
        - Extracts EXIF metadata (DateTimeOriginal, CreateDate, device model) using ExifTool
        - Builds a per-user date-to-model map from photos to infer missing video metadata
        - Automatically classifies:
            * MP4 → WhatsApp
            * Non-EXIF media → WhatsApp (photos) or Unknown Model (videos)
        - Performs SHA256 hashing to detect and skip duplicate files per user
        - Maintains persistent state:
            * processed.log → tracks processed source files (safe re-runs)
            * hashes.csv   → tracks deduplicated content
        - Supports crash-safe restarts without reprocessing completed files
        - Logs:
            * Run activity (per execution)
            * Errors
            * Slow iCloud downloads
        - Tracks progress with elapsed time, ETA, and per-user statistics
        - Reverts processed files back to iCloud "Free up space" (stub state) to minimise disk usage
        - Separates persistent state data from disposable run logs
        - Automatically exits after a batch (2000 files) to allow clean restarts and memory stability

.NOTES
    Author: Ron Perkins
    Version: 2.0.0
    Created: 2026-03-31
    Updated: 2026-04-02

.REQUIREMENTS
    - Windows 10/11
    - iCloud for Windows (Photos enabled with "Optimise Storage")
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

if (!(Test-Path $hashLog)) { "Person,Hash" | Out-File $hashLog }
$processed = @{}
if (Test-Path $logFile) { Get-Content $logFile | ForEach-Object { $processed[$_] = $true } }

$hashDB = @{}
if (Test-Path $hashLog) {
    Import-Csv $hashLog | ForEach-Object { $hashDB["$($_.Person)|$($_.Hash)"] = $true }
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
$currentFile = 0
$startTime = Get-Date
$dateModelMap = @{}

<#
 Process Users
#>
foreach ($src in $sources) {
    $sourcePath = $src.Path
    $person     = $src.Person

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
        $remainingFiles = $totalFiles - $processed.Count - 1
        Log "[$currentFile | $remainingFiles] Time: $elapsedStr | ETA: $etaStr | Processing: $person - $($file.Name)"

        $tempFile = Join-Path $tempDir ([System.IO.Path]::GetFileName($filePath))

        try {
            <#
             Wait for iCloud Download
            #>
            $maxWaitSec = if ($videoExt -contains $ext) { $maxWaitSecVideo } else { $maxWaitSecPhoto }
            $waited = 0
            try { $null = [System.IO.File]::ReadAllBytes($filePath) } catch { }
            $origFile = Get-Item $filePath
            while ($origFile.Attributes -band [System.IO.FileAttributes]::Offline) {
                if ($waited -ge $maxWaitSec) {
                    Add-Content $slowLog "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $filePath"
                    throw "iCloud did not download file after $maxWaitSec seconds"
                }
                Start-Sleep -Seconds 2
                $waited += 2
                $origFile = Get-Item $filePath
            }

            <#
             Copy to Temp
            #>
            Copy-Item $filePath -Destination $tempFile -Force -ErrorAction Stop
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
            $targetFolder = "$dest\$person\$typeFolder\$year\$subFolder"
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
             Exit after 2000 files so launcher can restart in a fresh process
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
