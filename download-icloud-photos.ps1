<#
.SYNOPSIS
    iCloud Photos Organizer for Windows

.DESCRIPTION
    This script processes iCloud Photos downloaded via iCloud for Windows across multiple user accounts,
    extracts metadata using ExifTool, and organises media into a structured folder hierarchy.

    Files are grouped by:
        Person → Media Type (Photos/Videos) → Year → Device Model or WhatsApp

    The script:
        - Forces download of iCloud "on-demand" files
        - Uses a temporary working directory to avoid iCloud file locking issues
        - Extracts EXIF metadata (date taken, device model)
        - Detects and skips duplicate files per user using SHA256 hashes
        - Separates WhatsApp / non-EXIF media
        - Logs processed files to allow safe re-runs
        - Tracks progress, ETA, and per-user statistics
        - Reverts files back to iCloud "cloud-only" state to save disk space

.NOTES
    Author: Ron Perkins
    Version: 1.0.0
    Created: 2026-03-31

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
$sources = @(
    @{ Path = "C:\Users\User1\Pictures\iCloud Photos\Photos";  Person = "User1" },
    @{ Path = "C:\Users\User2\Pictures\iCloud Photos\Photos"; Person = "User2" },
    @{ Path = "C:\Users\User3\Pictures\iCloud Photos\Photos";  Person = "User3" }
)

$dest     = "D:\Photos\Phone Media"
$tempDir  = "D:\Photos\Phone Media\Temp"
$logsDir  = "D:\Photos\Phone Media\Logs"
$logFile  = "$logsDir\processed.log"
$hashLog  = "$logsDir\hashes.csv"
$errorLog = "$logsDir\errors.log"
$runLog   = "$logsDir\run_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').log"
$exiftool = "C:\Tools\Exiftool\exiftool.exe"

$photoExt = @(".jpg",".jpeg",".heic",".png")
$videoExt = @(".mov",".mp4",".m4v")

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
if (!(Test-Path $logsDir)) { New-Item -ItemType Directory -Force -Path $logsDir | Out-Null }
if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force }
New-Item -ItemType Directory -Force -Path $tempDir | Out-Null

<#
 Initalize Logs
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

$shell = New-Object -ComObject Shell.Application

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
        if ($personFile -gt 1) { $currentFile--; break }
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
            $maxWaitSec = 120
            $waited = 0
            try { $null = [System.IO.File]::ReadAllBytes($filePath) } catch { }
            $origFile = Get-Item $filePath
            while ($origFile.Attributes -band [System.IO.FileAttributes]::Offline) {
                if ($waited -ge $maxWaitSec) { throw "iCloud did not download file after $maxWaitSec seconds" }
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
             Fetch Model / WhatsApp
            #>
            $modelStr = if ($lines.Count -gt 2 -and $lines[2]) { $lines[2].Trim() } else { "" }
            $subFolder = if (-not $modelStr) { "WhatsApp" } else { $modelStr }
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
             Evict back to iCloud Stub
            #>
            try {
                if ($currentFile % 50 -eq 0) {
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) | Out-Null
                    $shell = $null
                    [System.GC]::Collect()
                    [System.GC]::WaitForPendingFinalizers()
                    [System.GC]::Collect()
                    $shell = New-Object -ComObject Shell.Application
                }
                $shellFile = $shell.Namespace($file.DirectoryName).ParseName($file.Name)
                $verb = $shellFile.Verbs() | Where-Object { $_.Name -eq "Free up space" }
                if ($verb) { $verb.DoIt() }
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shellFile) | Out-Null
                $shellFile = $null
            } catch {
                Log "  Warning: Could not evict file back to stub: $($file.Name) - $_"
            }

            <#
             Stats
            #>
            if ($typeFolder -eq "Videos") { $stats[$person].Videos++ }
            elseif ($subFolder -eq "WhatsApp") { $stats[$person].WhatsApp++ }
            else { $stats[$person].Photos++ }

            Log "  Done -> $targetFolder"

            # Force garbage collection every 50 files
            if ($currentFile % 50 -eq 0) {
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
                [System.GC]::Collect()
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
    Log "  Photos:  $($stats[$person].Photos)"
    Log "  Videos:  $($stats[$person].Videos)"
    Log "  WhatsApp: $($stats[$person].WhatsApp)"
    Log "  Skipped: $($stats[$person].Skipped)"
    Log "  Errors:  $($stats[$person].Errors)"
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
if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force }
