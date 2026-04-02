<#
.SYNOPSIS
    iCloud Photos Processing Launcher for Windows

.DESCRIPTION
    This script acts as a resilient launcher for the main iCloud photo processing script
    (download-icloud-photos.ps1), repeatedly executing it in a fresh PowerShell process
    until all source media files have been processed.

    The launcher is designed to improve long-running stability by avoiding issues such as
    CHARTV.dll crashes and memory buildup that can occur during extended PowerShell sessions.

    Behaviour:
        - Starts the main processing script in a new PowerShell process
        - Waits for the script to complete (each run processes up to 2000 files)
        - Recounts total source files across all configured user paths
        - Compares against processed.log to determine progress
        - Repeats execution until all files are processed
        - Exits cleanly once processing is complete

    Key features:
        - Enables batch-based processing for improved reliability
        - Provides automatic restart and continuation between runs
        - Uses persistent processed.log to track overall progress
        - Prevents reprocessing of already handled files
        - Simple progress output between runs

.NOTES
    Author: Ron Perkins
    Version: 1.0.0
    Created: 2026-04-02

.REQUIREMENTS
    - Windows 10/11
    - PowerShell 5.1+
    - download-icloud-photos.ps1 (main processor script)
    - config.ps1 (shared configuration)

.LICENSE
    MIT License
    Copyright (c) 2026 Ron Perkins
#>

. "$PSScriptRoot\config.ps1"

$script = "$sysDir\download-icloud-photos.ps1"

while ($true) {
    Write-Host "Starting new run..."
    Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$script`"" -Wait
    Write-Host "Run complete. Checking if more files to process..."

    $totalFiles = 0
    foreach ($src in $sources) {
        $totalFiles += (Get-ChildItem -Path $src.Path -Recurse -File | Where-Object {
            $ext = $_.Extension.ToLower()
            ($photoExt -contains $ext) -or ($videoExt -contains $ext)
        }).Count
    }

    $processedCount = 0
    if (Test-Path $logFile) { $processedCount = (Get-Content $logFile).Count }

    Write-Host "Processed: $processedCount / $totalFiles"

    if ($processedCount -ge $totalFiles) {
        Write-Host "All files processed. Done!"
        break
    }
}
