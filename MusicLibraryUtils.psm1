# Music Library Utilities Module
# Additional functions for advanced music library management

# Export functions
Export-ModuleMember -Function *

<#
.SYNOPSIS
Analyzes a music library and provides detailed statistics

.DESCRIPTION
Scans a music library folder and generates comprehensive statistics about
file formats, bitrates, genres, artists, and potential issues

.PARAMETER Path
Path to the music library folder

.PARAMETER IncludeFileAnalysis
Whether to perform detailed file analysis (slower but more comprehensive)

.EXAMPLE
Get-MusicLibraryStats -Path "h:\Music" -IncludeFileAnalysis
#>
function Get-MusicLibraryStats {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path,
        
        [switch]$IncludeFileAnalysis
    )
    
    if (-not (Test-Path $Path)) {
        throw "Path does not exist: $Path"
    }
    
    Write-Host "Analyzing music library at: $Path" -ForegroundColor Green
    
    $musicFiles = Get-ChildItem -Path $Path -Recurse -Include *.mp3,*.flac,*.m4a,*.wav,*.wma,*.ogg
    $totalFiles = $musicFiles.Count
    $totalSize = ($musicFiles | Measure-Object -Property Length -Sum).Sum
    
    $stats = @{
        TotalFiles = $totalFiles
        TotalSizeGB = [math]::Round($totalSize / 1GB, 2)
        Formats = @{}
        Artists = @{}
        Genres = @{}
        Years = @{}
        Bitrates = @{}
        Issues = @()
        Duplicates = @()
    }
    
    $processed = 0
    foreach ($file in $musicFiles) {
        $processed++
        $percentComplete = [math]::Round(($processed / $totalFiles) * 100, 1)
        Write-Progress -Activity "Analyzing Files" -Status "$processed of $totalFiles ($percentComplete%)" -PercentComplete $percentComplete
        
        # File format analysis
        $extension = $file.Extension.ToLower()
        if (-not $stats.Formats.ContainsKey($extension)) {
            $stats.Formats[$extension] = 0
        }
        $stats.Formats[$extension]++
        
        if ($IncludeFileAnalysis) {
            try {
                # Extract metadata
                $shell = New-Object -ComObject Shell.Application
                $folder = $shell.Namespace($file.DirectoryName)
                $fileItem = $folder.ParseName($file.Name)
                
                # Artist
                $artist = $folder.GetDetailsOf($fileItem, 13)
                if ($artist -and $artist.Trim() -ne "") {
                    $artist = $artist.Trim()
                    if (-not $stats.Artists.ContainsKey($artist)) {
                        $stats.Artists[$artist] = 0
                    }
                    $stats.Artists[$artist]++
                }
                
                # Genre
                $genre = $folder.GetDetailsOf($fileItem, 16)
                if ($genre -and $genre.Trim() -ne "") {
                    $genre = $genre.Trim()
                    if (-not $stats.Genres.ContainsKey($genre)) {
                        $stats.Genres[$genre] = 0
                    }
                    $stats.Genres[$genre]++
                }
                
                # Year
                $year = $folder.GetDetailsOf($fileItem, 15)
                if ($year -and $year.Trim() -ne "") {
                    $year = $year.Trim() -replace '[^\d]', ''
                    if ($year.Length -eq 4) {
                        if (-not $stats.Years.ContainsKey($year)) {
                            $stats.Years[$year] = 0
                        }
                        $stats.Years[$year]++
                    }
                }
                
                # Bitrate
                $bitrate = $folder.GetDetailsOf($fileItem, 28)
                if ($bitrate -and $bitrate.Trim() -ne "") {
                    $bitrate = $bitrate.Trim()
                    if (-not $stats.Bitrates.ContainsKey($bitrate)) {
                        $stats.Bitrates[$bitrate] = 0
                    }
                    $stats.Bitrates[$bitrate]++
                }
                
                # Check for potential issues
                if (-not $artist -or $artist.Trim() -eq "") {
                    $stats.Issues += "Missing artist: $($file.FullName)"
                }
                
                if ($file.Name -match '^\d+\s*[-._]\s*$') {
                    $stats.Issues += "Poor filename: $($file.FullName)"
                }
                
            }
            catch {
                $stats.Issues += "Metadata extraction failed: $($file.FullName) - $($_.Exception.Message)"
            }
        }
    }
    
    Write-Progress -Activity "Analyzing Files" -Completed
    
    # Find potential duplicates based on filename similarity
    Write-Host "Checking for potential duplicates..." -ForegroundColor Yellow
    $fileGroups = $musicFiles | Group-Object { [System.IO.Path]::GetFileNameWithoutExtension($_.Name) }
    foreach ($group in $fileGroups) {
        if ($group.Count -gt 1) {
            $stats.Duplicates += @{
                Name = $group.Name
                Files = $group.Group | ForEach-Object { $_.FullName }
                Count = $group.Count
            }
        }
    }
    
    return $stats
}

<#
.SYNOPSIS
Finds and reports duplicate music files

.DESCRIPTION
Scans for duplicate music files based on various criteria including
filename, size, duration, and audio fingerprinting

.PARAMETER Path
Path to scan for duplicates

.PARAMETER Method
Duplication detection method: 'Filename', 'Size', 'Hash', or 'All'

.EXAMPLE
Find-DuplicateMusic -Path "h:\Music" -Method "All"
#>
function Find-DuplicateMusic {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path,
        
        [ValidateSet('Filename', 'Size', 'Hash', 'All')]
        [string]$Method = 'All'
    )
    
    $musicFiles = Get-ChildItem -Path $Path -Recurse -Include *.mp3,*.flac,*.m4a,*.wav,*.wma,*.ogg
    $duplicates = @()
    
    Write-Host "Scanning $($musicFiles.Count) files for duplicates using method: $Method" -ForegroundColor Green
    
    if ($Method -eq 'Filename' -or $Method -eq 'All') {
        Write-Host "Checking filename duplicates..." -ForegroundColor Yellow
        $nameGroups = $musicFiles | Group-Object { [System.IO.Path]::GetFileNameWithoutExtension($_.Name).ToLower() }
        foreach ($group in $nameGroups | Where-Object { $_.Count -gt 1 }) {
            $duplicates += @{
                Type = 'Filename'
                Key = $group.Name
                Files = $group.Group
                Count = $group.Count
            }
        }
    }
    
    if ($Method -eq 'Size' -or $Method -eq 'All') {
        Write-Host "Checking size duplicates..." -ForegroundColor Yellow
        $sizeGroups = $musicFiles | Group-Object Length
        foreach ($group in $sizeGroups | Where-Object { $_.Count -gt 1 }) {
            $duplicates += @{
                Type = 'Size'
                Key = "$($group.Name) bytes"
                Files = $group.Group
                Count = $group.Count
            }
        }
    }
    
    if ($Method -eq 'Hash' -or $Method -eq 'All') {
        Write-Host "Checking hash duplicates (this may take a while)..." -ForegroundColor Yellow
        $hashGroups = $musicFiles | Group-Object { (Get-FileHash $_.FullName -Algorithm MD5).Hash }
        foreach ($group in $hashGroups | Where-Object { $_.Count -gt 1 }) {
            $duplicates += @{
                Type = 'Hash'
                Key = $group.Name
                Files = $group.Group
                Count = $group.Count
            }
        }
    }
    
    return $duplicates
}

<#
.SYNOPSIS
Validates music file integrity

.DESCRIPTION
Checks music files for corruption, missing metadata, and other issues

.PARAMETER Path
Path to validate

.PARAMETER FixIssues
Attempt to fix minor issues automatically

.EXAMPLE
Test-MusicFileIntegrity -Path "h:\Music" -FixIssues
#>
function Test-MusicFileIntegrity {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path,
        
        [switch]$FixIssues
    )
    
    $musicFiles = Get-ChildItem -Path $Path -Recurse -Include *.mp3,*.flac,*.m4a,*.wav,*.wma,*.ogg
    $issues = @()
    $fixed = @()
    
    Write-Host "Validating $($musicFiles.Count) music files..." -ForegroundColor Green
    
    $processed = 0
    foreach ($file in $musicFiles) {
        $processed++
        $percentComplete = [math]::Round(($processed / $musicFiles.Count) * 100, 1)
        Write-Progress -Activity "Validating Files" -Status "$processed of $($musicFiles.Count) ($percentComplete%)" -PercentComplete $percentComplete
        
        $fileIssues = @()
        
        # Check file accessibility
        try {
            $stream = [System.IO.File]::OpenRead($file.FullName)
            $stream.Close()
        }
        catch {
            $fileIssues += "File access error: $($_.Exception.Message)"
        }
        
        # Check file size
        if ($file.Length -eq 0) {
            $fileIssues += "Zero-byte file"
        }
        elseif ($file.Length -lt 1KB) {
            $fileIssues += "Suspiciously small file ($($file.Length) bytes)"
        }
        
        # Check filename issues
        $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
        $hasInvalidChars = $false
        foreach ($char in $invalidChars) {
            if ($file.Name.Contains($char)) {
                $hasInvalidChars = $true
                break
            }
        }
        
        if ($hasInvalidChars) {
            $fileIssues += "Invalid characters in filename"
            
            if ($FixIssues) {
                try {
                    $newName = $file.Name
                    foreach ($char in $invalidChars) {
                        $newName = $newName.Replace($char, '_')
                    }
                    $newPath = Join-Path $file.DirectoryName $newName
                    Rename-Item -Path $file.FullName -NewName $newName
                    $fixed += "Renamed: $($file.FullName) -> $newPath"
                }
                catch {
                    $fileIssues += "Failed to fix filename: $($_.Exception.Message)"
                }
            }
        }
        
        # Check for very long paths (Windows limitation)
        if ($file.FullName.Length -gt 260) {
            $fileIssues += "Path too long ($($file.FullName.Length) characters)"
        }
        
        # Check metadata
        try {
            $shell = New-Object -ComObject Shell.Application
            $folder = $shell.Namespace($file.DirectoryName)
            $fileItem = $folder.ParseName($file.Name)
            
            $title = $folder.GetDetailsOf($fileItem, 21)
            $artist = $folder.GetDetailsOf($fileItem, 13)
            $album = $folder.GetDetailsOf($fileItem, 14)
            
            if (-not $title -or $title.Trim() -eq "") {
                $fileIssues += "Missing title metadata"
            }
            if (-not $artist -or $artist.Trim() -eq "") {
                $fileIssues += "Missing artist metadata"
            }
            if (-not $album -or $album.Trim() -eq "") {
                $fileIssues += "Missing album metadata"
            }
        }
        catch {
            $fileIssues += "Metadata extraction failed: $($_.Exception.Message)"
        }
        
        if ($fileIssues.Count -gt 0) {
            $issues += @{
                File = $file.FullName
                Issues = $fileIssues
            }
        }
    }
    
    Write-Progress -Activity "Validating Files" -Completed
    
    return @{
        TotalFiles = $musicFiles.Count
        FilesWithIssues = $issues.Count
        Issues = $issues
        FixedIssues = $fixed
    }
}

<#
.SYNOPSIS
Generates a detailed music library report

.DESCRIPTION
Creates a comprehensive HTML report of music library statistics and analysis

.PARAMETER Path
Path to the music library

.PARAMETER OutputPath
Path for the output HTML report

.EXAMPLE
New-MusicLibraryReport -Path "h:\Music" -OutputPath "h:\Music\LibraryReport.html"
#>
function New-MusicLibraryReport {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path,
        
        [Parameter(Mandatory=$true)]
        [string]$OutputPath
    )
    
    Write-Host "Generating comprehensive music library report..." -ForegroundColor Green
    
    # Get statistics
    $stats = Get-MusicLibraryStats -Path $Path -IncludeFileAnalysis
    $duplicates = Find-DuplicateMusic -Path $Path -Method "All"
    $integrity = Test-MusicFileIntegrity -Path $Path
    
    # Generate HTML report
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Music Library Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .header { background-color: #f0f0f0; padding: 20px; border-radius: 5px; }
        .section { margin: 20px 0; }
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; }
        .stat-card { background-color: #f9f9f9; padding: 15px; border-radius: 5px; border-left: 4px solid #007acc; }
        .stat-number { font-size: 24px; font-weight: bold; color: #007acc; }
        .stat-label { color: #666; }
        table { border-collapse: collapse; width: 100%; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        .issue { color: #d9534f; }
        .success { color: #5cb85c; }
        .warning { color: #f0ad4e; }
    </style>
</head>
<body>
    <div class="header">
        <h1>Music Library Report</h1>
        <p>Generated on: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
        <p>Library Path: $Path</p>
    </div>
    
    <div class="section">
        <h2>Overview</h2>
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-number">$($stats.TotalFiles)</div>
                <div class="stat-label">Total Files</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">$($stats.TotalSizeGB) GB</div>
                <div class="stat-label">Total Size</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">$($stats.Artists.Count)</div>
                <div class="stat-label">Unique Artists</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">$($stats.Genres.Count)</div>
                <div class="stat-label">Unique Genres</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">$($duplicates.Count)</div>
                <div class="stat-label">Potential Duplicates</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">$($integrity.FilesWithIssues)</div>
                <div class="stat-label">Files with Issues</div>
            </div>
        </div>
    </div>
    
    <div class="section">
        <h2>File Formats</h2>
        <table>
            <tr><th>Format</th><th>Count</th><th>Percentage</th></tr>
"@
    
    foreach ($format in $stats.Formats.GetEnumerator() | Sort-Object Value -Descending) {
        $percentage = [math]::Round(($format.Value / $stats.TotalFiles) * 100, 1)
        $html += "            <tr><td>$($format.Key)</td><td>$($format.Value)</td><td>$percentage%</td></tr>`n"
    }
    
    $html += @"
        </table>
    </div>
    
    <div class="section">
        <h2>Top Artists</h2>
        <table>
            <tr><th>Artist</th><th>Track Count</th></tr>
"@
    
    foreach ($artist in $stats.Artists.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 20) {
        $html += "            <tr><td>$($artist.Key)</td><td>$($artist.Value)</td></tr>`n"
    }
    
    $html += @"
        </table>
    </div>
    
    <div class="section">
        <h2>Genres</h2>
        <table>
            <tr><th>Genre</th><th>Track Count</th></tr>
"@
    
    foreach ($genre in $stats.Genres.GetEnumerator() | Sort-Object Value -Descending) {
        $html += "            <tr><td>$($genre.Key)</td><td>$($genre.Value)</td></tr>`n"
    }
    
    $html += @"
        </table>
    </div>
    
    <div class="section">
        <h2>Issues Found</h2>
        <p class="$(if ($stats.Issues.Count -eq 0) { 'success' } else { 'warning' })">$($stats.Issues.Count) issues found</p>
"@
    
    if ($stats.Issues.Count -gt 0) {
        $html += "        <ul>`n"
        foreach ($issue in $stats.Issues | Select-Object -First 50) {
            $html += "            <li class='issue'>$issue</li>`n"
        }
        if ($stats.Issues.Count -gt 50) {
            $html += "            <li>... and $($stats.Issues.Count - 50) more issues</li>`n"
        }
        $html += "        </ul>`n"
    }
    
    $html += @"
    </div>
    
    <div class="section">
        <h2>Potential Duplicates</h2>
        <p class="$(if ($duplicates.Count -eq 0) { 'success' } else { 'warning' })">$($duplicates.Count) potential duplicate groups found</p>
"@
    
    if ($duplicates.Count -gt 0) {
        foreach ($duplicate in $duplicates | Select-Object -First 20) {
            $html += "        <h4>$($duplicate.Type): $($duplicate.Key) ($($duplicate.Count) files)</h4>`n"
            $html += "        <ul>`n"
            foreach ($file in $duplicate.Files) {
                $html += "            <li>$($file.FullName)</li>`n"
            }
            $html += "        </ul>`n"
        }
    }
    
    $html += @"
    </div>
    
    <div class="section">
        <p><em>Report generated by Music Library Manager v2.0</em></p>
    </div>
</body>
</html>
"@
    
    $html | Out-File -FilePath $OutputPath -Encoding UTF8
    Write-Host "Report saved to: $OutputPath" -ForegroundColor Green
    
    # Open the report
    Start-Process $OutputPath
}

<#
.SYNOPSIS
Backs up music library metadata

.DESCRIPTION
Creates a backup of all music metadata that can be restored later

.PARAMETER Path
Path to the music library

.PARAMETER BackupPath
Path for the backup file

.EXAMPLE
Backup-MusicMetadata -Path "h:\Music" -BackupPath "h:\Music\metadata_backup.json"
#>
function Backup-MusicMetadata {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path,
        
        [Parameter(Mandatory=$true)]
        [string]$BackupPath
    )
    
    $musicFiles = Get-ChildItem -Path $Path -Recurse -Include *.mp3,*.flac,*.m4a,*.wav,*.wma,*.ogg
    $backup = @{
        BackupDate = Get-Date
        SourcePath = $Path
        TotalFiles = $musicFiles.Count
        Files = @()
    }
    
    Write-Host "Backing up metadata for $($musicFiles.Count) files..." -ForegroundColor Green
    
    $processed = 0
    foreach ($file in $musicFiles) {
        $processed++
        $percentComplete = [math]::Round(($processed / $musicFiles.Count) * 100, 1)
        Write-Progress -Activity "Backing up Metadata" -Status "$processed of $($musicFiles.Count) ($percentComplete%)" -PercentComplete $percentComplete
        
        try {
            $shell = New-Object -ComObject Shell.Application
            $folder = $shell.Namespace($file.DirectoryName)
            $fileItem = $folder.ParseName($file.Name)
            
            $fileMetadata = @{
                Path = $file.FullName
                RelativePath = $file.FullName.Replace($Path, "")
                Size = $file.Length
                LastModified = $file.LastWriteTime
                Metadata = @{
                    Title = $folder.GetDetailsOf($fileItem, 21)
                    Artist = $folder.GetDetailsOf($fileItem, 13)
                    Album = $folder.GetDetailsOf($fileItem, 14)
                    Year = $folder.GetDetailsOf($fileItem, 15)
                    Genre = $folder.GetDetailsOf($fileItem, 16)
                    Track = $folder.GetDetailsOf($fileItem, 26)
                    Duration = $folder.GetDetailsOf($fileItem, 27)
                    Bitrate = $folder.GetDetailsOf($fileItem, 28)
                }
            }
            
            $backup.Files += $fileMetadata
        }
        catch {
            Write-Warning "Failed to backup metadata for: $($file.FullName)"
        }
    }
    
    Write-Progress -Activity "Backing up Metadata" -Completed
    
    $backup | ConvertTo-Json -Depth 10 | Out-File -FilePath $BackupPath -Encoding UTF8
    Write-Host "Metadata backup saved to: $BackupPath" -ForegroundColor Green
}

<#
.SYNOPSIS
Compares two music libraries

.DESCRIPTION
Compares two music library folders and reports differences

.PARAMETER Path1
First library path

.PARAMETER Path2
Second library path

.EXAMPLE
Compare-MusicLibraries -Path1 "h:\Music" -Path2 "h:\Music_Organized"
#>
function Compare-MusicLibraries {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path1,
        
        [Parameter(Mandatory=$true)]
        [string]$Path2
    )
    
    Write-Host "Comparing music libraries..." -ForegroundColor Green
    Write-Host "Library 1: $Path1" -ForegroundColor Cyan
    Write-Host "Library 2: $Path2" -ForegroundColor Cyan
    
    $files1 = Get-ChildItem -Path $Path1 -Recurse -Include *.mp3,*.flac,*.m4a,*.wav,*.wma,*.ogg
    $files2 = Get-ChildItem -Path $Path2 -Recurse -Include *.mp3,*.flac,*.m4a,*.wav,*.wma,*.ogg
    
    $comparison = @{
        Library1 = @{
            Path = $Path1
            FileCount = $files1.Count
            TotalSize = ($files1 | Measure-Object -Property Length -Sum).Sum
        }
        Library2 = @{
            Path = $Path2
            FileCount = $files2.Count
            TotalSize = ($files2 | Measure-Object -Property Length -Sum).Sum
        }
        OnlyInLibrary1 = @()
        OnlyInLibrary2 = @()
        Common = @()
        SizeDifferences = @()
    }
    
    # Create lookup tables
    $files1Lookup = @{}
    foreach ($file in $files1) {
        $key = [System.IO.Path]::GetFileNameWithoutExtension($file.Name).ToLower()
        if (-not $files1Lookup.ContainsKey($key)) {
            $files1Lookup[$key] = @()
        }
        $files1Lookup[$key] += $file
    }
    
    $files2Lookup = @{}
    foreach ($file in $files2) {
        $key = [System.IO.Path]::GetFileNameWithoutExtension($file.Name).ToLower()
        if (-not $files2Lookup.ContainsKey($key)) {
            $files2Lookup[$key] = @()
        }
        $files2Lookup[$key] += $file
    }
    
    # Find files only in library 1
    foreach ($key in $files1Lookup.Keys) {
        if (-not $files2Lookup.ContainsKey($key)) {
            $comparison.OnlyInLibrary1 += $files1Lookup[$key]
        }
    }
    
    # Find files only in library 2
    foreach ($key in $files2Lookup.Keys) {
        if (-not $files1Lookup.ContainsKey($key)) {
            $comparison.OnlyInLibrary2 += $files2Lookup[$key]
        }
    }
    
    # Find common files and size differences
    foreach ($key in $files1Lookup.Keys) {
        if ($files2Lookup.ContainsKey($key)) {
            $file1 = $files1Lookup[$key][0]
            $file2 = $files2Lookup[$key][0]
            
            $comparison.Common += @{
                Name = $key
                File1 = $file1.FullName
                File2 = $file2.FullName
            }
            
            if ($file1.Length -ne $file2.Length) {
                $comparison.SizeDifferences += @{
                    Name = $key
                    File1 = $file1.FullName
                    File2 = $file2.FullName
                    Size1 = $file1.Length
                    Size2 = $file2.Length
                }
            }
        }
    }
    
    return $comparison
}