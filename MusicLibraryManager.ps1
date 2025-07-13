# Advanced Music Library Manager with Internet Search Integration
# This application searches the internet, compares local file information with online data,
# corrects metadata, organizes files, and creates comprehensive manifests

param(
    [Parameter(Mandatory=$true)]
    [string]$SourcePath,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "$SourcePath\Organized",
    
    [Parameter(Mandatory=$false)]
    [switch]$DryRun,
    
    [Parameter(Mandatory=$false)]
    [bool]$EnableWebSearch = $true,
    
    [Parameter(Mandatory=$false)]
    [int]$MaxConcurrentSearches = 5,
    
    [Parameter(Mandatory=$false)]
    [string]$LogPath = "$SourcePath\MusicLibraryManager.log"
)

# Import required modules
Add-Type -AssemblyName System.Net.Http

# Global variables
$Global:ProcessedFiles = @()
$Global:SearchCache = @{}
$Global:ErrorLog = @()
$Global:ManifestData = @{
    ProcessingDate = Get-Date
    SourcePath = $SourcePath
    OutputPath = $OutputPath
    TotalFiles = 0
    ProcessedFiles = 0
    CorrectedFiles = 0
    Artists = @{}
    Albums = @{}
    Genres = @{}
    Years = @{}
    Errors = @()
}

# Logging function
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    Write-Host $logEntry
    Add-Content -Path $LogPath -Value $logEntry
}

# Web search functions
class WebSearchProvider {
    [string]$Name
    [string]$BaseUrl
    [hashtable]$Headers
    
    WebSearchProvider([string]$name, [string]$baseUrl) {
        $this.Name = $name
        $this.BaseUrl = $baseUrl
        $this.Headers = @{
            'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
    }
    
    [object] Search([string]$query) {
        return $null
    }
}

class MusicBrainzSearchProvider : WebSearchProvider {
    MusicBrainzSearchProvider() : base("MusicBrainz", "https://musicbrainz.org/ws/2") {}
    
    [object] Search([string]$artist, [string]$album, [string]$track) {
        try {
            $query = ""
            if ($artist) { $query += "artist:$artist " }
            if ($album) { $query += "release:$album " }
            if ($track) { $query += "recording:$track" }
            
            $encodedQuery = [System.Uri]::EscapeDataString($query.Trim())
            $url = "$($this.BaseUrl)/recording/?query=$encodedQuery&fmt=json&limit=5"
            
            $response = Invoke-RestMethod -Uri $url -Headers $this.Headers -TimeoutSec 10
            return $response
        }
        catch {
            Write-Log "MusicBrainz search failed: $($_.Exception.Message)" "ERROR"
            return $null
        }
    }
}

class LastFmSearchProvider : WebSearchProvider {
    [string]$ApiKey
    
    LastFmSearchProvider([string]$apiKey) : base("Last.fm", "https://ws.audioscrobbler.com/2.0") {
        $this.ApiKey = $apiKey
    }
    
    [object] SearchArtist([string]$artist) {
        try {
            $url = "$($this.BaseUrl)/?method=artist.getinfo&artist=$([System.Uri]::EscapeDataString($artist))&api_key=$($this.ApiKey)&format=json"
            $response = Invoke-RestMethod -Uri $url -Headers $this.Headers -TimeoutSec 10
            return $response
        }
        catch {
            Write-Log "Last.fm artist search failed: $($_.Exception.Message)" "ERROR"
            return $null
        }
    }
    
    [object] SearchAlbum([string]$artist, [string]$album) {
        try {
            $url = "$($this.BaseUrl)/?method=album.getinfo&artist=$([System.Uri]::EscapeDataString($artist))&album=$([System.Uri]::EscapeDataString($album))&api_key=$($this.ApiKey)&format=json"
            $response = Invoke-RestMethod -Uri $url -Headers $this.Headers -TimeoutSec 10
            return $response
        }
        catch {
            Write-Log "Last.fm album search failed: $($_.Exception.Message)" "ERROR"
            return $null
        }
    }
}

class DiscogsSearchProvider : WebSearchProvider {
    [string]$Token
    
    DiscogsSearchProvider([string]$token) : base("Discogs", "https://api.discogs.com") {
        $this.Token = $token
        $this.Headers['Authorization'] = "Discogs token=$token"
    }
    
    [object] Search([string]$artist, [string]$album) {
        try {
            $query = "$artist $album"
            $encodedQuery = [System.Uri]::EscapeDataString($query)
            $url = "$($this.BaseUrl)/database/search?q=$encodedQuery&type=release"
            
            $response = Invoke-RestMethod -Uri $url -Headers $this.Headers -TimeoutSec 10
            return $response
        }
        catch {
            Write-Log "Discogs search failed: $($_.Exception.Message)" "ERROR"
            return $null
        }
    }
}

# Metadata extraction and correction
class MusicFileAnalyzer {
    [string]$FilePath
    [hashtable]$RawMetadata
    [hashtable]$CorrectedMetadata
    [hashtable]$OnlineData
    
    MusicFileAnalyzer([string]$filePath) {
        $this.FilePath = $filePath
        $this.RawMetadata = @{}
        $this.CorrectedMetadata = @{}
        $this.OnlineData = @{}
        $this.ExtractMetadata()
    }
    
    [void] ExtractMetadata() {
        try {
            # Extract metadata using Shell.Application
            $shell = New-Object -ComObject Shell.Application
            $folder = $shell.Namespace((Get-Item $this.FilePath).DirectoryName)
            $file = $folder.ParseName((Get-Item $this.FilePath).Name)
            
            # Common metadata properties
            $properties = @{
                'Title' = 21
                'Artist' = 13
                'Album' = 14
                'Year' = 15
                'Genre' = 16
                'Track' = 26
                'Duration' = 27
                'Bitrate' = 28
            }
            
            foreach ($prop in $properties.GetEnumerator()) {
                $value = $folder.GetDetailsOf($file, $prop.Value)
                if ($value -and $value.Trim() -ne "") {
                    $this.RawMetadata[$prop.Key] = $value.Trim()
                }
            }
            
            # Extract from filename and folder structure
            $this.ExtractFromPath()
            
        }
        catch {
            Write-Log "Metadata extraction failed for $($this.FilePath): $($_.Exception.Message)" "ERROR"
        }
    }
    
    [void] ExtractFromPath() {
        $fileInfo = Get-Item $this.FilePath
        $fileName = [System.IO.Path]::GetFileNameWithoutExtension($fileInfo.Name)
        $folderName = $fileInfo.Directory.Name
        $parentFolder = $fileInfo.Directory.Parent.Name
        
        # Try to extract year from folder name (including parentheses)
        if ($folderName -match '\((\d{4})\)') {
            $this.RawMetadata['YearFromFolder'] = $matches[1]
        }
        elseif ($folderName -match '(\d{4})') {
            $this.RawMetadata['YearFromFolder'] = $matches[1]
        }
        
        # Try to extract artist and album from folder structure
        # Handle formats like: (2000) - Nine Inch Nails - Things Falling Apart [16Bit-44.1kHz]
        if ($folderName -match '^\((\d{4})\)\s*-\s*(.+?)\s*-\s*(.+?)(?:\s*\[.*\])?$') {
            $this.RawMetadata['YearFromFolder'] = $matches[1]
            $this.RawMetadata['ArtistFromFolder'] = $matches[2].Trim()
            $this.RawMetadata['AlbumFromFolder'] = $matches[3].Trim()
        }
        # Handle formats like: 1989 - Artist - Album
        elseif ($folderName -match '^(\d{4})\s*-\s*(.+?)\s*-\s*(.+?)(?:\s*\[.*\])?$') {
            $this.RawMetadata['YearFromFolder'] = $matches[1]
            $this.RawMetadata['ArtistFromFolder'] = $matches[2].Trim()
            $this.RawMetadata['AlbumFromFolder'] = $matches[3].Trim()
        }
        # Handle formats like: Artist - Album
        elseif ($folderName -match '^(.+?)\s*-\s*(.+?)(?:\s*\[.*\])?$') {
            $this.RawMetadata['ArtistFromFolder'] = $matches[1].Trim()
            $this.RawMetadata['AlbumFromFolder'] = $matches[2].Trim()
        }
        
        # Try to extract track info from filename
        if ($fileName -match '^(\d+)\s*[-._]\s*(.+)$') {
            $this.RawMetadata['TrackFromFile'] = $matches[1]
            $this.RawMetadata['TitleFromFile'] = $matches[2].Trim()
        }
    }
    
    [void] SearchOnline([WebSearchProvider[]]$providers) {
        if (-not $Global:EnableWebSearch) { return }
        
        $artist = $this.GetBestValue('Artist')
        $album = $this.GetBestValue('Album')
        $title = $this.GetBestValue('Title')
        
        if (-not $artist -and -not $album -and -not $title) {
            Write-Log "Insufficient metadata for online search: $($this.FilePath)" "WARNING"
            return
        }
        
        foreach ($provider in $providers) {
            try {
                $searchKey = "$($provider.Name):${artist}:${album}:${title}"
                
                if ($Global:SearchCache.ContainsKey($searchKey)) {
                    $this.OnlineData[$provider.Name] = $Global:SearchCache[$searchKey]
                    continue
                }
                
                $result = $null
                switch ($provider.Name) {
                    "MusicBrainz" {
                        $result = $provider.Search($artist, $album, $title)
                    }
                    "Last.fm" {
                        if ($artist) {
                            $artistInfo = $provider.SearchArtist($artist)
                            $albumInfo = $provider.SearchAlbum($artist, $album)
                            $result = @{ Artist = $artistInfo; Album = $albumInfo }
                        }
                    }
                    "Discogs" {
                        $result = $provider.Search($artist, $album)
                    }
                }
                
                if ($result) {
                    $this.OnlineData[$provider.Name] = $result
                    $Global:SearchCache[$searchKey] = $result
                }
                
                Start-Sleep -Milliseconds 200  # Rate limiting
            }
            catch {
                Write-Log "Online search failed with $($provider.Name): $($_.Exception.Message)" "ERROR"
            }
        }
    }
    
    [string] GetBestValue([string]$property) {
        # Priority order for metadata sources
        $sources = @(
            $property,
            "${property}FromFolder",
            "${property}FromFile"
        )
        
        foreach ($source in $sources) {
            if ($this.RawMetadata.ContainsKey($source) -and $this.RawMetadata[$source]) {
                return $this.RawMetadata[$source]
            }
        }
        
        return ""
    }
    
    [void] CorrectMetadata() {
        # Start with raw metadata
        $this.CorrectedMetadata = $this.RawMetadata.Clone()
        
        # Apply corrections based on online data
        $this.ApplyOnlineCorrections()
        
        # Apply standardization rules
        $this.StandardizeMetadata()
        
        # Validate and clean
        $this.ValidateMetadata()
    }
    
    [void] ApplyOnlineCorrections() {
        # MusicBrainz corrections
        if ($this.OnlineData.ContainsKey('MusicBrainz') -and $this.OnlineData['MusicBrainz'].recordings) {
            $recording = $this.OnlineData['MusicBrainz'].recordings[0]
            if ($recording.title) {
                $this.CorrectedMetadata['Title'] = $recording.title
            }
            if ($recording.'artist-credit' -and $recording.'artist-credit'[0].name) {
                $this.CorrectedMetadata['Artist'] = $recording.'artist-credit'[0].name
            }
            if ($recording.releases -and $recording.releases[0].title) {
                $this.CorrectedMetadata['Album'] = $recording.releases[0].title
            }
        }
        
        # Last.fm corrections
        if ($this.OnlineData.ContainsKey('Last.fm')) {
            $lastfmData = $this.OnlineData['Last.fm']
            if ($lastfmData.Artist -and $lastfmData.Artist.artist) {
                if ($lastfmData.Artist.artist.tags -and $lastfmData.Artist.artist.tags.tag) {
                    $topGenre = $lastfmData.Artist.artist.tags.tag[0].name
                    if ($topGenre) {
                        $this.CorrectedMetadata['Genre'] = $topGenre
                    }
                }
            }
        }
        
        # Discogs corrections
        if ($this.OnlineData.ContainsKey('Discogs') -and $this.OnlineData['Discogs'].results) {
            $release = $this.OnlineData['Discogs'].results[0]
            if ($release.year) {
                $this.CorrectedMetadata['Year'] = $release.year.ToString()
            }
            if ($release.genre -and $release.genre.Count -gt 0) {
                $this.CorrectedMetadata['Genre'] = $release.genre[0]
            }
        }
    }
    
    [void] StandardizeMetadata() {
        # Standardize artist names
        if ($this.CorrectedMetadata.ContainsKey('Artist')) {
            $this.CorrectedMetadata['Artist'] = $this.StandardizeArtistName($this.CorrectedMetadata['Artist'])
        }
        
        # Standardize genre
        if ($this.CorrectedMetadata.ContainsKey('Genre')) {
            $this.CorrectedMetadata['Genre'] = $this.StandardizeGenre($this.CorrectedMetadata['Genre'])
        }
        
        # Ensure year is 4 digits
        if ($this.CorrectedMetadata.ContainsKey('Year')) {
            $year = $this.CorrectedMetadata['Year'] -replace '[^\d]', ''
            if ($year.Length -eq 4) {
                $this.CorrectedMetadata['Year'] = $year
            }
        }
    }
    
    [string] StandardizeArtistName([string]$artist) {
        # Remove common prefixes/suffixes
        $artist = $artist -replace '^(The|A|An)\s+', ''
        $artist = $artist -replace '\s+(feat\.|ft\.|featuring).*$', ''
        return $artist.Trim()
    }
    
    [string] StandardizeGenre([string]$genre) {
        # Remove duplicates and clean up genre string
        if ([string]::IsNullOrWhiteSpace($genre)) {
            return 'Unknown'
        }
        
        # Split by semicolon and take only the first unique genre
        $genres = $genre -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -Unique -First 1
        $cleanGenre = $genres[0]
        
        # Map common genre variations to standard names
        $genreMap = @{
            'Alternative Rock' = 'Alternative'
            'Alt Rock' = 'Alternative'
            'Hip Hop' = 'Hip-Hop'
            'Rap' = 'Hip-Hop'
            'R&B' = 'R&B/Soul'
            'Rhythm and Blues' = 'R&B/Soul'
            'Electronic Dance Music' = 'Electronic'
            'EDM' = 'Electronic'
            'Alternative Rock/Grunge' = 'Grunge'
        }
        
        if ($genreMap.ContainsKey($cleanGenre)) {
            return $genreMap[$cleanGenre]
        }
        
        # Limit length to prevent path issues
        if ($cleanGenre.Length -gt 20) {
            $cleanGenre = $cleanGenre.Substring(0, 20)
        }
        
        return $cleanGenre
    }
    
    [void] ValidateMetadata() {
        # Ensure required fields have values
        $requiredFields = @('Artist', 'Album', 'Title')
        
        foreach ($field in $requiredFields) {
            if (-not $this.CorrectedMetadata.ContainsKey($field) -or -not $this.CorrectedMetadata[$field]) {
                $fallback = $this.GetFallbackValue($field)
                if ($fallback) {
                    $this.CorrectedMetadata[$field] = $fallback
                    Write-Log "Used fallback for $field in $($this.FilePath): $fallback" "WARNING"
                }
            }
        }
    }
    
    [string] GetFallbackValue([string]$field) {
        switch ($field) {
            'Artist' {
                $value = $this.GetBestValue('Artist')
                if ($value) { return $value } else { return 'Unknown Artist' }
            }
            'Album' {
                $value = $this.GetBestValue('Album')
                if ($value) { return $value } else { return 'Unknown Album' }
            }
            'Title' {
                $fileName = [System.IO.Path]::GetFileNameWithoutExtension($this.FilePath)
                return $fileName -replace '^\d+\s*[-._]\s*', ''
            }
            default {
                return 'Unknown'
            }
        }
        return 'Unknown'
    }
}

# File organization functions
function Get-SafeFileName {
    param([string]$Name)
    
    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars() + [System.IO.Path]::GetInvalidPathChars()
    $safeName = $Name
    
    foreach ($char in $invalidChars) {
        $safeName = $safeName.Replace($char, '_')
    }
    
    return $safeName.Trim()
}

function New-OrganizedPath {
    param(
        [hashtable]$Metadata,
        [string]$BaseOutputPath,
        [string]$OriginalExtension
    )
    
    # Debug: Show what we're working with
    Write-Log "DEBUG - New-OrganizedPath input metadata:" "INFO"
    foreach ($key in $Metadata.Keys) {
        Write-Log "  $key = '$($Metadata[$key])'" "INFO"
    }
    
    $genre = if ($Metadata.ContainsKey('Genre') -and $Metadata['Genre']) { Get-SafeFileName($Metadata['Genre']) } else { 'Unknown' }
    $artist = if ($Metadata.ContainsKey('Artist') -and $Metadata['Artist']) { Get-SafeFileName($Metadata['Artist']) } else { 'Unknown Artist' }
    $year = if ($Metadata.ContainsKey('Year') -and $Metadata['Year']) { $Metadata['Year'] } else { 'Unknown' }
    $album = if ($Metadata.ContainsKey('Album') -and $Metadata['Album']) { Get-SafeFileName($Metadata['Album']) } else { 'Unknown Album' }
    $track = if ($Metadata.ContainsKey('Track') -and $Metadata['Track']) { $Metadata['Track'].PadLeft(2, '0') } else { '00' }
    $title = if ($Metadata.ContainsKey('Title') -and $Metadata['Title']) { Get-SafeFileName($Metadata['Title']) } else { 'Unknown Title' }
    
    Write-Log "DEBUG - Processed values: Genre='$genre', Artist='$artist', Year='$year', Album='$album', Track='$track', Title='$title'" "INFO"
    
    $artistFolder = Join-Path (Join-Path $BaseOutputPath $genre) $artist
    $albumFolder = Join-Path $artistFolder "$year - $album"
    $fileName = "$track - $title$OriginalExtension"
    $fullPath = Join-Path $albumFolder $fileName
    
    # Check path length and truncate if necessary (Windows 260 char limit)
    if ($fullPath.Length -gt 250) {
        # Truncate title to make path shorter
        $maxTitleLength = 250 - ($fullPath.Length - $title.Length)
        if ($maxTitleLength -gt 10) {
            $title = $title.Substring(0, [Math]::Min($title.Length, $maxTitleLength))
            $fileName = "$track - $title$OriginalExtension"
            $fullPath = Join-Path $albumFolder $fileName
        }
    }
    
    return @{
        ArtistFolder = $artistFolder
        AlbumFolder = $albumFolder
        FullPath = $fullPath
        FileName = $fileName
    }
}

function Copy-MusicFile {
    param(
        [string]$SourcePath,
        [string]$DestinationPath,
        [hashtable]$Metadata,
        [switch]$DryRun
    )
    
    if ($DryRun) {
        Write-Log "[DRY RUN] Would copy: $SourcePath -> $DestinationPath" "INFO"
        return $true
    }
    
    try {
        $destDir = Split-Path $DestinationPath -Parent
        if (-not (Test-Path $destDir)) {
            New-Item -Path $destDir -ItemType Directory -Force | Out-Null
        }
        
        Copy-Item -Path $SourcePath -Destination $DestinationPath -Force
        Write-Log "Copied: $SourcePath -> $DestinationPath" "INFO"
        return $true
    }
    catch {
        Write-Log "Failed to copy file: $($_.Exception.Message)" "ERROR"
        $Global:ErrorLog += @{
            File = $SourcePath
            Error = $_.Exception.Message
            Operation = 'Copy'
        }
        return $false
    }
}

# Manifest and reporting functions
function Update-Manifest {
    param(
        [string]$OriginalPath,
        [string]$NewPath,
        [hashtable]$OriginalMetadata,
        [hashtable]$CorrectedMetadata,
        [hashtable]$OnlineData
    )
    
    $fileEntry = @{
        OriginalPath = $OriginalPath
        NewPath = $NewPath
        OriginalMetadata = $OriginalMetadata
        CorrectedMetadata = $CorrectedMetadata
        OnlineData = $OnlineData
        ProcessedDate = Get-Date
        WasCorrected = $false
    }
    
    # Check if metadata was corrected
    foreach ($key in $CorrectedMetadata.Keys) {
        if ($OriginalMetadata[$key] -ne $CorrectedMetadata[$key]) {
            $fileEntry.WasCorrected = $true
            break
        }
    }
    
    $Global:ProcessedFiles += $fileEntry
    
    # Update statistics
    $artist = $CorrectedMetadata['Artist'] -or 'Unknown'
    $album = $CorrectedMetadata['Album'] -or 'Unknown'
    $genre = $CorrectedMetadata['Genre'] -or 'Unknown'
    $year = $CorrectedMetadata['Year'] -or 'Unknown'
    
    if (-not $Global:ManifestData.Artists.ContainsKey($artist)) {
        $Global:ManifestData.Artists[$artist] = 0
    }
    $Global:ManifestData.Artists[$artist]++
    
    if (-not $Global:ManifestData.Albums.ContainsKey($album)) {
        $Global:ManifestData.Albums[$album] = 0
    }
    $Global:ManifestData.Albums[$album]++
    
    if (-not $Global:ManifestData.Genres.ContainsKey($genre)) {
        $Global:ManifestData.Genres[$genre] = 0
    }
    $Global:ManifestData.Genres[$genre]++
    
    if (-not $Global:ManifestData.Years.ContainsKey($year)) {
        $Global:ManifestData.Years[$year] = 0
    }
    $Global:ManifestData.Years[$year]++
    
    if ($fileEntry.WasCorrected) {
        $Global:ManifestData.CorrectedFiles++
    }
}

function Export-Manifest {
    param([string]$OutputPath)
    
    $Global:ManifestData.TotalFiles = (Get-ChildItem -Path $SourcePath -Recurse -Include *.mp3,*.flac,*.m4a,*.wav,*.wma).Count
    $Global:ManifestData.ProcessedFiles = $Global:ProcessedFiles.Count
    $Global:ManifestData.Errors = $Global:ErrorLog
    
    $manifestPath = Join-Path $OutputPath "MusicLibraryManifest.json"
    
    try {
        $Global:ManifestData | ConvertTo-Json -Depth 10 | Out-File -FilePath $manifestPath -Encoding UTF8
        Write-Log "Manifest exported to: $manifestPath" "INFO"
        
        # Also create a summary report
        $summaryPath = Join-Path $OutputPath "ProcessingSummary.txt"
        $summary = @"
Music Library Processing Summary
================================
Processing Date: $($Global:ManifestData.ProcessingDate)
Source Path: $($Global:ManifestData.SourcePath)
Output Path: $($Global:ManifestData.OutputPath)

Statistics:
-----------
Total Files Found: $($Global:ManifestData.TotalFiles)
Files Processed: $($Global:ManifestData.ProcessedFiles)
Files Corrected: $($Global:ManifestData.CorrectedFiles)
Unique Artists: $($Global:ManifestData.Artists.Count)
Unique Albums: $($Global:ManifestData.Albums.Count)
Unique Genres: $($Global:ManifestData.Genres.Count)
Year Range: $($Global:ManifestData.Years.Keys | Sort-Object | Select-Object -First 1) - $($Global:ManifestData.Years.Keys | Sort-Object | Select-Object -Last 1)

Top Genres:
-----------
$($Global:ManifestData.Genres.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 10 | ForEach-Object { "$($_.Key): $($_.Value) tracks" } | Out-String)

Top Artists:
------------
$($Global:ManifestData.Artists.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 10 | ForEach-Object { "$($_.Key): $($_.Value) tracks" } | Out-String)

Errors:
-------
$($Global:ErrorLog.Count) errors encountered
"@
        
        $summary | Out-File -FilePath $summaryPath -Encoding UTF8
        Write-Log "Summary report exported to: $summaryPath" "INFO"
    }
    catch {
        Write-Log "Failed to export manifest: $($_.Exception.Message)" "ERROR"
    }
}

# File organization functions
function New-OrganizedPath {
    param(
        [hashtable]$Metadata,
        [string]$BaseOutputPath,
        [string]$OriginalExtension
    )
    
    # Use explicit null/empty checks instead of -or operator
    $artist = if ($Metadata.ContainsKey('Artist') -and ![string]::IsNullOrWhiteSpace($Metadata['Artist'])) { $Metadata['Artist'] } else { 'Unknown Artist' }
    $album = if ($Metadata.ContainsKey('Album') -and ![string]::IsNullOrWhiteSpace($Metadata['Album'])) { $Metadata['Album'] } else { 'Unknown Album' }
    $title = if ($Metadata.ContainsKey('Title') -and ![string]::IsNullOrWhiteSpace($Metadata['Title'])) { $Metadata['Title'] } else { 'Unknown Title' }
    $year = if ($Metadata.ContainsKey('Year') -and ![string]::IsNullOrWhiteSpace($Metadata['Year'])) { $Metadata['Year'] } else { 'Unknown' }
    $genre = if ($Metadata.ContainsKey('Genre') -and ![string]::IsNullOrWhiteSpace($Metadata['Genre'])) { $Metadata['Genre'] } else { 'Unknown' }
    $track = if ($Metadata.ContainsKey('Track') -and ![string]::IsNullOrWhiteSpace($Metadata['Track'])) { $Metadata['Track'] } else { '01' }
    
    # Debug logging
    Write-Log "DEBUG - New-OrganizedPath values: Artist='$artist', Album='$album', Title='$title', Year='$year', Genre='$genre', Track='$track'" "INFO"
    
    # Sanitize names for file system
    $artist = $artist -replace '[<>:"/\\|?*]', '_'
    $album = $album -replace '[<>:"/\\|?*]', '_'
    $title = $title -replace '[<>:"/\\|?*]', '_'
    $genre = $genre -replace '[<>:"/\\|?*]', '_'
    
    # Create folder structure: Genre\Artist\Year - Album
    $folderPath = Join-Path $BaseOutputPath $genre
    $folderPath = Join-Path $folderPath $artist
    if ($year -ne 'Unknown') {
        $folderPath = Join-Path $folderPath "$year - $album"
    } else {
        $folderPath = Join-Path $folderPath $album
    }
    
    # Create filename: Track - Title.ext
    $trackNum = $track.ToString().PadLeft(2, '0')
    $fileName = "$trackNum - $title$OriginalExtension"
    $fullPath = Join-Path $folderPath $fileName
    
    return @{
        FolderPath = $folderPath
        FileName = $fileName
        FullPath = $fullPath
    }
}

function Copy-MusicFile {
    param(
        [string]$SourcePath,
        [string]$DestinationPath,
        [hashtable]$Metadata,
        [switch]$DryRun
    )
    
    try {
        if ($DryRun) {
            Write-Log "[DRY RUN] Would copy: $SourcePath -> $DestinationPath" "INFO"
            return $true
        }
        
        $destDir = Split-Path $DestinationPath -Parent
        if (-not (Test-Path $destDir)) {
            New-Item -Path $destDir -ItemType Directory -Force | Out-Null
        }
        
        Copy-Item -Path $SourcePath -Destination $DestinationPath -Force
        Write-Log "Copied: $SourcePath -> $DestinationPath" "INFO"
        return $true
    }
    catch {
        Write-Log "Failed to copy file: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Update-Manifest {
    param(
        [string]$OriginalPath,
        [string]$NewPath,
        [hashtable]$OriginalMetadata,
        [hashtable]$CorrectedMetadata,
        [hashtable]$OnlineData
    )
    
    $fileInfo = @{
        OriginalPath = $OriginalPath
        NewPath = $NewPath
        OriginalMetadata = $OriginalMetadata
        CorrectedMetadata = $CorrectedMetadata
        OnlineData = $OnlineData
        ProcessedDate = Get-Date
    }
    
    $Global:ProcessedFiles += $fileInfo
    $Global:ManifestData.ProcessedFiles++
    
    # Update statistics
    $artist = $CorrectedMetadata['Artist'] -or 'Unknown Artist'
    $album = $CorrectedMetadata['Album'] -or 'Unknown Album'
    $genre = $CorrectedMetadata['Genre'] -or 'Unknown'
    $year = $CorrectedMetadata['Year'] -or 'Unknown'
    
    if ($Global:ManifestData.Artists.ContainsKey($artist)) {
        $Global:ManifestData.Artists[$artist]++
    } else {
        $Global:ManifestData.Artists[$artist] = 1
    }
    
    if ($Global:ManifestData.Albums.ContainsKey($album)) {
        $Global:ManifestData.Albums[$album]++
    } else {
        $Global:ManifestData.Albums[$album] = 1
    }
    
    if ($Global:ManifestData.Genres.ContainsKey($genre)) {
        $Global:ManifestData.Genres[$genre]++
    } else {
        $Global:ManifestData.Genres[$genre] = 1
    }
    
    if ($Global:ManifestData.Years.ContainsKey($year)) {
        $Global:ManifestData.Years[$year]++
    } else {
        $Global:ManifestData.Years[$year] = 1
    }
    
    # Check if metadata was corrected
    $wasCorrect = $false
    foreach ($key in $CorrectedMetadata.Keys) {
        if ($OriginalMetadata[$key] -ne $CorrectedMetadata[$key]) {
            $wasCorrect = $true
            break
        }
    }
    
    if ($wasCorrect) {
        $Global:ManifestData.CorrectedFiles++
    }
}

# Playlist generation functions
function New-Playlists {
    param([string]$OutputPath)
    
    Write-Log "Generating playlists..." "INFO"
    
    # Genre playlists
    foreach ($genre in $Global:ManifestData.Genres.Keys) {
        $playlistPath = Join-Path $OutputPath "Playlists" "By Genre" "$genre.m3u"
        $tracks = $Global:ProcessedFiles | Where-Object { $_.CorrectedMetadata['Genre'] -eq $genre }
        New-PlaylistFile -Tracks $tracks -PlaylistPath $playlistPath -Title "$genre Music"
    }
    
    # Artist playlists
    foreach ($artist in $Global:ManifestData.Artists.Keys) {
        $playlistPath = Join-Path $OutputPath "Playlists" "By Artist" "$artist.m3u"
        $tracks = $Global:ProcessedFiles | Where-Object { $_.CorrectedMetadata['Artist'] -eq $artist }
        New-PlaylistFile -Tracks $tracks -PlaylistPath $playlistPath -Title "$artist"
    }
    
    # Year playlists
    foreach ($year in $Global:ManifestData.Years.Keys) {
        if ($year -ne 'Unknown') {
            $playlistPath = Join-Path $OutputPath "Playlists" "By Year" "$year.m3u"
            $tracks = $Global:ProcessedFiles | Where-Object { $_.CorrectedMetadata['Year'] -eq $year }
            New-PlaylistFile -Tracks $tracks -PlaylistPath $playlistPath -Title "Music from $year"
        }
    }
}

function New-PlaylistFile {
    param(
        [array]$Tracks,
        [string]$PlaylistPath,
        [string]$Title
    )
    
    if ($Tracks.Count -eq 0) { return }
    
    try {
        $playlistDir = Split-Path $PlaylistPath -Parent
        if (-not (Test-Path $playlistDir)) {
            New-Item -Path $playlistDir -ItemType Directory -Force | Out-Null
        }
        
        $content = @("#EXTM3U")
        $content += "#PLAYLIST:$Title"
        
        foreach ($track in $Tracks) {
            $metadata = $track.CorrectedMetadata
            $artist = $metadata['Artist'] -or 'Unknown Artist'
            $title = $metadata['Title'] -or 'Unknown Title'
            $duration = $metadata['Duration'] -or ''
            
            if ($duration) {
                $content += "#EXTINF:$duration,$artist - $title"
            } else {
                $content += "#EXTINF:-1,$artist - $title"
            }
            $content += $track.NewPath
        }
        
        $content | Out-File -FilePath $PlaylistPath -Encoding UTF8
        Write-Log "Created playlist: $PlaylistPath ($($Tracks.Count) tracks)" "INFO"
    }
    catch {
        Write-Log "Failed to create playlist ${PlaylistPath}: $($_.Exception.Message)" "ERROR"
    }
}

# Main processing function
function Start-MusicLibraryProcessing {
    # Set global variables
    $Global:EnableWebSearch = $EnableWebSearch
    
    Write-Log "Starting Music Library Manager" "INFO"
    Write-Log "Source: $SourcePath" "INFO"
    Write-Log "Output: $OutputPath" "INFO"
    Write-Log "Dry Run: $DryRun" "INFO"
    Write-Log "Web Search: $EnableWebSearch" "INFO"
    
    # Initialize search providers
    $searchProviders = @()
    $searchProviders += [MusicBrainzSearchProvider]::new()
    
    # Note: Add API keys for Last.fm and Discogs if available
    # $searchProviders += [LastFmSearchProvider]::new("YOUR_LASTFM_API_KEY")
    # $searchProviders += [DiscogsSearchProvider]::new("YOUR_DISCOGS_TOKEN")
    
    # Get all music files
    $musicFiles = Get-ChildItem -Path $SourcePath -Recurse -Include *.mp3,*.flac,*.m4a,*.wav,*.wma
    Write-Log "Found $($musicFiles.Count) music files" "INFO"
    
    # Update total files count
    $Global:ManifestData.TotalFiles = $musicFiles.Count
    
    $processedCount = 0
    $totalFiles = $musicFiles.Count
    
    foreach ($file in $musicFiles) {
        $processedCount++
        $percentComplete = [math]::Round(($processedCount / $totalFiles) * 100, 1)
        
        Write-Progress -Activity "Processing Music Files" -Status "$processedCount of $totalFiles ($percentComplete%)" -PercentComplete $percentComplete
        Write-Log "Processing ($processedCount/$totalFiles): $($file.FullName)" "INFO"
        
        try {
            # Analyze file
            $analyzer = [MusicFileAnalyzer]::new($file.FullName)
            
            # Search online for corrections
            $analyzer.SearchOnline($searchProviders)
            
            # Apply corrections
            $analyzer.CorrectMetadata()
            
            # Debug: Show corrected metadata
            Write-Log "DEBUG - CorrectedMetadata for $($file.Name):" "INFO"
            foreach ($key in $analyzer.CorrectedMetadata.Keys) {
                Write-Log "  $key = '$($analyzer.CorrectedMetadata[$key])'" "INFO"
            }
            
            # Determine new file path
            $newPath = New-OrganizedPath -Metadata $analyzer.CorrectedMetadata -BaseOutputPath $OutputPath -OriginalExtension $file.Extension
            
            # Copy file to new location
            $copySuccess = Copy-MusicFile -SourcePath $file.FullName -DestinationPath $newPath.FullPath -Metadata $analyzer.CorrectedMetadata -DryRun:$DryRun
            
            if ($copySuccess) {
                # Update manifest
                Update-Manifest -OriginalPath $file.FullName -NewPath $newPath.FullPath -OriginalMetadata $analyzer.RawMetadata -CorrectedMetadata $analyzer.CorrectedMetadata -OnlineData $analyzer.OnlineData
            }
        }
        catch {
            Write-Log "Error processing $($file.FullName): $($_.Exception.Message)" "ERROR"
            $Global:ErrorLog += @{
                File = $file.FullName
                Error = $_.Exception.Message
                Operation = 'Processing'
            }
        }
        
        # Rate limiting for web searches
        if ($EnableWebSearch -and ($processedCount % 10) -eq 0) {
            Start-Sleep -Milliseconds 500
        }
    }
    
    Write-Progress -Activity "Processing Music Files" -Completed
    
    # Generate playlists
    if (-not $DryRun) {
        New-Playlists -OutputPath $OutputPath
    }
    
    # Export manifest and summary
    Export-Manifest -OutputPath $OutputPath
    
    Write-Log "Processing completed!" "INFO"
    Write-Log "Processed: $($Global:ProcessedFiles.Count) files" "INFO"
    Write-Log "Corrected: $($Global:ManifestData.CorrectedFiles) files" "INFO"
    Write-Log "Errors: $($Global:ErrorLog.Count)" "INFO"
    Write-Log "Artists: $($Global:ManifestData.Artists.Count)" "INFO"
    Write-Log "Albums: $($Global:ManifestData.Albums.Count)" "INFO"
    Write-Log "Genres: $($Global:ManifestData.Genres.Count)" "INFO"
}

# Start processing
Start-MusicLibraryProcessing