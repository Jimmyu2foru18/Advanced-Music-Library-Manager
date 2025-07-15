# Music Library Organization Script
# This script will reorganize your music collection with proper naming conventions
# and remove all existing playlist files

Write-Host "Starting Music Library Organization..." -ForegroundColor Green

# Function to sanitize folder/file names
function Sanitize-Name {
    param([string]$name)
    # Remove invalid characters and clean up the name
    $sanitized = $name -replace '[\\/:*?"<>|]', ''
    $sanitized = $sanitized -replace '\s+', ' '  # Replace multiple spaces with single space
    $sanitized = $sanitized.Trim()
    return $sanitized
}

# Function to extract artist from folder name
function Get-Artist {
    param([string]$folderName)
    
    # Common patterns to extract artist names
    if ($folderName -match '^\d{4}\s*-?\s*(.+?)\s*-\s*(.+)$') {
        return $matches[1].Trim()
    }
    elseif ($folderName -match '^(.+?)\s*-\s*(.+)$') {
        return $matches[1].Trim()
    }
    elseif ($folderName -match '^\[?\d{4}\]?\s*(.+)$') {
        return $matches[1].Trim()
    }
    else {
        return $folderName
    }
}

# Function to extract album from folder name
function Get-Album {
    param([string]$folderName)
    
    if ($folderName -match '^\d{4}\s*-?\s*(.+?)\s*-\s*(.+)$') {
        return $matches[2].Trim()
    }
    elseif ($folderName -match '^(.+?)\s*-\s*(.+)$') {
        return $matches[2].Trim()
    }
    else {
        return $folderName
    }
}

# Function to extract year from folder name
function Get-Year {
    param([string]$folderName)
    
    if ($folderName -match '^(\d{4})') {
        return $matches[1]
    }
    elseif ($folderName -match '\[(\d{4})\]') {
        return $matches[1]
    }
    else {
        return "Unknown"
    }
}

# Function to determine genre based on artist
function Get-Genre {
    param([string]$artist)
    
    $artist = $artist.ToLower()
    
    # Define genre mappings
    $genreMap = @{
        'slipknot' = 'Metal'
        'disturbed' = 'Metal'
        'system of a down' = 'Metal'
        'rise against' = 'Punk Rock'
        'nirvana' = 'Grunge'
        'pearl jam' = 'Grunge'
        'red hot chili peppers' = 'Alternative Rock'
        'bon jovi' = 'Rock'
        'eminem' = 'Hip Hop'
        'd12' = 'Hip Hop'
        'j. cole' = 'Hip Hop'
        'nas' = 'Hip Hop'
        'gorillaz' = 'Alternative'
        'linkin park' = 'Nu Metal'
        'maroon 5' = 'Pop Rock'
        'evanescence' = 'Gothic Rock'
        'rezz' = 'Electronic'
        'joy division' = 'Post-Punk'
        'stone sour' = 'Alternative Metal'
        'falling in reverse' = 'Post-Hardcore'
        'arctic monkeys' = 'Indie Rock'
    }
    
    foreach ($key in $genreMap.Keys) {
        if ($artist -like "*$key*") {
            return $genreMap[$key]
        }
    }
    
    return 'Unknown'
}

# Step 1: Remove all existing playlist files
Write-Host "Removing existing playlist files..." -ForegroundColor Yellow
$playlistFiles = Get-ChildItem -Path "h:\Music" -Filter "*.m3u" -Recurse
foreach ($playlist in $playlistFiles) {
    Write-Host "Removing: $($playlist.FullName)" -ForegroundColor Red
    Remove-Item $playlist.FullName -Force
}
Write-Host "Removed $($playlistFiles.Count) playlist files." -ForegroundColor Green

# Step 2: Create organized directory structure
$baseDir = "h:\Music"
$organizedDir = "h:\Music_Organized"

if (Test-Path $organizedDir) {
    Write-Host "Removing existing organized directory..." -ForegroundColor Yellow
    Remove-Item $organizedDir -Recurse -Force
}

New-Item -ItemType Directory -Path $organizedDir -Force | Out-Null

# Step 3: Process each album folder
Write-Host "Processing album folders..." -ForegroundColor Yellow

$albumFolders = Get-ChildItem -Path $baseDir -Directory | Where-Object { $_.Name -notlike "*_Organized*" }

foreach ($folder in $albumFolders) {
    Write-Host "Processing: $($folder.Name)" -ForegroundColor Cyan
    
    $artist = Get-Artist $folder.Name
    $album = Get-Album $folder.Name
    $year = Get-Year $folder.Name
    $genre = Get-Genre $artist
    
    # Sanitize names
    $artist = Sanitize-Name $artist
    $album = Sanitize-Name $album
    $genre = Sanitize-Name $genre
    
    # Create organized folder structure: Genre\Artist\Year - Album
    $genreDir = Join-Path $organizedDir $genre
    $artistDir = Join-Path $genreDir $artist
    $albumDir = Join-Path $artistDir "$year - $album"
    
    # Create directories if they don't exist
    if (!(Test-Path $genreDir)) { New-Item -ItemType Directory -Path $genreDir -Force | Out-Null }
    if (!(Test-Path $artistDir)) { New-Item -ItemType Directory -Path $artistDir -Force | Out-Null }
    if (!(Test-Path $albumDir)) { New-Item -ItemType Directory -Path $albumDir -Force | Out-Null }
    
    # Copy music files and artwork
    $musicFiles = Get-ChildItem -Path $folder.FullName -File | Where-Object { 
        $_.Extension -in @('.mp3', '.flac', '.wav', '.m4a', '.aac', '.ogg', '.jpg', '.jpeg', '.png', '.bmp', '.gif') -and
        $_.Name -notlike '*.m3u*' -and
        $_.Name -notlike '*.nfo*' -and
        $_.Name -notlike '*.txt*'
    }
    
    foreach ($file in $musicFiles) {
        $destPath = Join-Path $albumDir $file.Name
        Copy-Item $file.FullName $destPath -Force
    }
    
    Write-Host "  -> Organized to: $albumDir" -ForegroundColor Green
}

# Step 4: Create new organized playlists by genre
Write-Host "Creating organized playlists by genre..." -ForegroundColor Yellow

$genres = Get-ChildItem -Path $organizedDir -Directory
foreach ($genreFolder in $genres) {
    $playlistPath = Join-Path $organizedDir "$($genreFolder.Name) - Complete Collection.m3u"
    $musicFiles = Get-ChildItem -Path $genreFolder.FullName -File -Recurse | Where-Object { 
        $_.Extension -in @('.mp3', '.flac', '.wav', '.m4a', '.aac', '.ogg') 
    }
    
    $playlistContent = "#EXTM3U`n"
    foreach ($musicFile in $musicFiles) {
        $relativePath = $musicFile.FullName.Replace($organizedDir + "\", "")
        $playlistContent += "$relativePath`n"
    }
    
    Set-Content -Path $playlistPath -Value $playlistContent -Encoding UTF8
    Write-Host "Created playlist: $($genreFolder.Name) - Complete Collection.m3u" -ForegroundColor Green
}

# Step 5: Create artist-specific playlists
Write-Host "Creating artist-specific playlists..." -ForegroundColor Yellow

foreach ($genreFolder in $genres) {
    $artists = Get-ChildItem -Path $genreFolder.FullName -Directory
    foreach ($artistFolder in $artists) {
        $playlistPath = Join-Path $artistFolder.FullName "$($artistFolder.Name) - Complete Discography.m3u"
        $musicFiles = Get-ChildItem -Path $artistFolder.FullName -File -Recurse | Where-Object { 
            $_.Extension -in @('.mp3', '.flac', '.wav', '.m4a', '.aac', '.ogg') 
        }
        
        $playlistContent = "#EXTM3U`n"
        foreach ($musicFile in $musicFiles) {
            $relativePath = $musicFile.Name
            $playlistContent += "$relativePath`n"
        }
        
        Set-Content -Path $playlistPath -Value $playlistContent -Encoding UTF8
        Write-Host "Created artist playlist: $($artistFolder.Name) - Complete Discography.m3u" -ForegroundColor Green
    }
}

Write-Host "`nMusic Library Organization Complete!" -ForegroundColor Green
Write-Host "Organized music is located in: $organizedDir" -ForegroundColor Cyan
Write-Host "Original files remain in: $baseDir" -ForegroundColor Cyan
Write-Host "`nNew structure: Genre\Artist\Year - Album" -ForegroundColor Yellow
Write-Host "Playlists created: Genre collections and artist discographies" -ForegroundColor Yellow

# Display summary
$totalGenres = (Get-ChildItem -Path $organizedDir -Directory).Count
$totalArtists = (Get-ChildItem -Path $organizedDir -Directory -Recurse | Where-Object { $_.Parent.Name -ne (Split-Path $organizedDir -Leaf) }).Count
$totalAlbums = (Get-ChildItem -Path $organizedDir -Directory -Recurse | Where-Object { $_.GetFiles() -ne $null }).Count
$totalTracks = (Get-ChildItem -Path $organizedDir -File -Recurse | Where-Object { $_.Extension -in @('.mp3', '.flac', '.wav', '.m4a', '.aac', '.ogg') }).Count

Write-Host "`n=== ORGANIZATION SUMMARY ===" -ForegroundColor Magenta
Write-Host "Genres: $totalGenres" -ForegroundColor White
Write-Host "Artists: $totalArtists" -ForegroundColor White
Write-Host "Albums: $totalAlbums" -ForegroundColor White
Write-Host "Tracks: $totalTracks" -ForegroundColor White
Write-Host "============================" -ForegroundColor Magenta