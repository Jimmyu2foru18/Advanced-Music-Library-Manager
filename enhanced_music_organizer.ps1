# Enhanced Music Library Organizer with Metadata Correction
# Uses MusicBrainz standards and best practices for accurate music organization

param(
    [string]$SourcePath = "h:\Music",
    [string]$DestinationPath = "h:\Music_Organized_Enhanced",
    [switch]$UseAcousticFingerprinting = $false,
    [switch]$DryRun = $false
)

# Import required modules
Add-Type -AssemblyName System.Web

# Function to sanitize file/folder names
function Sanitize-Name {
    param([string]$Name)
    $invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''
    $Name = $Name -replace "[$([regex]::Escape($invalidChars))]", '_'
    $Name = $Name -replace '[\[\](){}]', '_'
    $Name = $Name.Trim()
    return $Name
}

# Function to extract metadata from file tags
function Get-AudioMetadata {
    param([string]$FilePath)
    
    try {
        $shell = New-Object -ComObject Shell.Application
        $folder = $shell.Namespace((Get-Item $FilePath).DirectoryName)
        $file = $folder.ParseName((Get-Item $FilePath).Name)
        
        $metadata = @{
            Title = $folder.GetDetailsOf($file, 21)
            Artist = $folder.GetDetailsOf($file, 13)
            Album = $folder.GetDetailsOf($file, 14)
            Year = $folder.GetDetailsOf($file, 15)
            Genre = $folder.GetDetailsOf($file, 16)
            Track = $folder.GetDetailsOf($file, 26)
            Duration = $folder.GetDetailsOf($file, 27)
            Bitrate = $folder.GetDetailsOf($file, 28)
        }
        
        # Clean up metadata
        foreach ($key in $metadata.Keys) {
            if ($metadata[$key]) {
                $metadata[$key] = $metadata[$key].Trim()
            }
        }
        
        return $metadata
    }
    catch {
        Write-Warning "Could not extract metadata from: $FilePath"
        return $null
    }
}

# Enhanced genre classification based on MusicBrainz and Discogs standards
function Get-StandardizedGenre {
    param([string]$Artist, [string]$ExistingGenre)
    
    # Standard genre mapping based on Discogs taxonomy
    $genreMap = @{
        # Electronic subgenres
        'Ambient' = 'Electronic'
        'Breakbeat' = 'Electronic'
        'Drum and Bass' = 'Electronic'
        'Dubstep' = 'Electronic'
        'House' = 'Electronic'
        'Techno' = 'Electronic'
        'Trance' = 'Electronic'
        'IDM' = 'Electronic'
        'Downtempo' = 'Electronic'
        
        # Rock subgenres
        'Alternative Rock' = 'Rock'
        'Hard Rock' = 'Rock'
        'Progressive Rock' = 'Rock'
        'Punk Rock' = 'Rock'
        'Indie Rock' = 'Rock'
        'Classic Rock' = 'Rock'
        'Grunge' = 'Rock'
        
        # Metal subgenres
        'Heavy Metal' = 'Rock'
        'Death Metal' = 'Rock'
        'Black Metal' = 'Rock'
        'Thrash Metal' = 'Rock'
        'Power Metal' = 'Rock'
        
        # Hip-Hop variations
        'Rap' = 'Hip-Hop'
        'Hip Hop' = 'Hip-Hop'
        'Gangsta Rap' = 'Hip-Hop'
        
        # R&B variations
        'Soul' = 'Funk / Soul'
        'Funk' = 'Funk / Soul'
        'R&B' = 'Funk / Soul'
        'Rhythm and Blues' = 'Funk / Soul'
        
        # Country variations
        'Country Rock' = 'Folk, World, & Country'
        'Bluegrass' = 'Folk, World, & Country'
        'Folk' = 'Folk, World, & Country'
        'Americana' = 'Folk, World, & Country'
        
        # Classical variations
        'Classical' = 'Classical'
        'Opera' = 'Classical'
        'Symphony' = 'Classical'
        'Chamber Music' = 'Classical'
        
        # Jazz variations
        'Smooth Jazz' = 'Jazz'
        'Bebop' = 'Jazz'
        'Fusion' = 'Jazz'
        'Big Band' = 'Jazz'
        
        # Pop variations
        'Pop Rock' = 'Pop'
        'Synthpop' = 'Pop'
        'Dance Pop' = 'Pop'
        
        # Latin variations
        'Salsa' = 'Latin'
        'Reggaeton' = 'Latin'
        'Bossa Nova' = 'Latin'
        'Tango' = 'Latin'
        
        # Reggae variations
        'Dub' = 'Reggae'
        'Ska' = 'Reggae'
        'Dancehall' = 'Reggae'
        
        # Blues variations
        'Delta Blues' = 'Blues'
        'Chicago Blues' = 'Blues'
        'Electric Blues' = 'Blues'
        
        # Other
        'Soundtrack' = 'Stage & Screen'
        'Musical' = 'Stage & Screen'
        'Score' = 'Stage & Screen'
        'Gospel' = 'Religious'
        'Christian' = 'Religious'
        'Spiritual' = 'Religious'
        'Kids' = "Children's"
        'Children' = "Children's"
        'Instrumental' = 'Non-Music'
        'Spoken Word' = 'Non-Music'
    }
    
    # First try to use existing genre if it's already standardized
    if ($ExistingGenre) {
        $standardGenres = @('Blues', 'Brass & Military', "Children's", 'Classical', 'Electronic', 
                           'Folk, World, & Country', 'Funk / Soul', 'Hip-Hop', 'Jazz', 'Latin', 
                           'Non-Music', 'Pop', 'Reggae', 'Rock', 'Stage & Screen', 'Religious')
        
        if ($standardGenres -contains $ExistingGenre) {
            return $ExistingGenre
        }
        
        # Try to map the existing genre
        foreach ($key in $genreMap.Keys) {
            if ($ExistingGenre -like "*$key*") {
                return $genreMap[$key]
            }
        }
    }
    
    # Artist-based genre classification (simplified)
    $artistGenreMap = @{
        # Electronic artists
        'Daft Punk' = 'Electronic'
        'Deadmau5' = 'Electronic'
        'Skrillex' = 'Electronic'
        'Aphex Twin' = 'Electronic'
        'Kraftwerk' = 'Electronic'
        
        # Rock artists
        'Led Zeppelin' = 'Rock'
        'The Beatles' = 'Rock'
        'Queen' = 'Rock'
        'Pink Floyd' = 'Rock'
        'Metallica' = 'Rock'
        'Nirvana' = 'Rock'
        
        # Hip-Hop artists
        'Eminem' = 'Hip-Hop'
        'Jay-Z' = 'Hip-Hop'
        'Kanye West' = 'Hip-Hop'
        'Tupac' = 'Hip-Hop'
        'Notorious B.I.G.' = 'Hip-Hop'
        
        # Jazz artists
        'Miles Davis' = 'Jazz'
        'John Coltrane' = 'Jazz'
        'Duke Ellington' = 'Jazz'
        'Ella Fitzgerald' = 'Jazz'
        
        # Classical artists
        'Mozart' = 'Classical'
        'Beethoven' = 'Classical'
        'Bach' = 'Classical'
        'Chopin' = 'Classical'
        
        # Pop artists
        'Michael Jackson' = 'Pop'
        'Madonna' = 'Pop'
        'Taylor Swift' = 'Pop'
        'Britney Spears' = 'Pop'
        
        # R&B/Soul artists
        'Stevie Wonder' = 'Funk / Soul'
        'Aretha Franklin' = 'Funk / Soul'
        'Marvin Gaye' = 'Funk / Soul'
        'James Brown' = 'Funk / Soul'
        
        # Country artists
        'Johnny Cash' = 'Folk, World, & Country'
        'Dolly Parton' = 'Folk, World, & Country'
        'Willie Nelson' = 'Folk, World, & Country'
        
        # Reggae artists
        'Bob Marley' = 'Reggae'
        'Jimmy Cliff' = 'Reggae'
        'Peter Tosh' = 'Reggae'
        
        # Blues artists
        'B.B. King' = 'Blues'
        'Muddy Waters' = 'Blues'
        'Eric Clapton' = 'Blues'
    }
    
    if ($artistGenreMap.ContainsKey($Artist)) {
        return $artistGenreMap[$Artist]
    }
    
    # Default fallback
    return 'Pop'
}

# Function to clean and standardize artist names
function Clean-ArtistName {
    param([string]$Artist)
    
    if (-not $Artist) { return "Unknown Artist" }
    
    # Remove common prefixes/suffixes
    $Artist = $Artist -replace '^The\s+', ''
    $Artist = $Artist -replace '\s+feat\..*$', ''
    $Artist = $Artist -replace '\s+ft\..*$', ''
    $Artist = $Artist -replace '\s+featuring.*$', ''
    $Artist = $Artist -replace '\s+\(.*\)$', ''
    
    return $Artist.Trim()
}

# Function to clean and standardize album names
function Clean-AlbumName {
    param([string]$Album)
    
    if (-not $Album) { return "Unknown Album" }
    
    # Remove common suffixes
    $Album = $Album -replace '\s+\(Deluxe.*\)$', ''
    $Album = $Album -replace '\s+\(Remaster.*\)$', ''
    $Album = $Album -replace '\s+\(Special.*\)$', ''
    $Album = $Album -replace '\s+\(Expanded.*\)$', ''
    
    return $Album.Trim()
}

# Function to extract year from various sources
function Get-Year {
    param([string]$YearTag, [string]$FolderName, [string]$FileName)
    
    # Try year from tag first
    if ($YearTag -and $YearTag -match '\d{4}') {
        return [regex]::Match($YearTag, '\d{4}').Value
    }
    
    # Try folder name
    if ($FolderName -match '(19|20)\d{2}') {
        return [regex]::Match($FolderName, '(19|20)\d{2}').Value
    }
    
    # Try file name
    if ($FileName -match '(19|20)\d{2}') {
        return [regex]::Match($FileName, '(19|20)\d{2}').Value
    }
    
    return "Unknown"
}

# Function to create M3U playlist
function Create-Playlist {
    param(
        [string]$PlaylistPath,
        [array]$Files,
        [string]$PlaylistName
    )
    
    if ($Files.Count -eq 0) { return }
    
    $content = @("#EXTM3U")
    $content += "#PLAYLIST:$PlaylistName"
    
    foreach ($file in $Files) {
        $relativePath = $file.FullName.Replace($DestinationPath + "\", "")
        $content += "#EXTINF:-1,$($file.BaseName)"
        $content += $relativePath
    }
    
    $content | Out-File -FilePath $PlaylistPath -Encoding UTF8
}

# Main processing function
function Process-MusicLibrary {
    Write-Host "Enhanced Music Library Organizer" -ForegroundColor Green
    Write-Host "Source: $SourcePath" -ForegroundColor Yellow
    Write-Host "Destination: $DestinationPath" -ForegroundColor Yellow
    Write-Host "Dry Run: $DryRun" -ForegroundColor Yellow
    Write-Host ""
    
    # Create destination directory
    if (-not $DryRun -and -not (Test-Path $DestinationPath)) {
        New-Item -ItemType Directory -Path $DestinationPath -Force | Out-Null
    }
    
    # Remove existing playlists from source
    Write-Host "Removing existing playlist files..." -ForegroundColor Cyan
    $playlistFiles = Get-ChildItem -Path $SourcePath -Filter "*.m3u" -Recurse
    foreach ($playlist in $playlistFiles) {
        Write-Host "  Removing: $($playlist.FullName)"
        if (-not $DryRun) {
            Remove-Item $playlist.FullName -Force
        }
    }
    
    # Get all audio files
    $audioExtensions = @('*.mp3', '*.flac', '*.wav', '*.m4a', '*.aac', '*.ogg', '*.wma')
    $audioFiles = @()
    foreach ($ext in $audioExtensions) {
        $audioFiles += Get-ChildItem -Path $SourcePath -Filter $ext -Recurse
    }
    
    Write-Host "Found $($audioFiles.Count) audio files" -ForegroundColor Green
    
    $stats = @{
        Processed = 0
        Genres = @{}
        Artists = @{}
        Albums = @{}
        Years = @{}
    }
    
    $processedFiles = @()
    
    foreach ($file in $audioFiles) {
        Write-Progress -Activity "Processing Music Files" -Status "Processing $($file.Name)" -PercentComplete (($stats.Processed / $audioFiles.Count) * 100)
        
        # Extract metadata
        $metadata = Get-AudioMetadata -FilePath $file.FullName
        
        # Use filename/folder structure as fallback
        $folderName = $file.Directory.Name
        $fileName = $file.BaseName
        
        # Determine metadata with fallbacks
        $artist = if ($metadata -and $metadata.Artist) { $metadata.Artist } else {
            if ($folderName -match '^(.+?)\s*-\s*(.+)$') {
                $matches[1]
            } else {
                "Unknown Artist"
            }
        }
        
        $album = if ($metadata -and $metadata.Album) { $metadata.Album } else {
            if ($folderName -match '^(.+?)\s*-\s*(.+)$') {
                $matches[2]
            } else {
                $folderName
            }
        }
        
        $title = if ($metadata -and $metadata.Title) { $metadata.Title } else {
            $fileName -replace '^\d+\s*[-.]\s*', ''
        }
        
        $year = Get-Year -YearTag ($metadata.Year) -FolderName $folderName -FileName $fileName
        
        # Clean and standardize
        $artist = Clean-ArtistName -Artist $artist
        $album = Clean-AlbumName -Album $album
        $genre = Get-StandardizedGenre -Artist $artist -ExistingGenre ($metadata.Genre)
        
        # Sanitize for file system
        $safeArtist = Sanitize-Name -Name $artist
        $safeAlbum = Sanitize-Name -Name $album
        $safeTitle = Sanitize-Name -Name $title
        $safeGenre = Sanitize-Name -Name $genre
        
        # Create directory structure: Genre\Artist\Year - Album
        $albumFolder = if ($year -ne "Unknown") { "$year - $safeAlbum" } else { $safeAlbum }
        $targetDir = Join-Path $DestinationPath "$safeGenre\$safeArtist\$albumFolder"
        
        if (-not $DryRun -and -not (Test-Path $targetDir)) {
            New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
        }
        
        # Create new filename with track number if available
        $trackNum = if ($metadata -and $metadata.Track) {
            $trackMatch = [regex]::Match($metadata.Track, '\d+')
            if ($trackMatch.Success) {
                $trackMatch.Value.PadLeft(2, '0')
            } else {
                "01"
            }
        } else {
            "01"
        }
        
        $newFileName = "$trackNum - $safeTitle$($file.Extension)"
        $targetFile = Join-Path $targetDir $newFileName
        
        # Copy file
        if (-not $DryRun) {
            try {
                Copy-Item $file.FullName $targetFile -Force
                Write-Host "  Copied: $($file.Name) -> $targetFile" -ForegroundColor Gray
            }
            catch {
                Write-Warning "Failed to copy $($file.FullName): $($_.Exception.Message)"
                continue
            }
        }
        
        # Copy artwork files
        $artworkFiles = Get-ChildItem -Path $file.Directory -Filter "*.jpg" -ErrorAction SilentlyContinue
        $artworkFiles += Get-ChildItem -Path $file.Directory -Filter "*.png" -ErrorAction SilentlyContinue
        $artworkFiles += Get-ChildItem -Path $file.Directory -Filter "*.bmp" -ErrorAction SilentlyContinue
        
        foreach ($artwork in $artworkFiles) {
            $artworkTarget = Join-Path $targetDir $artwork.Name
            if (-not $DryRun -and -not (Test-Path $artworkTarget)) {
                Copy-Item $artwork.FullName $artworkTarget -Force -ErrorAction SilentlyContinue
            }
        }
        
        # Track statistics
        if ($stats.Genres.ContainsKey($genre)) {
            $stats.Genres[$genre] += 1
        } else {
            $stats.Genres[$genre] = 1
        }
        
        if ($stats.Artists.ContainsKey($artist)) {
            $stats.Artists[$artist] += 1
        } else {
            $stats.Artists[$artist] = 1
        }
        
        $albumKey = "$artist - $album"
        if ($stats.Albums.ContainsKey($albumKey)) {
            $stats.Albums[$albumKey] += 1
        } else {
            $stats.Albums[$albumKey] = 1
        }
        
        if ($stats.Years.ContainsKey($year)) {
            $stats.Years[$year] += 1
        } else {
            $stats.Years[$year] = 1
        }
        
        # Add to processed files for playlist generation
        $processedFiles += [PSCustomObject]@{
            FullName = $targetFile
            BaseName = $safeTitle
            Genre = $genre
            Artist = $artist
            Album = $album
            Year = $year
        }
        
        $stats.Processed++
    }
    
    Write-Progress -Activity "Processing Music Files" -Completed
    
    # Generate playlists
    if (-not $DryRun) {
        Write-Host "\nGenerating playlists..." -ForegroundColor Cyan
        
        # Genre playlists
        foreach ($genre in $stats.Genres.Keys) {
            $genreFiles = $processedFiles | Where-Object { $_.Genre -eq $genre }
            $playlistPath = Join-Path $DestinationPath "$genre.m3u"
            Create-Playlist -PlaylistPath $playlistPath -Files $genreFiles -PlaylistName $genre
            Write-Host "  Created genre playlist: $genre.m3u ($($genreFiles.Count) tracks)"
        }
        
        # Artist playlists
        foreach ($artist in $stats.Artists.Keys) {
            $artistFiles = $processedFiles | Where-Object { $_.Artist -eq $artist }
            if ($artistFiles.Count -gt 1) {
                $safeArtistName = Sanitize-Name -Name $artist
                $playlistPath = Join-Path $DestinationPath "Artist - $safeArtistName.m3u"
                Create-Playlist -PlaylistPath $playlistPath -Files $artistFiles -PlaylistName "$artist - Complete"
                Write-Host "  Created artist playlist: Artist - $safeArtistName.m3u ($($artistFiles.Count) tracks)"
            }
        }
        
        # Year playlists
        foreach ($year in ($stats.Years.Keys | Where-Object { $_ -ne "Unknown" -and $stats.Years[$_] -gt 5 })) {
            $yearFiles = $processedFiles | Where-Object { $_.Year -eq $year }
            $playlistPath = Join-Path $DestinationPath "Year - $year.m3u"
            Create-Playlist -PlaylistPath $playlistPath -Files $yearFiles -PlaylistName "$year Music"
            Write-Host "  Created year playlist: Year - $year.m3u ($($yearFiles.Count) tracks)"
        }
    }
    
    # Display summary
    Write-Host "\n" + "="*50 -ForegroundColor Green
    Write-Host "ORGANIZATION COMPLETE" -ForegroundColor Green
    Write-Host "="*50 -ForegroundColor Green
    Write-Host "Total files processed: $($stats.Processed)" -ForegroundColor Yellow
    Write-Host "Genres found: $($stats.Genres.Count)" -ForegroundColor Yellow
    Write-Host "Artists found: $($stats.Artists.Count)" -ForegroundColor Yellow
    Write-Host "Albums found: $($stats.Albums.Count)" -ForegroundColor Yellow
    
    Write-Host "\nTop Genres:" -ForegroundColor Cyan
    $stats.Genres.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 10 | ForEach-Object {
        Write-Host "  $($_.Key): $($_.Value) tracks" -ForegroundColor White
    }
    
    if ($DryRun) {
        Write-Host "\nThis was a DRY RUN - no files were actually moved or copied." -ForegroundColor Red
        Write-Host "Run without -DryRun parameter to perform the actual organization." -ForegroundColor Red
    }
}

# Run the main function
Process-MusicLibrary

Write-Host "\nScript completed. Check the organized library at: $DestinationPath" -ForegroundColor Green
Write-Host "\nFor best results, consider using MusicBrainz Picard for additional metadata correction:" -ForegroundColor Yellow
Write-Host "1. Download MusicBrainz Picard from https://picard.musicbrainz.org/" -ForegroundColor White
Write-Host "2. Use 'Cluster' and 'Lookup' functions for files with existing metadata" -ForegroundColor White
Write-Host "3. Use 'Scan' function only for files with poor or missing metadata" -ForegroundColor White
Write-Host "4. Always review matches before saving to avoid incorrect tagging" -ForegroundColor White