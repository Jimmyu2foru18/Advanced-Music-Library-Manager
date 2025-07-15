# Comprehensive Music Metadata Correction and Organization Script
# Based on MusicBrainz standards and best practices for accurate music library management

param(
    [string]$SourcePath = "h:\Music",
    [string]$DestinationPath = "h:\Music_Fixed",
    [switch]$UsePicard = $false,
    [switch]$DryRun = $false,
    [switch]$DownloadPicard = $false
)

# Function to download and install MusicBrainz Picard
function Install-MusicBrainzPicard {
    Write-Host "Downloading MusicBrainz Picard..." -ForegroundColor Cyan
    
    $picardUrl = "https://github.com/metabrainz/picard/releases/latest/download/MusicBrainz-Picard-2.10-win-x86_64.exe"
    $picardInstaller = "$env:TEMP\MusicBrainz-Picard-Installer.exe"
    
    try {
        Invoke-WebRequest -Uri $picardUrl -OutFile $picardInstaller -UseBasicParsing
        Write-Host "Starting Picard installation..." -ForegroundColor Green
        Start-Process -FilePath $picardInstaller -ArgumentList "/S" -Wait
        Write-Host "MusicBrainz Picard installed successfully!" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Warning "Failed to download/install MusicBrainz Picard: $($_.Exception.Message)"
        return $false
    }
}

# Function to sanitize file/folder names
function Sanitize-Name {
    param([string]$Name)
    if (-not $Name) { return "Unknown" }
    
    $invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''
    $Name = $Name -replace "[$([regex]::Escape($invalidChars))]", '_'
    $Name = $Name -replace '[\[\](){}]', '_'
    $Name = $Name -replace '\s+', ' '
    $Name = $Name.Trim()
    
    # Remove common problematic patterns
    $Name = $Name -replace '\.$', ''
    $Name = $Name -replace '^\s*-\s*', ''
    $Name = $Name -replace '\s*-\s*$', ''
    
    return $Name
}

# Enhanced metadata extraction using multiple methods
function Get-EnhancedMetadata {
    param([string]$FilePath)
    
    $metadata = @{
        Title = $null
        Artist = $null
        Album = $null
        Year = $null
        Genre = $null
        Track = $null
        AlbumArtist = $null
        Duration = $null
    }
    
    try {
        # Method 1: Shell.Application (Windows built-in)
        $shell = New-Object -ComObject Shell.Application
        $folder = $shell.Namespace((Get-Item $FilePath).DirectoryName)
        $file = $folder.ParseName((Get-Item $FilePath).Name)
        
        $shellMetadata = @{
            Title = $folder.GetDetailsOf($file, 21)
            Artist = $folder.GetDetailsOf($file, 13)
            Album = $folder.GetDetailsOf($file, 14)
            Year = $folder.GetDetailsOf($file, 15)
            Genre = $folder.GetDetailsOf($file, 16)
            Track = $folder.GetDetailsOf($file, 26)
            Duration = $folder.GetDetailsOf($file, 27)
        }
        
        foreach ($key in $shellMetadata.Keys) {
            if ($shellMetadata[$key] -and $shellMetadata[$key].Trim()) {
                $metadata[$key] = $shellMetadata[$key].Trim()
            }
        }
        
        # Method 2: Try TagLib# if available (more accurate)
        if (Test-Path "$env:ProgramFiles\TagLib\TagLibSharp.dll") {
            Add-Type -Path "$env:ProgramFiles\TagLib\TagLibSharp.dll"
            $tagFile = [TagLib.File]::Create($FilePath)
            
            if ($tagFile.Tag) {
                if ($tagFile.Tag.Title) { $metadata.Title = $tagFile.Tag.Title }
                if ($tagFile.Tag.FirstPerformer) { $metadata.Artist = $tagFile.Tag.FirstPerformer }
                if ($tagFile.Tag.Album) { $metadata.Album = $tagFile.Tag.Album }
                if ($tagFile.Tag.Year -gt 0) { $metadata.Year = $tagFile.Tag.Year.ToString() }
                if ($tagFile.Tag.FirstGenre) { $metadata.Genre = $tagFile.Tag.FirstGenre }
                if ($tagFile.Tag.Track -gt 0) { $metadata.Track = $tagFile.Tag.Track.ToString() }
                if ($tagFile.Tag.FirstAlbumArtist) { $metadata.AlbumArtist = $tagFile.Tag.FirstAlbumArtist }
            }
            
            $tagFile.Dispose()
        }
        
    }
    catch {
        Write-Verbose "Metadata extraction failed for ${FilePath}: $($_.Exception.Message)"
    }
    
    return $metadata
}

# Advanced genre classification based on multiple sources
function Get-ImprovedGenre {
    param(
        [string]$Artist, 
        [string]$ExistingGenre, 
        [string]$Album,
        [string]$FolderPath
    )
    
    # Standard MusicBrainz/Discogs genre taxonomy
    $standardGenres = @(
        'Blues', 'Brass & Military', "Children's", 'Classical', 'Electronic', 
        'Folk, World, & Country', 'Funk / Soul', 'Hip-Hop', 'Jazz', 'Latin', 
        'Non-Music', 'Pop', 'Reggae', 'Rock', 'Stage & Screen', 'Religious'
    )
    
    # If existing genre is already standard, use it
    if ($ExistingGenre -and $standardGenres -contains $ExistingGenre) {
        return $ExistingGenre
    }
    
    # Genre keyword mapping
    $genreKeywords = @{
        'Electronic' = @('electronic', 'techno', 'house', 'trance', 'ambient', 'dubstep', 'edm', 'synth', 'electro', 'dance', 'club', 'rave', 'breakbeat', 'drum and bass', 'dnb', 'jungle', 'garage', 'hardcore', 'hardstyle', 'minimal', 'progressive', 'psytrance', 'chillout', 'downtempo', 'trip hop', 'idm', 'glitch')
        'Rock' = @('rock', 'metal', 'punk', 'grunge', 'alternative', 'indie', 'hard rock', 'classic rock', 'progressive rock', 'psychedelic', 'garage rock', 'post rock', 'math rock', 'stoner', 'doom', 'sludge', 'black metal', 'death metal', 'thrash', 'heavy metal', 'power metal', 'gothic', 'industrial', 'nu metal', 'metalcore', 'hardcore', 'post hardcore', 'emo', 'screamo')
        'Hip-Hop' = @('hip hop', 'rap', 'hip-hop', 'gangsta', 'trap', 'drill', 'grime', 'boom bap', 'conscious', 'underground', 'old school', 'east coast', 'west coast', 'southern', 'crunk', 'mumble')
        'Jazz' = @('jazz', 'bebop', 'swing', 'fusion', 'smooth jazz', 'free jazz', 'cool jazz', 'hard bop', 'post bop', 'avant garde jazz', 'latin jazz', 'acid jazz', 'nu jazz', 'contemporary jazz')
        'Classical' = @('classical', 'orchestra', 'symphony', 'concerto', 'sonata', 'chamber', 'baroque', 'romantic', 'modern classical', 'contemporary classical', 'minimalist', 'opera', 'choral', 'string quartet')
        'Funk / Soul' = @('funk', 'soul', 'r&b', 'rhythm and blues', 'motown', 'neo soul', 'contemporary r&b', 'new jack swing', 'quiet storm', 'northern soul', 'southern soul', 'deep funk', 'p-funk')
        'Folk, World, & Country' = @('folk', 'country', 'bluegrass', 'americana', 'world', 'traditional', 'celtic', 'irish', 'scottish', 'african', 'indian', 'asian', 'middle eastern', 'latin american', 'european', 'acoustic', 'singer songwriter', 'indie folk', 'alt country', 'outlaw country', 'honky tonk')
        'Pop' = @('pop', 'pop rock', 'dance pop', 'electropop', 'synthpop', 'teen pop', 'bubblegum', 'adult contemporary', 'soft rock', 'easy listening', 'lounge', 'yacht rock')
        'Reggae' = @('reggae', 'ska', 'dub', 'dancehall', 'roots reggae', 'lovers rock', 'ragga', 'reggaeton', 'rocksteady')
        'Blues' = @('blues', 'delta blues', 'chicago blues', 'electric blues', 'acoustic blues', 'country blues', 'rhythm and blues', 'blues rock', 'boogie', 'jump blues')
        'Latin' = @('latin', 'salsa', 'merengue', 'bachata', 'cumbia', 'tango', 'bossa nova', 'samba', 'mambo', 'cha cha', 'rumba', 'flamenco', 'mariachi', 'ranchera', 'tejano', 'norteño')
        'Stage & Screen' = @('soundtrack', 'score', 'musical', 'broadway', 'film', 'movie', 'tv', 'television', 'game', 'video game', 'anime', 'theme')
        'Religious' = @('gospel', 'christian', 'spiritual', 'hymn', 'praise', 'worship', 'contemporary christian', 'southern gospel', 'black gospel', 'country gospel', 'religious', 'sacred')
        "Children's" = @('children', 'kids', 'nursery', 'lullaby', 'educational', 'family')
        'Non-Music' = @('spoken word', 'audiobook', 'comedy', 'interview', 'speech', 'poetry', 'meditation', 'nature sounds', 'white noise')
    }
    
    # Check existing genre against keywords
    if ($ExistingGenre) {
        $genreLower = $ExistingGenre.ToLower()
        foreach ($genre in $genreKeywords.Keys) {
            foreach ($keyword in $genreKeywords[$genre]) {
                if ($genreLower -like "*$keyword*") {
                    return $genre
                }
            }
        }
    }
    
    # Check artist name against keywords
    if ($Artist) {
        $artistLower = $Artist.ToLower()
        foreach ($genre in $genreKeywords.Keys) {
            foreach ($keyword in $genreKeywords[$genre]) {
                if ($artistLower -like "*$keyword*") {
                    return $genre
                }
            }
        }
    }
    
    # Check album name against keywords
    if ($Album) {
        $albumLower = $Album.ToLower()
        foreach ($genre in $genreKeywords.Keys) {
            foreach ($keyword in $genreKeywords[$genre]) {
                if ($albumLower -like "*$keyword*") {
                    return $genre
                }
            }
        }
    }
    
    # Check folder path for genre hints
    if ($FolderPath) {
        $folderLower = $FolderPath.ToLower()
        foreach ($genre in $genreKeywords.Keys) {
            foreach ($keyword in $genreKeywords[$genre]) {
                if ($folderLower -like "*$keyword*") {
                    return $genre
                }
            }
        }
    }
    
    # Comprehensive artist-to-genre mapping
    $artistGenreMap = @{
        # Electronic
        'Daft Punk' = 'Electronic'; 'Deadmau5' = 'Electronic'; 'Skrillex' = 'Electronic'
        'Aphex Twin' = 'Electronic'; 'Kraftwerk' = 'Electronic'; 'Moby' = 'Electronic'
        'The Chemical Brothers' = 'Electronic'; 'Fatboy Slim' = 'Electronic'
        'Underworld' = 'Electronic'; 'Prodigy' = 'Electronic'; 'Massive Attack' = 'Electronic'
        
        # Rock
        'Led Zeppelin' = 'Rock'; 'The Beatles' = 'Rock'; 'Queen' = 'Rock'
        'Pink Floyd' = 'Rock'; 'Metallica' = 'Rock'; 'Nirvana' = 'Rock'
        'AC/DC' = 'Rock'; 'Black Sabbath' = 'Rock'; 'Deep Purple' = 'Rock'
        'The Rolling Stones' = 'Rock'; 'The Who' = 'Rock'; 'Guns N Roses' = 'Rock'
        'Iron Maiden' = 'Rock'; 'Judas Priest' = 'Rock'; 'Megadeth' = 'Rock'
        'Slayer' = 'Rock'; 'Anthrax' = 'Rock'; 'Pearl Jam' = 'Rock'
        'Soundgarden' = 'Rock'; 'Alice in Chains' = 'Rock'; 'Stone Temple Pilots' = 'Rock'
        'Radiohead' = 'Rock'; 'Foo Fighters' = 'Rock'; 'Red Hot Chili Peppers' = 'Rock'
        
        # Hip-Hop
        'Eminem' = 'Hip-Hop'; 'Jay-Z' = 'Hip-Hop'; 'Kanye West' = 'Hip-Hop'
        'Tupac' = 'Hip-Hop'; 'Notorious B.I.G.' = 'Hip-Hop'; 'Nas' = 'Hip-Hop'
        'Dr. Dre' = 'Hip-Hop'; 'Snoop Dogg' = 'Hip-Hop'; 'Ice Cube' = 'Hip-Hop'
        'Wu-Tang Clan' = 'Hip-Hop'; 'Public Enemy' = 'Hip-Hop'; 'N.W.A' = 'Hip-Hop'
        'Kendrick Lamar' = 'Hip-Hop'; 'Drake' = 'Hip-Hop'; 'J. Cole' = 'Hip-Hop'
        
        # Jazz
        'Miles Davis' = 'Jazz'; 'John Coltrane' = 'Jazz'; 'Duke Ellington' = 'Jazz'
        'Ella Fitzgerald' = 'Jazz'; 'Louis Armstrong' = 'Jazz'; 'Charlie Parker' = 'Jazz'
        'Thelonious Monk' = 'Jazz'; 'Bill Evans' = 'Jazz'; 'Chet Baker' = 'Jazz'
        'Herbie Hancock' = 'Jazz'; 'Weather Report' = 'Jazz'; 'Pat Metheny' = 'Jazz'
        
        # Classical
        'Mozart' = 'Classical'; 'Beethoven' = 'Classical'; 'Bach' = 'Classical'
        'Chopin' = 'Classical'; 'Tchaikovsky' = 'Classical'; 'Vivaldi' = 'Classical'
        'Brahms' = 'Classical'; 'Debussy' = 'Classical'; 'Stravinsky' = 'Classical'
        
        # Pop
        'Michael Jackson' = 'Pop'; 'Madonna' = 'Pop'; 'Taylor Swift' = 'Pop'
        'Britney Spears' = 'Pop'; 'Justin Timberlake' = 'Pop'; 'Ariana Grande' = 'Pop'
        'Ed Sheeran' = 'Pop'; 'Adele' = 'Pop'; 'Bruno Mars' = 'Pop'
        
        # Funk/Soul
        'Stevie Wonder' = 'Funk / Soul'; 'Aretha Franklin' = 'Funk / Soul'
        'Marvin Gaye' = 'Funk / Soul'; 'James Brown' = 'Funk / Soul'
        'Parliament-Funkadelic' = 'Funk / Soul'; 'Sly & The Family Stone' = 'Funk / Soul'
        'Earth Wind & Fire' = 'Funk / Soul'; 'The Temptations' = 'Funk / Soul'
        
        # Country/Folk
        'Johnny Cash' = 'Folk, World, & Country'; 'Dolly Parton' = 'Folk, World, & Country'
        'Willie Nelson' = 'Folk, World, & Country'; 'Hank Williams' = 'Folk, World, & Country'
        'Bob Dylan' = 'Folk, World, & Country'; 'Neil Young' = 'Folk, World, & Country'
        
        # Reggae
        'Bob Marley' = 'Reggae'; 'Jimmy Cliff' = 'Reggae'; 'Peter Tosh' = 'Reggae'
        'Burning Spear' = 'Reggae'; 'Black Uhuru' = 'Reggae'; 'Steel Pulse' = 'Reggae'
        
        # Blues
        'B.B. King' = 'Blues'; 'Muddy Waters' = 'Blues'; 'Eric Clapton' = 'Blues'
        'Stevie Ray Vaughan' = 'Blues'; 'Robert Johnson' = 'Blues'; 'Howlin Wolf' = 'Blues'
        
        # Latin
        'Shakira' = 'Latin'; 'Manu Chao' = 'Latin'; 'Gipsy Kings' = 'Latin'
        'Buena Vista Social Club' = 'Latin'; 'Santana' = 'Latin'
    }
    
    # Check artist mapping
    if ($Artist -and $artistGenreMap.ContainsKey($Artist)) {
        return $artistGenreMap[$Artist]
    }
    
    # Partial artist name matching
    if ($Artist) {
        foreach ($mappedArtist in $artistGenreMap.Keys) {
            if ($Artist -like "*$mappedArtist*" -or $mappedArtist -like "*$Artist*") {
                return $artistGenreMap[$mappedArtist]
            }
        }
    }
    
    # Default fallback
    return 'Pop'
}

# Function to extract and clean metadata with multiple fallback strategies
function Get-CleanedMetadata {
    param(
        [string]$FilePath,
        [object]$RawMetadata
    )
    
    $file = Get-Item $FilePath
    $folderName = $file.Directory.Name
    $fileName = $file.BaseName
    
    # Initialize cleaned metadata
    $cleaned = @{
        Title = $null
        Artist = $null
        Album = $null
        Year = $null
        Genre = $null
        Track = $null
    }
    
    # Strategy 1: Use existing metadata if good quality
    if ($RawMetadata.Title -and $RawMetadata.Title.Length -gt 2) {
        $cleaned.Title = $RawMetadata.Title
    }
    if ($RawMetadata.Artist -and $RawMetadata.Artist.Length -gt 1) {
        $cleaned.Artist = $RawMetadata.Artist
    }
    if ($RawMetadata.Album -and $RawMetadata.Album.Length -gt 1) {
        $cleaned.Album = $RawMetadata.Album
    }
    if ($RawMetadata.Year -and $RawMetadata.Year -match '\d{4}') {
        $cleaned.Year = [regex]::Match($RawMetadata.Year, '\d{4}').Value
    }
    if ($RawMetadata.Track -and $RawMetadata.Track -match '\d+') {
        $cleaned.Track = [regex]::Match($RawMetadata.Track, '\d+').Value
    }
    
    # Strategy 2: Parse folder structure
    # Common patterns: "Artist - Album", "Year - Album", "Artist - Year - Album"
    if (-not $cleaned.Artist -or -not $cleaned.Album) {
        if ($folderName -match '^(.+?)\s*-\s*(.+)$') {
            $part1 = $matches[1].Trim()
            $part2 = $matches[2].Trim()
            
            if ($part1 -match '^\d{4}$') {
                # Pattern: "Year - Album"
                if (-not $cleaned.Year) { $cleaned.Year = $part1 }
                if (-not $cleaned.Album) { $cleaned.Album = $part2 }
            } elseif ($part2 -match '^\d{4}') {
                # Pattern: "Artist - Year"
                if (-not $cleaned.Artist) { $cleaned.Artist = $part1 }
                if (-not $cleaned.Year) { $cleaned.Year = $part2 }
            } else {
                # Pattern: "Artist - Album"
                if (-not $cleaned.Artist) { $cleaned.Artist = $part1 }
                if (-not $cleaned.Album) { $cleaned.Album = $part2 }
            }
        }
        
        # Three-part pattern: "Artist - Year - Album"
        if ($folderName -match '^(.+?)\s*-\s*(\d{4})\s*-\s*(.+)$') {
            if (-not $cleaned.Artist) { $cleaned.Artist = $matches[1].Trim() }
            if (-not $cleaned.Year) { $cleaned.Year = $matches[2] }
            if (-not $cleaned.Album) { $cleaned.Album = $matches[3].Trim() }
        }
    }
    
    # Strategy 3: Parse filename
    if (-not $cleaned.Title) {
        # Remove track number prefix
        $titleFromFile = $fileName -replace '^\d+\s*[-.]\s*', ''
        # Remove common suffixes
        $titleFromFile = $titleFromFile -replace '\s*\(.*\)$', ''
        $titleFromFile = $titleFromFile -replace '\s*\[.*\]$', ''
        
        if ($titleFromFile.Length -gt 2) {
            $cleaned.Title = $titleFromFile
        }
    }
    
    # Strategy 4: Extract track number from filename
    if (-not $cleaned.Track -and $fileName -match '^(\d+)') {
        $cleaned.Track = $matches[1]
    }
    
    # Strategy 5: Extract year from various sources
    if (-not $cleaned.Year) {
        # Check parent folder for year
        $parentFolder = $file.Directory.Parent.Name
        if ($parentFolder -match '(19|20)\d{2}') {
            $cleaned.Year = [regex]::Match($parentFolder, '(19|20)\d{2}').Value
        }
        # Check filename for year
        elseif ($fileName -match '(19|20)\d{2}') {
            $cleaned.Year = [regex]::Match($fileName, '(19|20)\d{2}').Value
        }
    }
    
    # Strategy 6: Use folder name as album if still missing
    if (-not $cleaned.Album) {
        $cleaned.Album = $folderName
    }
    
    # Strategy 7: Use parent folder as artist if still missing
    if (-not $cleaned.Artist) {
        $cleaned.Artist = $file.Directory.Parent.Name
    }
    
    # Clean up all fields
    foreach ($key in $cleaned.Keys) {
        if ($cleaned[$key]) {
            $cleaned[$key] = $cleaned[$key].Trim()
            # Remove common artifacts
            $cleaned[$key] = $cleaned[$key] -replace '^\s*-\s*', ''
            $cleaned[$key] = $cleaned[$key] -replace '\s*-\s*$', ''
            $cleaned[$key] = $cleaned[$key] -replace '\s+', ' '
        }
    }
    
    # Set defaults for missing fields
    if (-not $cleaned.Title -or $cleaned.Title.Length -lt 2) { $cleaned.Title = $fileName }
    if (-not $cleaned.Artist -or $cleaned.Artist.Length -lt 2) { $cleaned.Artist = "Unknown Artist" }
    if (-not $cleaned.Album -or $cleaned.Album.Length -lt 2) { $cleaned.Album = "Unknown Album" }
    if (-not $cleaned.Year) { $cleaned.Year = "Unknown" }
    if (-not $cleaned.Track) { $cleaned.Track = "01" }
    
    # Determine genre
    $cleaned.Genre = Get-ImprovedGenre -Artist $cleaned.Artist -ExistingGenre $RawMetadata.Genre -Album $cleaned.Album -FolderPath $file.DirectoryName
    
    return $cleaned
}

# Function to create M3U playlist with proper encoding
function Create-EnhancedPlaylist {
    param(
        [string]$PlaylistPath,
        [array]$Files,
        [string]$PlaylistName
    )
    
    if ($Files.Count -eq 0) { return }
    
    $content = @("#EXTM3U")
    $content += "#PLAYLIST:$PlaylistName"
    $content += "#EXTENC:UTF-8"
    
    foreach ($file in $Files) {
        $relativePath = $file.FullName.Replace($DestinationPath + "\", "")
        $duration = if ($file.Duration) { $file.Duration } else { "-1" }
        $artist = if ($file.Artist) { $file.Artist } else { "Unknown" }
        $title = if ($file.Title) { $file.Title } else { $file.BaseName }
        
        $content += "#EXTINF:$duration,$artist - $title"
        $content += $relativePath
    }
    
    try {
        $content | Out-File -FilePath $PlaylistPath -Encoding UTF8 -Force
    }
    catch {
        Write-Warning "Failed to create playlist ${PlaylistPath}: $($_.Exception.Message)"
    }
}

# Main processing function
function Start-ComprehensiveMusicFix {
    Write-Host "=" * 60 -ForegroundColor Green
    Write-Host "COMPREHENSIVE MUSIC METADATA CORRECTION & ORGANIZATION" -ForegroundColor Green
    Write-Host "=" * 60 -ForegroundColor Green
    Write-Host "Source: $SourcePath" -ForegroundColor Yellow
    Write-Host "Destination: $DestinationPath" -ForegroundColor Yellow
    Write-Host "Dry Run: $DryRun" -ForegroundColor Yellow
    Write-Host "Use Picard: $UsePicard" -ForegroundColor Yellow
    Write-Host ""
    
    # Download Picard if requested
    if ($DownloadPicard) {
        $picardInstalled = Install-MusicBrainzPicard
        if (-not $picardInstalled) {
            Write-Host "Continuing without Picard..." -ForegroundColor Yellow
        }
    }
    
    # Create destination directory
    if (-not $DryRun -and -not (Test-Path $DestinationPath)) {
        New-Item -ItemType Directory -Path $DestinationPath -Force | Out-Null
    }
    
    # Remove existing playlists from source
    Write-Host "Cleaning up existing playlist files..." -ForegroundColor Cyan
    $playlistFiles = Get-ChildItem -Path $SourcePath -Filter "*.m3u" -Recurse -ErrorAction SilentlyContinue
    $playlistCount = 0
    foreach ($playlist in $playlistFiles) {
        Write-Host "  Removing: $($playlist.Name)" -ForegroundColor Gray
        if (-not $DryRun) {
            try {
                Remove-Item $playlist.FullName -Force
                $playlistCount++
            }
            catch {
                Write-Warning "Could not remove $($playlist.FullName): $($_.Exception.Message)"
            }
        }
    }
    Write-Host "Removed $playlistCount playlist files" -ForegroundColor Green
    
    # Get all audio files
    Write-Host "Scanning for audio files..." -ForegroundColor Cyan
    $audioExtensions = @('*.mp3', '*.flac', '*.wav', '*.m4a', '*.aac', '*.ogg', '*.wma', '*.ape', '*.wv')
    $audioFiles = @()
    foreach ($ext in $audioExtensions) {
        $audioFiles += Get-ChildItem -Path $SourcePath -Filter $ext -Recurse -ErrorAction SilentlyContinue
    }
    
    Write-Host "Found $($audioFiles.Count) audio files" -ForegroundColor Green
    
    if ($audioFiles.Count -eq 0) {
        Write-Host "No audio files found in $SourcePath" -ForegroundColor Red
        return
    }
    
    # Initialize statistics
    $stats = @{
        Processed = 0
        Successful = 0
        Failed = 0
        Genres = @{}
        Artists = @{}
        Albums = @{}
        Years = @{}
    }
    
    $processedFiles = @()
    $startTime = Get-Date
    
    Write-Host "\nProcessing audio files..." -ForegroundColor Cyan
    
    foreach ($file in $audioFiles) {
        $stats.Processed++
        $percentComplete = ($stats.Processed / $audioFiles.Count) * 100
        
        Write-Progress -Activity "Processing Music Files" -Status "Processing $($file.Name) ($($stats.Processed)/$($audioFiles.Count))" -PercentComplete $percentComplete
        
        try {
            # Extract raw metadata
            $rawMetadata = Get-EnhancedMetadata -FilePath $file.FullName
            
            # Clean and enhance metadata
            $metadata = Get-CleanedMetadata -FilePath $file.FullName -RawMetadata $rawMetadata
            
            # Sanitize for file system
            $safeArtist = Sanitize-Name -Name $metadata.Artist
            $safeAlbum = Sanitize-Name -Name $metadata.Album
            $safeTitle = Sanitize-Name -Name $metadata.Title
            $safeGenre = Sanitize-Name -Name $metadata.Genre
            
            # Create directory structure: Genre\Artist\Year - Album
            $albumFolder = if ($metadata.Year -ne "Unknown") { 
                "$($metadata.Year) - $safeAlbum" 
            } else { 
                $safeAlbum 
            }
            $targetDir = Join-Path $DestinationPath "$safeGenre\$safeArtist\$albumFolder"
            
            if (-not $DryRun -and -not (Test-Path $targetDir)) {
                New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
            }
            
            # Create new filename with proper track numbering
            $trackNum = if ($metadata.Track -match '\d+') {
                [int]([regex]::Match($metadata.Track, '\d+').Value)
            } else {
                1
            }
            $trackNumFormatted = $trackNum.ToString().PadLeft(2, '0')
            
            $newFileName = "$trackNumFormatted - $safeTitle$($file.Extension)"
            $targetFile = Join-Path $targetDir $newFileName
            
            # Copy file
            if (-not $DryRun) {
                Copy-Item $file.FullName $targetFile -Force
            }
            
            # Copy artwork files
            $artworkExtensions = @('*.jpg', '*.jpeg', '*.png', '*.bmp', '*.gif', '*.tiff')
            foreach ($artExt in $artworkExtensions) {
                $artworkFiles = Get-ChildItem -Path $file.Directory -Filter $artExt -ErrorAction SilentlyContinue
                foreach ($artwork in $artworkFiles) {
                    $artworkTarget = Join-Path $targetDir $artwork.Name
                    if (-not $DryRun -and -not (Test-Path $artworkTarget)) {
                        Copy-Item $artwork.FullName $artworkTarget -Force -ErrorAction SilentlyContinue
                    }
                }
            }
            
            # Update statistics
            if ($stats.Genres.ContainsKey($metadata.Genre)) {
                $stats.Genres[$metadata.Genre]++
            } else {
                $stats.Genres[$metadata.Genre] = 1
            }
            
            if ($stats.Artists.ContainsKey($metadata.Artist)) {
                $stats.Artists[$metadata.Artist]++
            } else {
                $stats.Artists[$metadata.Artist] = 1
            }
            
            $albumKey = "$($metadata.Artist) - $($metadata.Album)"
            if ($stats.Albums.ContainsKey($albumKey)) {
                $stats.Albums[$albumKey]++
            } else {
                $stats.Albums[$albumKey] = 1
            }
            
            if ($stats.Years.ContainsKey($metadata.Year)) {
                $stats.Years[$metadata.Year]++
            } else {
                $stats.Years[$metadata.Year] = 1
            }
            
            # Add to processed files for playlist generation
            $processedFiles += [PSCustomObject]@{
                FullName = $targetFile
                BaseName = $safeTitle
                Title = $metadata.Title
                Genre = $metadata.Genre
                Artist = $metadata.Artist
                Album = $metadata.Album
                Year = $metadata.Year
                Track = $trackNum
                Duration = $rawMetadata.Duration
            }
            
            $stats.Successful++
            
            if ($stats.Processed % 100 -eq 0) {
                Write-Host "  Processed $($stats.Processed) files..." -ForegroundColor Gray
            }
            
        }
        catch {
            Write-Warning "Failed to process $($file.FullName): $($_.Exception.Message)"
            $stats.Failed++
        }
    }
    
    Write-Progress -Activity "Processing Music Files" -Completed
    
    # Generate enhanced playlists
    if (-not $DryRun -and $processedFiles.Count -gt 0) {
        Write-Host "\nGenerating enhanced playlists..." -ForegroundColor Cyan
        
        # Genre playlists
        foreach ($genre in $stats.Genres.Keys) {
            $genreFiles = $processedFiles | Where-Object { $_.Genre -eq $genre } | Sort-Object Artist, Album, Track
            if ($genreFiles.Count -gt 0) {
                $playlistPath = Join-Path $DestinationPath "[Genre] $genre.m3u"
                Create-EnhancedPlaylist -PlaylistPath $playlistPath -Files $genreFiles -PlaylistName "Genre: $genre"
                Write-Host "  Created: [Genre] $genre.m3u ($($genreFiles.Count) tracks)" -ForegroundColor Green
            }
        }
        
        # Artist playlists (for artists with multiple tracks)
        $artistsWithMultipleTracks = $stats.Artists.Keys | Where-Object { $stats.Artists[$_] -gt 1 }
        foreach ($artist in $artistsWithMultipleTracks) {
            $artistFiles = $processedFiles | Where-Object { $_.Artist -eq $artist } | Sort-Object Album, Track
            $safeArtistName = Sanitize-Name -Name $artist
            $playlistPath = Join-Path $DestinationPath "[Artist] $safeArtistName.m3u"
            Create-EnhancedPlaylist -PlaylistPath $playlistPath -Files $artistFiles -PlaylistName "Artist: $artist"
            Write-Host "  Created: [Artist] $safeArtistName.m3u ($($artistFiles.Count) tracks)" -ForegroundColor Green
        }
        
        # Year playlists (for years with significant content)
        $significantYears = $stats.Years.Keys | Where-Object { $_ -ne "Unknown" -and $stats.Years[$_] -gt 10 }
        foreach ($year in $significantYears) {
            $yearFiles = $processedFiles | Where-Object { $_.Year -eq $year } | Sort-Object Artist, Album, Track
            $playlistPath = Join-Path $DestinationPath "[Year] $year.m3u"
            Create-EnhancedPlaylist -PlaylistPath $playlistPath -Files $yearFiles -PlaylistName "Year: $year"
            Write-Host "  Created: [Year] $year.m3u ($($yearFiles.Count) tracks)" -ForegroundColor Green
        }
        
        # Decade playlists
        $decades = @{}
        foreach ($year in ($stats.Years.Keys | Where-Object { $_ -ne "Unknown" -and $_ -match '^\d{4}$' })) {
            $decade = [math]::Floor([int]$year / 10) * 10
            $decadeKey = "${decade}s"
            if ($decades.ContainsKey($decadeKey)) {
                $decades[$decadeKey] += $stats.Years[$year]
            } else {
                $decades[$decadeKey] = $stats.Years[$year]
            }
        }
        
        foreach ($decade in ($decades.Keys | Where-Object { $decades[$_] -gt 20 })) {
            $decadeStart = [int]($decade -replace 's', '')
            $decadeEnd = $decadeStart + 9
            $decadeFiles = $processedFiles | Where-Object { 
                $_.Year -ne "Unknown" -and $_.Year -match '^\d{4}$' -and 
                [int]$_.Year -ge $decadeStart -and [int]$_.Year -le $decadeEnd 
            } | Sort-Object Year, Artist, Album, Track
            
            if ($decadeFiles.Count -gt 0) {
                $playlistPath = Join-Path $DestinationPath "[Decade] $decade.m3u"
                Create-EnhancedPlaylist -PlaylistPath $playlistPath -Files $decadeFiles -PlaylistName "Decade: $decade"
                Write-Host "  Created: [Decade] $decade.m3u ($($decadeFiles.Count) tracks)" -ForegroundColor Green
            }
        }
    }
    
    $endTime = Get-Date
    $duration = $endTime - $startTime
    
    # Display comprehensive summary
    Write-Host "\n" + "=" * 60 -ForegroundColor Green
    Write-Host "COMPREHENSIVE MUSIC ORGANIZATION COMPLETE" -ForegroundColor Green
    Write-Host "=" * 60 -ForegroundColor Green
    Write-Host "Processing Time: $($duration.ToString('hh\:mm\:ss'))" -ForegroundColor Yellow
    Write-Host "Total Files Processed: $($stats.Processed)" -ForegroundColor Yellow
    Write-Host "Successfully Organized: $($stats.Successful)" -ForegroundColor Green
    Write-Host "Failed: $($stats.Failed)" -ForegroundColor Red
    Write-Host "Unique Genres: $($stats.Genres.Count)" -ForegroundColor Yellow
    Write-Host "Unique Artists: $($stats.Artists.Count)" -ForegroundColor Yellow
    Write-Host "Unique Albums: $($stats.Albums.Count)" -ForegroundColor Yellow
    
    Write-Host "\nGenre Distribution:" -ForegroundColor Cyan
    $stats.Genres.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
        $percentage = ($_.Value / $stats.Successful) * 100
        Write-Host "  $($_.Key): $($_.Value) tracks ($($percentage.ToString('F1'))%)" -ForegroundColor White
    }
    
    Write-Host "\nTop Artists:" -ForegroundColor Cyan
    $stats.Artists.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 10 | ForEach-Object {
        Write-Host "  $($_.Key): $($_.Value) tracks" -ForegroundColor White
    }
    
    Write-Host "\nYear Distribution:" -ForegroundColor Cyan
    $stats.Years.GetEnumerator() | Sort-Object Key | ForEach-Object {
        if ($_.Value -gt 5) {
            Write-Host "  $($_.Key): $($_.Value) tracks" -ForegroundColor White
        }
    }
    
    if ($DryRun) {
        Write-Host "\n" + "!" * 60 -ForegroundColor Red
        Write-Host "THIS WAS A DRY RUN - NO FILES WERE ACTUALLY MOVED" -ForegroundColor Red
        Write-Host "Run without -DryRun parameter to perform the actual organization" -ForegroundColor Red
        Write-Host "!" * 60 -ForegroundColor Red
    } else {
        Write-Host "\nOrganized library location: $DestinationPath" -ForegroundColor Green
    }
    
    Write-Host "\nRecommendations for further improvement:" -ForegroundColor Yellow
    Write-Host "1. Use MusicBrainz Picard for acoustic fingerprinting and metadata correction" -ForegroundColor White
    Write-Host "2. Review and manually correct any 'Unknown Artist' or 'Unknown Album' entries" -ForegroundColor White
    Write-Host "3. Consider using additional tools like Mp3tag for batch metadata editing" -ForegroundColor White
    Write-Host "4. Regularly backup your organized music library" -ForegroundColor White
}

# Execute the main function
Start-ComprehensiveMusicFix

Write-Host "\nScript execution completed." -ForegroundColor Green