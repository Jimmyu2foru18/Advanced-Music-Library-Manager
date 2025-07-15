# Test New-OrganizedPath function

# Define the functions we need
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    Write-Host "[$Level] $Message"
}

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
    
    return @{
        ArtistFolder = $artistFolder
        AlbumFolder = $albumFolder
        FullPath = Join-Path $albumFolder $fileName
        FileName = $fileName
    }
}

# Create test metadata
$testMetadata = @{
    'Artist' = 'Test Artist'
    'Album' = 'Test Album'
    'Title' = 'Test Title'
    'Genre' = 'Test Genre'
    'Year' = '2023'
    'Track' = '1'
}

Write-Host "Test metadata:"
foreach ($key in $testMetadata.Keys) {
    Write-Host "  $key = '$($testMetadata[$key])'"
}

Write-Host "`nTesting New-OrganizedPath function..."

try {
    $result = New-OrganizedPath -Metadata $testMetadata -BaseOutputPath "C:\Test" -OriginalExtension ".mp3"
    
    Write-Host "Result:"
    Write-Host "  FullPath: $($result.FullPath)"
    Write-Host "  FileName: $($result.FileName)"
    Write-Host "  ArtistFolder: $($result.ArtistFolder)"
    Write-Host "  AlbumFolder: $($result.AlbumFolder)"
}
catch {
    Write-Host "Error: $($_.Exception.Message)"
    Write-Host "Stack trace: $($_.ScriptStackTrace)"
}