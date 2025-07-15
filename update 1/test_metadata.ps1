# Test metadata extraction
param([string]$TestFile)

if (-not $TestFile) {
    $TestFile = "H:\Music\1989 - Bleach @320\01 - Blew.mp3"
}

Write-Host "Testing metadata extraction for: $TestFile"

if (-not (Test-Path $TestFile)) {
    Write-Host "File not found: $TestFile"
    exit
}

try {
    # Extract metadata using Shell.Application
    $shell = New-Object -ComObject Shell.Application
    $folder = $shell.Namespace((Get-Item $TestFile).DirectoryName)
    $file = $folder.ParseName((Get-Item $TestFile).Name)
    
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
    
    $metadata = @{}
    
    foreach ($prop in $properties.GetEnumerator()) {
        $value = $folder.GetDetailsOf($file, $prop.Value)
        Write-Host "$($prop.Key) (index $($prop.Value)): '$value'"
        if ($value -and $value.Trim() -ne "") {
            $metadata[$prop.Key] = $value.Trim()
        }
    }
    
    Write-Host "`nExtracted metadata:"
    foreach ($key in $metadata.Keys) {
        Write-Host "  $key = '$($metadata[$key])'"
    }
    
    Write-Host "`nTesting hashtable key checks:"
    Write-Host "Genre exists: $($metadata.ContainsKey('Genre'))"
    Write-Host "Genre value: '$($metadata['Genre'])'"
    Write-Host "Genre boolean test: $([bool]$metadata['Genre'])"
    Write-Host "Genre if test: $(if ($metadata['Genre']) { 'TRUE' } else { 'FALSE' })"
    
}
catch {
    Write-Host "Error: $($_.Exception.Message)"
}