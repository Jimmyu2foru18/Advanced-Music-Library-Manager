# Advanced Music Library Manager

A comprehensive music library organization tool that combines local file analysis with internet search capabilities to automatically correct and enhance music metadata, organize files, and create playlists.

## Features

### 🎵 **Intelligent Metadata Correction**
- **Web Search Integration**: Searches MusicBrainz, Last.fm, Discogs, and Spotify for accurate metadata
- **Multiple Fallback Sources**: Uses folder names, file names, and existing tags when online data isn't available
- **Smart Genre Classification**: Advanced genre mapping and artist-based genre detection
- **Standardization**: Cleans and standardizes artist names, album titles, and genres

### 📁 **Advanced File Organization**
- **Flexible Folder Structures**: Multiple organization patterns (Genre\Artist\Year - Album, etc.)
- **Safe File Naming**: Removes invalid characters and creates consistent naming conventions
- **Artwork Handling**: Copies and organizes album artwork alongside music files
- **Duplicate Detection**: Identifies and handles duplicate files intelligently

### 🔍 **Comprehensive Search & Analysis**
- **Multi-Format Support**: MP3, FLAC, M4A, WAV, WMA, OGG
- **Metadata Extraction**: Uses multiple methods to extract existing metadata
- **Online Verification**: Cross-references local data with authoritative music databases
- **Confidence Scoring**: Rates the reliability of metadata corrections

### 📊 **Detailed Reporting & Manifests**
- **Processing Manifests**: Complete JSON logs of all changes and corrections
- **Statistical Reports**: Detailed breakdowns by genre, artist, year, and album
- **Error Tracking**: Comprehensive error logging and recovery suggestions
- **Before/After Comparisons**: Shows original vs. corrected metadata

### 🎧 **Playlist Generation**
- **Multiple Playlist Types**: By genre, artist, year, decade, and custom criteria
- **Smart Filtering**: Minimum track requirements and quality thresholds
- **M3U Format**: Compatible with most media players
- **Automatic Updates**: Regenerates playlists after organization

### 🖥️ **User-Friendly Interface**
- **GUI Application**: Easy-to-use Windows Forms interface
- **Command Line Support**: Full PowerShell script for automation
- **Dry Run Mode**: Preview changes before applying them
- **Real-time Progress**: Live updates during processing

## Quick Start

### Option 1: GUI Application (Recommended for beginners)

1. **Launch the GUI**:
   ```powershell
   powershell -ExecutionPolicy Bypass -File "h:\Music\MusicLibraryGUI.ps1"
   ```

2. **Configure Settings**:
   - Set your source music folder
   - Choose output destination
   - Enable/disable web search
   - Set processing options

3. **Start Processing**:
   - Click "Start Processing"
   - Monitor progress in real-time
   - Review results in the Results tab

### Option 2: Command Line (Advanced users)

1. **Basic Usage**:
   ```powershell
   powershell -ExecutionPolicy Bypass -File "h:\Music\MusicLibraryManager.ps1" -SourcePath "h:\Music" -OutputPath "h:\Music_Organized" -DryRun
   ```

2. **Full Processing**:
   ```powershell
   powershell -ExecutionPolicy Bypass -File "h:\Music\MusicLibraryManager.ps1" -SourcePath "h:\Music" -OutputPath "h:\Music_Organized" -EnableWebSearch
   ```

## Configuration

### API Keys (Optional but Recommended)

For enhanced metadata accuracy, obtain free API keys from:

1. **Last.fm API**:
   - Visit: https://www.last.fm/api/account/create
   - Add your API key to `MusicLibraryConfig.json`

2. **Discogs API**:
   - Visit: https://www.discogs.com/settings/developers
   - Generate a personal access token

3. **Spotify API**:
   - Visit: https://developer.spotify.com/dashboard
   - Create an app and get Client ID/Secret

### Configuration File

Edit `MusicLibraryConfig.json` to customize:

```json
{
  "SearchProviders": {
    "LastFm": {
      "Enabled": true,
      "ApiKey": "your_lastfm_api_key_here"
    }
  },
  "FileOrganization": {
    "Structure": "Genre\\Artist\\Year - Album",
    "FileNaming": "Track - Title"
  }
}
```

## File Organization Patterns

### Supported Folder Structures:
- `Genre\Artist\Year - Album` (Default)
- `Artist\Year - Album`
- `Artist\Album`
- `Year\Artist - Album`
- `Genre\Year\Artist - Album`

### File Naming Conventions:
- `Track - Title` (Default)
- `Track. Title`
- `Artist - Title`
- `Title`
- `Track - Artist - Title`

## Command Line Parameters

| Parameter | Description | Default |
|-----------|-------------|----------|
| `-SourcePath` | Source music folder path | Required |
| `-OutputPath` | Output destination folder | `$SourcePath\Organized` |
| `-DryRun` | Preview mode (no files moved) | `$false` |
| `-EnableWebSearch` | Enable internet metadata search | `$true` |
| `-MaxConcurrentSearches` | Concurrent web searches | `5` |
| `-LogPath` | Custom log file location | `$SourcePath\MusicLibraryManager.log` |

## Output Structure

After processing, your organized library will contain:

```
Music_Organized/
├── Rock/
│   ├── Nirvana/
│   │   └── 1991 - Nevermind/
│   │       ├── 01 - Smells Like Teen Spirit.mp3
│   │       ├── 02 - In Bloom.mp3
│   │       └── cover.jpg
│   └── Pearl Jam/
│       └── 1991 - Ten/
├── Hip-Hop/
├── Electronic/
├── Playlists/
│   ├── By Genre/
│   │   ├── Rock.m3u
│   │   └── Hip-Hop.m3u
│   ├── By Artist/
│   └── By Year/
├── MusicLibraryManifest.json
└── ProcessingSummary.txt
```

## Troubleshooting

### Common Issues

1. **"Execution Policy" Error**:
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

2. **"Collection was modified" Warnings**:
   - These are usually harmless and indicate concurrent file access
   - The script will continue processing

3. **Web Search Timeouts**:
   - Check internet connection
   - Reduce `MaxConcurrentSearches` parameter
   - Some APIs have rate limits

4. **Missing Metadata**:
   - Enable web search for better results
   - Check folder and file naming conventions
   - Review the error log for specific issues

### Performance Tips

1. **Large Libraries (>10,000 files)**:
   - Process in smaller batches
   - Disable web search for initial organization
   - Use SSD storage for better performance

2. **Network Optimization**:
   - Stable internet connection recommended
   - Consider processing during off-peak hours
   - Some APIs have daily limits

3. **System Resources**:
   - Close other applications during processing
   - Ensure sufficient disk space (2x library size)
   - Monitor memory usage for very large libraries

## Advanced Features

### Custom Genre Mapping

Add custom genre mappings in `MusicLibraryConfig.json`:

```json
"GenreMapping": {
  "Alternative Rock": "Alternative",
  "Hip Hop": "Hip-Hop",
  "Your Custom Genre": "Standard Genre"
}
```

### Artist-Specific Genres

Define genres for specific artists:

```json
"ArtistGenreMapping": {
  "Nirvana": "Grunge",
  "Eminem": "Hip-Hop",
  "Your Artist": "Your Genre"
}
```

### Batch Processing

Process multiple folders:

```powershell
$folders = @("h:\Music1", "h:\Music2", "h:\Music3")
foreach ($folder in $folders) {
    & "h:\Music\MusicLibraryManager.ps1" -SourcePath $folder -OutputPath "$folder\Organized"
}
```

## Data Sources

The application uses these authoritative music databases:

- **MusicBrainz**: Open music encyclopedia (primary source)
- **Last.fm**: Social music platform with extensive tagging
- **Discogs**: Comprehensive music database and marketplace
- **Spotify**: Streaming service with detailed metadata

## Privacy & Security

- **Local Processing**: All file operations happen locally
- **API Calls**: Only metadata queries sent to external services
- **No File Upload**: Your music files never leave your computer
- **Secure Storage**: API keys stored locally in configuration files

## Support

### Getting Help

1. **Check the Log**: Review `MusicLibraryManager.log` for detailed error information
2. **Dry Run First**: Always test with `-DryRun` before actual processing
3. **Backup Important Files**: Create backups of irreplaceable music collections

### Reporting Issues

When reporting problems, include:
- PowerShell version (`$PSVersionTable.PSVersion`)
- Windows version
- Sample file paths and names
- Relevant log entries
- Configuration settings used

## License

This project is provided as-is for personal use. Please respect copyright laws and only organize music you legally own.

## Changelog

### Version 2.0
- Added GUI application
- Enhanced web search integration
- Improved metadata correction algorithms
- Added comprehensive configuration system
- Better error handling and logging
- Support for additional audio formats

### Version 1.0
- Initial release
- Basic file organization
- Simple metadata extraction
- Playlist generation

---

**Happy organizing! 🎵**