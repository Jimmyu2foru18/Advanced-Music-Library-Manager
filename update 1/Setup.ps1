# Music Library Manager Setup Script
# This script helps set up the Music Library Manager for first-time use

param(
    [switch]$InstallDependencies,
    [switch]$CreateDesktopShortcut,
    [switch]$SetupApiKeys,
    [switch]$RunTests
)

# Colors for output
$Green = "Green"
$Yellow = "Yellow"
$Red = "Red"
$Cyan = "Cyan"

function Write-Header {
    param([string]$Text)
    Write-Host "`n" -NoNewline
    Write-Host "=" * 60 -ForegroundColor $Cyan
    Write-Host $Text -ForegroundColor $Cyan
    Write-Host "=" * 60 -ForegroundColor $Cyan
}

function Write-Success {
    param([string]$Text)
    Write-Host "✓ $Text" -ForegroundColor $Green
}

function Write-Warning {
    param([string]$Text)
    Write-Host "⚠ $Text" -ForegroundColor $Yellow
}

function Write-Error {
    param([string]$Text)
    Write-Host "✗ $Text" -ForegroundColor $Red
}

function Test-Prerequisites {
    Write-Header "Checking Prerequisites"
    
    $allGood = $true
    
    # Check PowerShell version
    $psVersion = $PSVersionTable.PSVersion
    if ($psVersion.Major -ge 5) {
        Write-Success "PowerShell version: $($psVersion.ToString())"
    } else {
        Write-Error "PowerShell 5.0 or higher required. Current version: $($psVersion.ToString())"
        $allGood = $false
    }
    
    # Check Windows version
    $osVersion = [System.Environment]::OSVersion.Version
    if ($osVersion.Major -ge 10 -or ($osVersion.Major -eq 6 -and $osVersion.Minor -ge 1)) {
        Write-Success "Windows version: $($osVersion.ToString())"
    } else {
        Write-Warning "Windows 7 or higher recommended. Current version: $($osVersion.ToString())"
    }
    
    # Check execution policy
    $executionPolicy = Get-ExecutionPolicy
    if ($executionPolicy -eq "Restricted") {
        Write-Warning "PowerShell execution policy is Restricted. You may need to change it."
        Write-Host "Run this command as Administrator: Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine" -ForegroundColor $Yellow
    } else {
        Write-Success "PowerShell execution policy: $executionPolicy"
    }
    
    # Check .NET Framework
    try {
        $netVersion = [System.Runtime.InteropServices.RuntimeInformation]::FrameworkDescription
        Write-Success ".NET Framework: $netVersion"
    } catch {
        Write-Warning "Could not determine .NET Framework version"
    }
    
    # Check available disk space
    $drive = (Get-Location).Drive
    $freeSpace = (Get-WmiObject -Class Win32_LogicalDisk -Filter "DeviceID='$($drive.Name)'").FreeSpace
    $freeSpaceGB = [math]::Round($freeSpace / 1GB, 2)
    
    if ($freeSpaceGB -gt 10) {
        Write-Success "Available disk space: $freeSpaceGB GB"
    } else {
        Write-Warning "Low disk space: $freeSpaceGB GB. Consider freeing up space."
    }
    
    return $allGood
}

function Test-ScriptFiles {
    Write-Header "Checking Script Files"
    
    $requiredFiles = @(
        "MusicLibraryManager.ps1",
        "MusicLibraryGUI.ps1",
        "MusicLibraryConfig.json",
        "MusicLibraryUtils.psm1",
        "Launch_Music_Manager.bat",
        "README.md"
    )
    
    $allFilesPresent = $true
    
    foreach ($file in $requiredFiles) {
        $filePath = Join-Path $PSScriptRoot $file
        if (Test-Path $filePath) {
            $fileSize = (Get-Item $filePath).Length
            Write-Success "$file ($([math]::Round($fileSize / 1KB, 1)) KB)"
        } else {
            Write-Error "Missing file: $file"
            $allFilesPresent = $false
        }
    }
    
    return $allFilesPresent
}

function Install-Dependencies {
    Write-Header "Installing Dependencies"
    
    # Check if running as administrator
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    
    if (-not $isAdmin) {
        Write-Warning "Not running as administrator. Some installations may fail."
    }
    
    # Install NuGet provider if needed
    try {
        $nugetProvider = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
        if (-not $nugetProvider) {
            Write-Host "Installing NuGet provider..." -ForegroundColor $Yellow
            Install-PackageProvider -Name NuGet -Force -Scope CurrentUser
            Write-Success "NuGet provider installed"
        } else {
            Write-Success "NuGet provider already installed"
        }
    } catch {
        Write-Warning "Could not install NuGet provider: $($_.Exception.Message)"
    }
    
    # Set PSGallery as trusted
    try {
        $psGallery = Get-PSRepository -Name PSGallery
        if ($psGallery.InstallationPolicy -ne "Trusted") {
            Write-Host "Setting PSGallery as trusted..." -ForegroundColor $Yellow
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
            Write-Success "PSGallery set as trusted"
        } else {
            Write-Success "PSGallery already trusted"
        }
    } catch {
        Write-Warning "Could not configure PSGallery: $($_.Exception.Message)"
    }
    
    # Install useful modules
    $modules = @(
        @{ Name = "PowerShellGet"; Description = "PowerShell module management" },
        @{ Name = "ImportExcel"; Description = "Excel file handling" }
    )
    
    foreach ($module in $modules) {
        try {
            $installedModule = Get-Module -Name $module.Name -ListAvailable
            if (-not $installedModule) {
                Write-Host "Installing $($module.Name)..." -ForegroundColor $Yellow
                Install-Module -Name $module.Name -Scope CurrentUser -Force
                Write-Success "$($module.Name) installed - $($module.Description)"
            } else {
                Write-Success "$($module.Name) already installed - $($module.Description)"
            }
        } catch {
            Write-Warning "Could not install $($module.Name): $($_.Exception.Message)"
        }
    }
}

function New-DesktopShortcut {
    Write-Header "Creating Desktop Shortcut"
    
    try {
        $desktopPath = [Environment]::GetFolderPath("Desktop")
        $shortcutPath = Join-Path $desktopPath "Music Library Manager.lnk"
        $batchPath = Join-Path $PSScriptRoot "Launch_Music_Manager.bat"
        
        $shell = New-Object -ComObject WScript.Shell
        $shortcut = $shell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = $batchPath
        $shortcut.WorkingDirectory = $PSScriptRoot
        $shortcut.Description = "Advanced Music Library Manager"
        $shortcut.IconLocation = "shell32.dll,168"  # Music note icon
        $shortcut.Save()
        
        Write-Success "Desktop shortcut created: $shortcutPath"
    } catch {
        Write-Error "Failed to create desktop shortcut: $($_.Exception.Message)"
    }
}

function Set-ApiKeys {
    Write-Header "API Keys Setup"
    
    Write-Host "The Music Library Manager can use several online services to enhance metadata:" -ForegroundColor $Cyan
    Write-Host ""
    Write-Host "1. MusicBrainz (Free, no API key required)" -ForegroundColor $Green
    Write-Host "2. Last.fm (Free API key required)" -ForegroundColor $Yellow
    Write-Host "3. Discogs (Free API key required)" -ForegroundColor $Yellow
    Write-Host "4. Spotify (Free API key required)" -ForegroundColor $Yellow
    Write-Host ""
    
    $setupKeys = Read-Host "Would you like to set up API keys now? (y/n)"
    
    if ($setupKeys -eq 'y' -or $setupKeys -eq 'Y') {
        $configPath = Join-Path $PSScriptRoot "MusicLibraryConfig.json"
        
        try {
            $config = Get-Content $configPath | ConvertFrom-Json -AsHashtable
            
            Write-Host ""
            Write-Host "Last.fm API Setup:" -ForegroundColor $Cyan
            Write-Host "1. Visit: https://www.last.fm/api/account/create"
            Write-Host "2. Create an account and get your API key"
            $lastfmKey = Read-Host "Enter your Last.fm API key (or press Enter to skip)"
            
            if ($lastfmKey) {
                $config.SearchProviders.LastFm.ApiKey = $lastfmKey
                $config.SearchProviders.LastFm.Enabled = $true
                Write-Success "Last.fm API key configured"
            }
            
            Write-Host ""
            Write-Host "Discogs API Setup:" -ForegroundColor $Cyan
            Write-Host "1. Visit: https://www.discogs.com/settings/developers"
            Write-Host "2. Generate a personal access token"
            $discogsToken = Read-Host "Enter your Discogs token (or press Enter to skip)"
            
            if ($discogsToken) {
                $config.SearchProviders.Discogs.Token = $discogsToken
                $config.SearchProviders.Discogs.Enabled = $true
                Write-Success "Discogs API token configured"
            }
            
            Write-Host ""
            Write-Host "Spotify API Setup:" -ForegroundColor $Cyan
            Write-Host "1. Visit: https://developer.spotify.com/dashboard"
            Write-Host "2. Create an app and get Client ID and Secret"
            $spotifyId = Read-Host "Enter your Spotify Client ID (or press Enter to skip)"
            
            if ($spotifyId) {
                $spotifySecret = Read-Host "Enter your Spotify Client Secret"
                $config.SearchProviders.Spotify.ClientId = $spotifyId
                $config.SearchProviders.Spotify.ClientSecret = $spotifySecret
                $config.SearchProviders.Spotify.Enabled = $true
                Write-Success "Spotify API credentials configured"
            }
            
            # Save configuration
            $config | ConvertTo-Json -Depth 10 | Out-File -FilePath $configPath -Encoding UTF8
            Write-Success "Configuration saved"
            
        } catch {
            Write-Error "Failed to configure API keys: $($_.Exception.Message)"
        }
    } else {
        Write-Host "API keys can be configured later using the GUI or by editing MusicLibraryConfig.json" -ForegroundColor $Yellow
    }
}

function Invoke-Tests {
    Write-Header "Running Tests"
    
    # Test basic functionality
    try {
        # Test configuration loading
        $configPath = Join-Path $PSScriptRoot "MusicLibraryConfig.json"
        $config = Get-Content $configPath | ConvertFrom-Json
        Write-Success "Configuration file loads correctly"
        
        # Test module import
        $modulePath = Join-Path $PSScriptRoot "MusicLibraryUtils.psm1"
        Import-Module $modulePath -Force
        Write-Success "Utility module imports correctly"
        
        # Test COM object creation (for metadata extraction)
        $shell = New-Object -ComObject Shell.Application
        Write-Success "COM objects can be created (metadata extraction will work)"
        
        # Test web connectivity
        $response = Invoke-WebRequest -Uri "https://musicbrainz.org" -TimeoutSec 5 -UseBasicParsing
        if ($response.StatusCode -eq 200) {
            Write-Success "Internet connectivity to MusicBrainz confirmed"
        }
        
    } catch {
        Write-Warning "Test failed: $($_.Exception.Message)"
    }
    
    # Test sample music file processing (if available)
    $musicFiles = Get-ChildItem -Path $PSScriptRoot -Include *.mp3,*.flac,*.m4a -Recurse | Select-Object -First 1
    if ($musicFiles) {
        try {
            $file = $musicFiles[0]
            $shell = New-Object -ComObject Shell.Application
            $folder = $shell.Namespace($file.DirectoryName)
            $fileItem = $folder.ParseName($file.Name)
            $title = $folder.GetDetailsOf($fileItem, 21)
            Write-Success "Metadata extraction test successful on: $($file.Name)"
        } catch {
            Write-Warning "Metadata extraction test failed: $($_.Exception.Message)"
        }
    } else {
        Write-Warning "No music files found for testing"
    }
}

function Show-QuickStart {
    Write-Header "Quick Start Guide"
    
    Write-Host "Your Music Library Manager is now set up! Here's how to get started:" -ForegroundColor $Green
    Write-Host ""
    Write-Host "Option 1: Use the GUI (Recommended for beginners)" -ForegroundColor $Cyan
    Write-Host "  • Double-click the desktop shortcut 'Music Library Manager'" -ForegroundColor $Yellow
    Write-Host "  • Or run: Launch_Music_Manager.bat" -ForegroundColor $Yellow
    Write-Host ""
    Write-Host "Option 2: Use the Command Line (Advanced users)" -ForegroundColor $Cyan
    Write-Host "  • powershell -ExecutionPolicy Bypass -File MusicLibraryManager.ps1 -SourcePath 'h:\Music' -DryRun" -ForegroundColor $Yellow
    Write-Host ""
    Write-Host "Important Tips:" -ForegroundColor $Cyan
    Write-Host "  • Always start with -DryRun to preview changes" -ForegroundColor $Yellow
    Write-Host "  • Backup your music library before processing" -ForegroundColor $Yellow
    Write-Host "  • Check the README.md file for detailed documentation" -ForegroundColor $Yellow
    Write-Host "  • Configure API keys for better metadata accuracy" -ForegroundColor $Yellow
    Write-Host ""
    Write-Host "Need help? Check the README.md file or the built-in help:" -ForegroundColor $Green
    Write-Host "  Get-Help .\MusicLibraryManager.ps1 -Full" -ForegroundColor $Yellow
}

# Main setup process
Write-Header "Music Library Manager Setup"
Write-Host "Welcome to the Music Library Manager setup!" -ForegroundColor $Green
Write-Host "This script will help you get everything configured and ready to use." -ForegroundColor $Green

# Run prerequisite checks
$prereqsOk = Test-Prerequisites
$filesOk = Test-ScriptFiles

if (-not $prereqsOk -or -not $filesOk) {
    Write-Error "Setup cannot continue due to missing prerequisites or files."
    exit 1
}

# Install dependencies if requested
if ($InstallDependencies) {
    Install-Dependencies
}

# Create desktop shortcut if requested
if ($CreateDesktopShortcut) {
    New-DesktopShortcut
}

# Setup API keys if requested
if ($SetupApiKeys) {
    Set-ApiKeys
}

# Run tests if requested
if ($RunTests) {
    Invoke-Tests
}

# If no specific options were provided, run interactive setup
if (-not $InstallDependencies -and -not $CreateDesktopShortcut -and -not $SetupApiKeys -and -not $RunTests) {
    Write-Host ""
    Write-Host "Would you like to run the full interactive setup? This will:" -ForegroundColor $Cyan
    Write-Host "  • Install helpful PowerShell modules" -ForegroundColor $Yellow
    Write-Host "  • Create a desktop shortcut" -ForegroundColor $Yellow
    Write-Host "  • Help you configure API keys" -ForegroundColor $Yellow
    Write-Host "  • Run basic tests" -ForegroundColor $Yellow
    
    $runSetup = Read-Host "Run full setup? (y/n)"
    
    if ($runSetup -eq 'y' -or $runSetup -eq 'Y') {
        Install-Dependencies
        New-DesktopShortcut
        Set-ApiKeys
        Invoke-Tests
    }
}

# Show quick start guide
Show-QuickStart

Write-Host ""
Write-Success "Setup completed successfully!"
Write-Host "Enjoy organizing your music library! 🎵" -ForegroundColor $Green