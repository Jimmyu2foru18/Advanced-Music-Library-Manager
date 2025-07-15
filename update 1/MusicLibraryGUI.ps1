# Music Library Manager - GUI Application
# Advanced music organization tool with web search integration

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web

# Load configuration
$configPath = Join-Path $PSScriptRoot "MusicLibraryConfig.json"
$config = @{}
if (Test-Path $configPath) {
    try {
        $config = Get-Content $configPath | ConvertFrom-Json -AsHashtable
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error loading configuration: $($_.Exception.Message)", "Configuration Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
}

# Global variables
$Global:ProcessingCancelled = $false
$Global:CurrentOperation = ""

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Music Library Manager v2.0"
$form.Size = New-Object System.Drawing.Size(800, 600)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false
$form.Icon = [System.Drawing.SystemIcons]::Application

# Create tab control
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Size = New-Object System.Drawing.Size(780, 550)
$tabControl.Location = New-Object System.Drawing.Point(10, 10)
$form.Controls.Add($tabControl)

# Main Processing Tab
$mainTab = New-Object System.Windows.Forms.TabPage
$mainTab.Text = "Main Processing"
$tabControl.TabPages.Add($mainTab)

# Source folder selection
$lblSource = New-Object System.Windows.Forms.Label
$lblSource.Text = "Source Music Folder:"
$lblSource.Location = New-Object System.Drawing.Point(20, 20)
$lblSource.Size = New-Object System.Drawing.Size(150, 20)
$mainTab.Controls.Add($lblSource)

$txtSource = New-Object System.Windows.Forms.TextBox
$txtSource.Location = New-Object System.Drawing.Point(20, 45)
$txtSource.Size = New-Object System.Drawing.Size(500, 25)
$txtSource.Text = "h:\Music"
$mainTab.Controls.Add($txtSource)

$btnBrowseSource = New-Object System.Windows.Forms.Button
$btnBrowseSource.Text = "Browse..."
$btnBrowseSource.Location = New-Object System.Drawing.Point(530, 43)
$btnBrowseSource.Size = New-Object System.Drawing.Size(80, 28)
$mainTab.Controls.Add($btnBrowseSource)

# Output folder selection
$lblOutput = New-Object System.Windows.Forms.Label
$lblOutput.Text = "Output Folder:"
$lblOutput.Location = New-Object System.Drawing.Point(20, 80)
$lblOutput.Size = New-Object System.Drawing.Size(150, 20)
$mainTab.Controls.Add($lblOutput)

$txtOutput = New-Object System.Windows.Forms.TextBox
$txtOutput.Location = New-Object System.Drawing.Point(20, 105)
$txtOutput.Size = New-Object System.Drawing.Size(500, 25)
$txtOutput.Text = "h:\Music_Organized"
$mainTab.Controls.Add($txtOutput)

$btnBrowseOutput = New-Object System.Windows.Forms.Button
$btnBrowseOutput.Text = "Browse..."
$btnBrowseOutput.Location = New-Object System.Drawing.Point(530, 103)
$btnBrowseOutput.Size = New-Object System.Drawing.Size(80, 28)
$mainTab.Controls.Add($btnBrowseOutput)

# Processing options
$grpOptions = New-Object System.Windows.Forms.GroupBox
$grpOptions.Text = "Processing Options"
$grpOptions.Location = New-Object System.Drawing.Point(20, 150)
$grpOptions.Size = New-Object System.Drawing.Size(590, 120)
$mainTab.Controls.Add($grpOptions)

$chkWebSearch = New-Object System.Windows.Forms.CheckBox
$chkWebSearch.Text = "Enable Web Search for Metadata Correction"
$chkWebSearch.Location = New-Object System.Drawing.Point(15, 25)
$chkWebSearch.Size = New-Object System.Drawing.Size(300, 20)
$chkWebSearch.Checked = $true
$grpOptions.Controls.Add($chkWebSearch)

$chkDryRun = New-Object System.Windows.Forms.CheckBox
$chkDryRun.Text = "Dry Run (Preview Only - No Files Moved)"
$chkDryRun.Location = New-Object System.Drawing.Point(15, 50)
$chkDryRun.Size = New-Object System.Drawing.Size(300, 20)
$chkDryRun.Checked = $true
$grpOptions.Controls.Add($chkDryRun)

$chkCreatePlaylists = New-Object System.Windows.Forms.CheckBox
$chkCreatePlaylists.Text = "Generate Playlists (by Genre, Artist, Year)"
$chkCreatePlaylists.Location = New-Object System.Drawing.Point(15, 75)
$chkCreatePlaylists.Size = New-Object System.Drawing.Size(300, 20)
$chkCreatePlaylists.Checked = $true
$grpOptions.Controls.Add($chkCreatePlaylists)

$chkBackup = New-Object System.Windows.Forms.CheckBox
$chkBackup.Text = "Create Backup of Original Files"
$chkBackup.Location = New-Object System.Drawing.Point(320, 25)
$chkBackup.Size = New-Object System.Drawing.Size(250, 20)
$chkBackup.Checked = $false
$grpOptions.Controls.Add($chkBackup)

$chkCopyArtwork = New-Object System.Windows.Forms.CheckBox
$chkCopyArtwork.Text = "Copy Album Artwork"
$chkCopyArtwork.Location = New-Object System.Drawing.Point(320, 50)
$chkCopyArtwork.Size = New-Object System.Drawing.Size(250, 20)
$chkCopyArtwork.Checked = $true
$grpOptions.Controls.Add($chkCopyArtwork)

# Progress section
$grpProgress = New-Object System.Windows.Forms.GroupBox
$grpProgress.Text = "Processing Progress"
$grpProgress.Location = New-Object System.Drawing.Point(20, 290)
$grpProgress.Size = New-Object System.Drawing.Size(590, 120)
$mainTab.Controls.Add($grpProgress)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(15, 25)
$progressBar.Size = New-Object System.Drawing.Size(560, 25)
$grpProgress.Controls.Add($progressBar)

$lblProgress = New-Object System.Windows.Forms.Label
$lblProgress.Text = "Ready to process..."
$lblProgress.Location = New-Object System.Drawing.Point(15, 55)
$lblProgress.Size = New-Object System.Drawing.Size(560, 20)
$grpProgress.Controls.Add($lblProgress)

$lblStats = New-Object System.Windows.Forms.Label
$lblStats.Text = "Files: 0 | Processed: 0 | Corrected: 0 | Errors: 0"
$lblStats.Location = New-Object System.Drawing.Point(15, 80)
$lblStats.Size = New-Object System.Drawing.Size(560, 20)
$grpProgress.Controls.Add($lblStats)

# Control buttons
$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = "Start Processing"
$btnStart.Location = New-Object System.Drawing.Point(20, 430)
$btnStart.Size = New-Object System.Drawing.Size(120, 35)
$btnStart.BackColor = [System.Drawing.Color]::LightGreen
$mainTab.Controls.Add($btnStart)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Cancel"
$btnCancel.Location = New-Object System.Drawing.Point(150, 430)
$btnCancel.Size = New-Object System.Drawing.Size(80, 35)
$btnCancel.Enabled = $false
$mainTab.Controls.Add($btnCancel)

$btnViewLog = New-Object System.Windows.Forms.Button
$btnViewLog.Text = "View Log"
$btnViewLog.Location = New-Object System.Drawing.Point(240, 430)
$btnViewLog.Size = New-Object System.Drawing.Size(80, 35)
$mainTab.Controls.Add($btnViewLog)

$btnOpenOutput = New-Object System.Windows.Forms.Button
$btnOpenOutput.Text = "Open Output"
$btnOpenOutput.Location = New-Object System.Drawing.Point(330, 430)
$btnOpenOutput.Size = New-Object System.Drawing.Size(100, 35)
$mainTab.Controls.Add($btnOpenOutput)

# Configuration Tab
$configTab = New-Object System.Windows.Forms.TabPage
$configTab.Text = "Configuration"
$tabControl.TabPages.Add($configTab)

# API Keys section
$grpApiKeys = New-Object System.Windows.Forms.GroupBox
$grpApiKeys.Text = "API Keys (Optional - for enhanced metadata)"
$grpApiKeys.Location = New-Object System.Drawing.Point(20, 20)
$grpApiKeys.Size = New-Object System.Drawing.Size(590, 150)
$configTab.Controls.Add($grpApiKeys)

$lblLastFm = New-Object System.Windows.Forms.Label
$lblLastFm.Text = "Last.fm API Key:"
$lblLastFm.Location = New-Object System.Drawing.Point(15, 25)
$lblLastFm.Size = New-Object System.Drawing.Size(120, 20)
$grpApiKeys.Controls.Add($lblLastFm)

$txtLastFmKey = New-Object System.Windows.Forms.TextBox
$txtLastFmKey.Location = New-Object System.Drawing.Point(140, 23)
$txtLastFmKey.Size = New-Object System.Drawing.Size(300, 25)
$txtLastFmKey.PasswordChar = '*'
$grpApiKeys.Controls.Add($txtLastFmKey)

$lblDiscogs = New-Object System.Windows.Forms.Label
$lblDiscogs.Text = "Discogs Token:"
$lblDiscogs.Location = New-Object System.Drawing.Point(15, 55)
$lblDiscogs.Size = New-Object System.Drawing.Size(120, 20)
$grpApiKeys.Controls.Add($lblDiscogs)

$txtDiscogsToken = New-Object System.Windows.Forms.TextBox
$txtDiscogsToken.Location = New-Object System.Drawing.Point(140, 53)
$txtDiscogsToken.Size = New-Object System.Drawing.Size(300, 25)
$txtDiscogsToken.PasswordChar = '*'
$grpApiKeys.Controls.Add($txtDiscogsToken)

$lblSpotifyId = New-Object System.Windows.Forms.Label
$lblSpotifyId.Text = "Spotify Client ID:"
$lblSpotifyId.Location = New-Object System.Drawing.Point(15, 85)
$lblSpotifyId.Size = New-Object System.Drawing.Size(120, 20)
$grpApiKeys.Controls.Add($lblSpotifyId)

$txtSpotifyId = New-Object System.Windows.Forms.TextBox
$txtSpotifyId.Location = New-Object System.Drawing.Point(140, 83)
$txtSpotifyId.Size = New-Object System.Drawing.Size(300, 25)
$grpApiKeys.Controls.Add($txtSpotifyId)

$lblSpotifySecret = New-Object System.Windows.Forms.Label
$lblSpotifySecret.Text = "Spotify Secret:"
$lblSpotifySecret.Location = New-Object System.Drawing.Point(15, 115)
$lblSpotifySecret.Size = New-Object System.Drawing.Size(120, 20)
$grpApiKeys.Controls.Add($lblSpotifySecret)

$txtSpotifySecret = New-Object System.Windows.Forms.TextBox
$txtSpotifySecret.Location = New-Object System.Drawing.Point(140, 113)
$txtSpotifySecret.Size = New-Object System.Drawing.Size(300, 25)
$txtSpotifySecret.PasswordChar = '*'
$grpApiKeys.Controls.Add($txtSpotifySecret)

# Organization settings
$grpOrganization = New-Object System.Windows.Forms.GroupBox
$grpOrganization.Text = "File Organization Settings"
$grpOrganization.Location = New-Object System.Drawing.Point(20, 180)
$grpOrganization.Size = New-Object System.Drawing.Size(590, 120)
$configTab.Controls.Add($grpOrganization)

$lblStructure = New-Object System.Windows.Forms.Label
$lblStructure.Text = "Folder Structure:"
$lblStructure.Location = New-Object System.Drawing.Point(15, 25)
$lblStructure.Size = New-Object System.Drawing.Size(120, 20)
$grpOrganization.Controls.Add($lblStructure)

$cmbStructure = New-Object System.Windows.Forms.ComboBox
$cmbStructure.Location = New-Object System.Drawing.Point(140, 23)
$cmbStructure.Size = New-Object System.Drawing.Size(300, 25)
$cmbStructure.DropDownStyle = "DropDownList"
$cmbStructure.Items.AddRange(@(
    "Genre\Artist\Year - Album",
    "Artist\Year - Album",
    "Artist\Album",
    "Year\Artist - Album",
    "Genre\Year\Artist - Album"
))
$cmbStructure.SelectedIndex = 0
$grpOrganization.Controls.Add($cmbStructure)

$lblNaming = New-Object System.Windows.Forms.Label
$lblNaming.Text = "File Naming:"
$lblNaming.Location = New-Object System.Drawing.Point(15, 55)
$lblNaming.Size = New-Object System.Drawing.Size(120, 20)
$grpOrganization.Controls.Add($lblNaming)

$cmbNaming = New-Object System.Windows.Forms.ComboBox
$cmbNaming.Location = New-Object System.Drawing.Point(140, 53)
$cmbNaming.Size = New-Object System.Drawing.Size(300, 25)
$cmbNaming.DropDownStyle = "DropDownList"
$cmbNaming.Items.AddRange(@(
    "Track - Title",
    "Track. Title",
    "Artist - Title",
    "Title",
    "Track - Artist - Title"
))
$cmbNaming.SelectedIndex = 0
$grpOrganization.Controls.Add($cmbNaming)

$btnSaveConfig = New-Object System.Windows.Forms.Button
$btnSaveConfig.Text = "Save Configuration"
$btnSaveConfig.Location = New-Object System.Drawing.Point(20, 320)
$btnSaveConfig.Size = New-Object System.Drawing.Size(150, 35)
$btnSaveConfig.BackColor = [System.Drawing.Color]::LightBlue
$configTab.Controls.Add($btnSaveConfig)

$btnLoadConfig = New-Object System.Windows.Forms.Button
$btnLoadConfig.Text = "Load Configuration"
$btnLoadConfig.Location = New-Object System.Drawing.Point(180, 320)
$btnLoadConfig.Size = New-Object System.Drawing.Size(150, 35)
$configTab.Controls.Add($btnLoadConfig)

# Results Tab
$resultsTab = New-Object System.Windows.Forms.TabPage
$resultsTab.Text = "Results & Statistics"
$tabControl.TabPages.Add($resultsTab)

$txtResults = New-Object System.Windows.Forms.RichTextBox
$txtResults.Location = New-Object System.Drawing.Point(20, 20)
$txtResults.Size = New-Object System.Drawing.Size(720, 400)
$txtResults.ReadOnly = $true
$txtResults.Font = New-Object System.Drawing.Font("Consolas", 9)
$resultsTab.Controls.Add($txtResults)

$btnExportResults = New-Object System.Windows.Forms.Button
$btnExportResults.Text = "Export Results"
$btnExportResults.Location = New-Object System.Drawing.Point(20, 430)
$btnExportResults.Size = New-Object System.Drawing.Size(120, 35)
$resultsTab.Controls.Add($btnExportResults)

$btnClearResults = New-Object System.Windows.Forms.Button
$btnClearResults.Text = "Clear Results"
$btnClearResults.Location = New-Object System.Drawing.Point(150, 430)
$btnClearResults.Size = New-Object System.Drawing.Size(100, 35)
$resultsTab.Controls.Add($btnClearResults)

# Event handlers
$btnBrowseSource.Add_Click({
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderDialog.Description = "Select source music folder"
    $folderDialog.SelectedPath = $txtSource.Text
    
    if ($folderDialog.ShowDialog() -eq "OK") {
        $txtSource.Text = $folderDialog.SelectedPath
        $txtOutput.Text = Join-Path $folderDialog.SelectedPath "Organized"
    }
})

$btnBrowseOutput.Add_Click({
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderDialog.Description = "Select output folder"
    $folderDialog.SelectedPath = $txtOutput.Text
    
    if ($folderDialog.ShowDialog() -eq "OK") {
        $txtOutput.Text = $folderDialog.SelectedPath
    }
})

$btnStart.Add_Click({
    if (-not (Test-Path $txtSource.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Source folder does not exist!", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    $Global:ProcessingCancelled = $false
    $btnStart.Enabled = $false
    $btnCancel.Enabled = $true
    
    # Build parameters
    $params = @{
        SourcePath = $txtSource.Text
        OutputPath = $txtOutput.Text
        DryRun = $chkDryRun.Checked
        EnableWebSearch = $chkWebSearch.Checked
    }
    
    # Start processing in background
    Start-ProcessingJob -Parameters $params
})

$btnCancel.Add_Click({
    $Global:ProcessingCancelled = $true
    $btnCancel.Enabled = $false
    $lblProgress.Text = "Cancelling..."
})

$btnViewLog.Add_Click({
    $logPath = Join-Path $txtSource.Text "MusicLibraryManager.log"
    if (Test-Path $logPath) {
        Start-Process notepad.exe $logPath
    } else {
        [System.Windows.Forms.MessageBox]::Show("Log file not found!", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

$btnOpenOutput.Add_Click({
    if (Test-Path $txtOutput.Text) {
        Start-Process explorer.exe $txtOutput.Text
    } else {
        [System.Windows.Forms.MessageBox]::Show("Output folder does not exist yet!", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

$btnSaveConfig.Add_Click({
    Save-Configuration
})

$btnLoadConfig.Add_Click({
    Load-Configuration
})

$btnExportResults.Add_Click({
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    $saveDialog.FileName = "MusicLibraryResults_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    
    if ($saveDialog.ShowDialog() -eq "OK") {
        $txtResults.Text | Out-File -FilePath $saveDialog.FileName -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("Results exported successfully!", "Export Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

$btnClearResults.Add_Click({
    $txtResults.Clear()
})

# Helper functions
function Start-ProcessingJob {
    param([hashtable]$Parameters)
    
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.Open()
    
    $powershell = [powershell]::Create()
    $powershell.Runspace = $runspace
    
    $scriptBlock = {
        param($SourcePath, $OutputPath, $DryRun, $EnableWebSearch)
        
        # Load the main processing script
        $mainScript = Join-Path $using:PSScriptRoot "MusicLibraryManager.ps1"
        if (Test-Path $mainScript) {
            & $mainScript -SourcePath $SourcePath -OutputPath $OutputPath -DryRun:$DryRun -EnableWebSearch:$EnableWebSearch
        }
    }
    
    $powershell.AddScript($scriptBlock)
    $powershell.AddParameters($Parameters)
    
    $asyncResult = $powershell.BeginInvoke()
    
    # Monitor progress
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 1000
    $timer.Add_Tick({
        if ($asyncResult.IsCompleted -or $Global:ProcessingCancelled) {
            $timer.Stop()
            $btnStart.Enabled = $true
            $btnCancel.Enabled = $false
            
            if ($Global:ProcessingCancelled) {
                $lblProgress.Text = "Processing cancelled by user"
            } else {
                $lblProgress.Text = "Processing completed!"
                $progressBar.Value = 100
                
                # Load results
                Load-ProcessingResults
            }
            
            $powershell.EndInvoke($asyncResult)
            $powershell.Dispose()
            $runspace.Close()
        }
    })
    
    $timer.Start()
}

function Load-ProcessingResults {
    $manifestPath = Join-Path $txtOutput.Text "MusicLibraryManifest.json"
    $summaryPath = Join-Path $txtOutput.Text "ProcessingSummary.txt"
    
    $results = ""
    
    if (Test-Path $summaryPath) {
        $results = Get-Content $summaryPath -Raw
    } elseif (Test-Path $manifestPath) {
        try {
            $manifest = Get-Content $manifestPath | ConvertFrom-Json
            $results = "Processing Results`n" +
                      "==================`n" +
                      "Total Files: $($manifest.TotalFiles)`n" +
                      "Processed: $($manifest.ProcessedFiles)`n" +
                      "Corrected: $($manifest.CorrectedFiles)`n" +
                      "Artists: $($manifest.Artists.Count)`n" +
                      "Albums: $($manifest.Albums.Count)`n" +
                      "Genres: $($manifest.Genres.Count)`n" +
                      "Errors: $($manifest.Errors.Count)`n"
        }
        catch {
            $results = "Error loading results: $($_.Exception.Message)"
        }
    } else {
        $results = "No results available yet. Run processing first."
    }
    
    $txtResults.Text = $results
    $tabControl.SelectedTab = $resultsTab
}

function Save-Configuration {
    try {
        $config.SearchProviders.LastFm.ApiKey = $txtLastFmKey.Text
        $config.SearchProviders.Discogs.Token = $txtDiscogsToken.Text
        $config.SearchProviders.Spotify.ClientId = $txtSpotifyId.Text
        $config.SearchProviders.Spotify.ClientSecret = $txtSpotifySecret.Text
        $config.FileOrganization.Structure = $cmbStructure.SelectedItem
        $config.FileOrganization.FileNaming = $cmbNaming.SelectedItem
        
        $config | ConvertTo-Json -Depth 10 | Out-File -FilePath $configPath -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("Configuration saved successfully!", "Save Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error saving configuration: $($_.Exception.Message)", "Save Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

function Load-Configuration {
    if (Test-Path $configPath) {
        try {
            $loadedConfig = Get-Content $configPath | ConvertFrom-Json -AsHashtable
            
            if ($loadedConfig.SearchProviders.LastFm.ApiKey) {
                $txtLastFmKey.Text = $loadedConfig.SearchProviders.LastFm.ApiKey
            }
            if ($loadedConfig.SearchProviders.Discogs.Token) {
                $txtDiscogsToken.Text = $loadedConfig.SearchProviders.Discogs.Token
            }
            if ($loadedConfig.SearchProviders.Spotify.ClientId) {
                $txtSpotifyId.Text = $loadedConfig.SearchProviders.Spotify.ClientId
            }
            if ($loadedConfig.SearchProviders.Spotify.ClientSecret) {
                $txtSpotifySecret.Text = $loadedConfig.SearchProviders.Spotify.ClientSecret
            }
            if ($loadedConfig.FileOrganization.Structure) {
                $cmbStructure.SelectedItem = $loadedConfig.FileOrganization.Structure
            }
            if ($loadedConfig.FileOrganization.FileNaming) {
                $cmbNaming.SelectedItem = $loadedConfig.FileOrganization.FileNaming
            }
            
            [System.Windows.Forms.MessageBox]::Show("Configuration loaded successfully!", "Load Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Error loading configuration: $($_.Exception.Message)", "Load Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
}

# Load initial configuration
Load-Configuration

# Show the form
[System.Windows.Forms.Application]::EnableVisualStyles()
$form.ShowDialog()