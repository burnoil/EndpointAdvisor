# PowerShell GUI for editing JSON content in a Git repository
# Requires Git and WinGet for Git installation if not present

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# Initialize logging
$logFile = Join-Path $PSScriptRoot "JsonEditorGui.log"
function Write-Log {
    param ($Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] $Message"
    Write-Host $logMessage
    Add-Content -Path $logFile -Value $logMessage -ErrorAction SilentlyContinue
}

Write-Log "Starting JsonEditorGui.ps1"

# Function to check if Git is installed
function Test-GitInstalled {
    try {
        git --version | Out-Null
        Write-Log "Git is installed."
        return $true
    }
    catch {
        Write-Log "Git not found: $($_.Exception.Message)"
        return $false
    }
}

# Function to install Git using WinGet
function Install-Git {
    param ($StatusTextBlock)
    try {
        $StatusTextBlock.Text = "Installing Git via WinGet..."
        Write-Log "Installing Git via WinGet..."
        winget install --id Git.Git --silent --accept-package-agreements --accept-source-agreements | ForEach-Object {
            Write-Log "WinGet output: $_"
            $StatusTextBlock.Text = "Installing Git: $_"
            [System.Windows.Forms.Application]::DoEvents() # Update GUI
        }
        # Refresh environment to ensure git is available
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
        if (Test-GitInstalled) {
            $StatusTextBlock.Text = "Git installed successfully."
            Write-Log "Git installed successfully."
            return $true
        }
        else {
            $StatusTextBlock.Text = "Git installation completed but git command not found."
            Write-Log "Git installation completed but git command not found."
            return $false
        }
    }
    catch {
        $StatusTextBlock.Text = "Failed to install Git: $($_.Exception.Message)"
        Write-Log "Failed to install Git: $($_.Exception.Message)"
        return $false
    }
}

# Function to ensure WinGet is available
function Ensure-WinGet {
    param ($StatusTextBlock)
    try {
        winget --version | Out-Null
        Write-Log "WinGet is available."
        return $true
    }
    catch {
        try {
            $StatusTextBlock.Text = "Installing WinGet..."
            Write-Log "Installing WinGet..."
            $progressPreference = 'SilentlyContinue'
            Install-PackageProvider -Name NuGet -Force | Out-Null
            Install-Module -Name Microsoft.WinGet.Client -Force -Repository PSGallery | Out-Null
            Repair-WinGetPackageManager | Out-Null
            $StatusTextBlock.Text = "WinGet installed successfully."
            Write-Log "WinGet installed successfully."
            return $true
        }
        catch {
            $StatusTextBlock.Text = "Failed to install WinGet: $($_.Exception.Message)"
            Write-Log "Failed to install WinGet: $($_.Exception.Message)"
            return $false
        }
    }
}

# Function to configure Git user identity
function Configure-GitIdentity {
    try {
        Write-Log "Checking Git user identity..."
        $currentName = git config user.name
        $currentEmail = git config user.email
        if (-not $currentName -or -not $currentEmail) {
            Write-Log "Configuring Git user identity."
            git config --global user.name "JsonEditorUser" 2>&1 | ForEach-Object { Write-Log "Git config user.name: $_" }
            git config --global user.email "jsoneditor@example.com" 2>&1 | ForEach-Object { Write-Log "Git config user.email: $_" }
            Write-Log "Git user identity configured."
        }
        else {
            Write-Log "Git user identity already configured: $currentName <$currentEmail>"
        }
        return $true
    }
    catch {
        Write-Log "Failed to configure Git user identity: $($_.Exception.Message)"
        return $false
    }
}

# Function to download JSON file
function Get-JsonContent {
    param ($Url, $LocalFilePath)
    try {
        if (-not $Url) {
            throw "URL is null or empty."
        }
        Write-Log "Attempting to download JSON from $Url"
        # Add cache-busting query parameter
        $timestamp = [DateTime]::Now.ToFileTimeUtc()
        $fetchUrl = "$Url?t=$timestamp"
        Write-Log "Constructed URL: $fetchUrl"
        $headers = @{
            "Cache-Control" = "no-cache"
            "Pragma" = "no-cache"
        }
        # Validate URL
        try {
            $uri = [System.Uri]$fetchUrl
            Write-Log "URL validated successfully: $uri"
        }
        catch {
            Write-Log "Invalid URL format: $($_.Exception.Message)"
            throw "Invalid URL: $($_.Exception.Message)"
        }
        $response = Invoke-WebRequest -Uri $fetchUrl -UseBasicParsing -Headers $headers
        $jsonContent = $response.Content | ConvertFrom-Json
        # Log a snippet of the fetched JSON for verification
        $jsonSnippet = $response.Content.Substring(0, [Math]::Min(200, $response.Content.Length))
        Write-Log "Fetched JSON snippet: $jsonSnippet"
        # Validate Links arrays
        Write-Log "Announcements Links: $($jsonContent.Announcements.Links | ConvertTo-Json -Compress)"
        Write-Log "Support Links: $($jsonContent.Support.Links | ConvertTo-Json -Compress)"
        if ($null -eq $jsonContent.Announcements.Links) { $jsonContent.Announcements.Links = @() }
        if ($null -eq $jsonContent.Support.Links) { $jsonContent.Support.Links = @() }
        # Exclude EarlyAdopter
        $filteredContent = [PSCustomObject]@{
            Announcements = [PSCustomObject]@{
                Text = $jsonContent.Announcements.Text
                Details = $jsonContent.Announcements.Details
                Links = @($jsonContent.Announcements.Links | Where-Object { $_.Name -and $_.Url })
            }
            Support = [PSCustomObject]@{
                Text = $jsonContent.Support.Text
                Links = @($jsonContent.Support.Links | Where-Object { $_.Name -and $_.Url })
            }
        }
        Write-Log "JSON downloaded and filtered successfully."
        return $filteredContent
    }
    catch {
        Write-Log "Failed to download JSON from $Url - Error: $($_.Exception.Message)"
        # Fallback to local file
        try {
            Write-Log "Attempting to read JSON from local file: $LocalFilePath"
            if (Test-Path $LocalFilePath) {
                $localRaw = Get-Content -Path $LocalFilePath -Raw
                $jsonContent = $localRaw | ConvertFrom-Json
                # Validate Links arrays
                Write-Log "Announcements Links (local): $($jsonContent.Announcements.Links | ConvertTo-Json -Compress)"
                Write-Log "Support Links (local): $($jsonContent.Support.Links | ConvertTo-Json -Compress)"
                if ($null -eq $jsonContent.Announcements.Links) { $jsonContent.Announcements.Links = @() }
                if ($null -eq $jsonContent.Support.Links) { $jsonContent.Support.Links = @() }
                $filteredContent = [PSCustomObject]@{
                    Announcements = [PSCustomObject]@{
                        Text = $jsonContent.Announcements.Text
                        Details = $jsonContent.Announcements.Details
                        Links = @($jsonContent.Announcements.Links | Where-Object { $_.Name -and $_.Url })
                    }
                    Support = [PSCustomObject]@{
                        Text = $jsonContent.Support.Text
                        Links = @($jsonContent.Support.Links | Where-Object { $_.Name -and $_.Url })
                    }
                }
                Write-Log "Local JSON read and filtered successfully."
                return $filteredContent
            }
            else {
                Write-Log "Local JSON file not found: $LocalFilePath"
                throw "Local JSON file not found."
            }
        }
        catch {
            Write-Log "Failed to read local JSON: $($_.Exception.Message)"
            throw "Failed to fetch JSON: $($_.Exception.Message)"
        }
    }
}

# Function to clone or pull repository
function Sync-Repository {
    param ($RepoUrl, $LocalPath, $Branch = "main")
    try {
        if (Test-Path $LocalPath) {
            Write-Log "Pulling repository at $LocalPath"
            Set-Location $LocalPath
            git pull origin $Branch 2>&1 | ForEach-Object { Write-Log "Git pull: $_" }
        }
        else {
            Write-Log "Cloning repository from $RepoUrl to $LocalPath"
            git clone $RepoUrl $LocalPath 2>&1 | ForEach-Object { Write-Log "Git clone: $_" }
            Set-Location $LocalPath
        }
        # Return to neutral directory
        Set-Location $env:TEMP
        Write-Log "Repository synced successfully."
        return $true
    }
    catch {
        Write-Log "Failed to sync repository: $($_.Exception.Message)"
        Set-Location $env:TEMP
        return $false
    }
}

# Function to sanitize strings for JSON
function Sanitize-JsonString {
    param ($InputString)
    if ($null -eq $InputString) { return "" }
    # Escape quotes and control characters, but preserve URL characters
    $sanitized = $InputString -replace '"', '\"' -replace '\r\n', '\n' -replace '\r', '\n' -replace '\t', '\t'
    return $sanitized
}

# Function to format JSON with exact 2-space indentation
function Format-Json {
    param ($JsonString)
    $lines = $JsonString -split '\n'
    $result = @()
    $indentLevel = 0
    foreach ($line in $lines) {
        $trimmed = $line.Trim()
        if ($trimmed -eq '' -or $trimmed -eq '}' -or $trimmed -eq ']' -or $trimmed -eq '},' -or $trimmed -eq '],') {
            $indentLevel = [Math]::Max(0, $indentLevel - 1)
        }
        $result += ('  ' * $indentLevel) + $trimmed
        if ($trimmed -eq '{' -or $trimmed -eq '[' -or $trimmed -match '^[^:]+:\s*[{[]$') {
            $indentLevel++
        }
    }
    return $result -join "`n"
}

# Function to commit and push changes
function Save-Changes {
    param ($LocalPath, $FilePath, $JsonContent, $CommitMessage)
    try {
        Write-Log "Saving changes to $FilePath"
        Set-Location $LocalPath

        # Create updated JSON object, excluding EarlyAdopter
        $updatedJson = [PSCustomObject]@{
            Announcements = [PSCustomObject]@{
                Text = Sanitize-JsonString $JsonContent.Announcements.Text
                Details = Sanitize-JsonString $JsonContent.Announcements.Details
                Links = @($JsonContent.Announcements.Links | Where-Object { $_.Name -and $_.Url } | ForEach-Object { 
                    [PSCustomObject]@{ 
                        Name = Sanitize-JsonString $_.Name
                        Url = $_.Url # URLs are preserved as-is
                    } 
                })
            }
            Support = [PSCustomObject]@{
                Text = Sanitize-JsonString $JsonContent.Support.Text
                Links = @($JsonContent.Support.Links | Where-Object { $_.Name -and $_.Url } | ForEach-Object { 
                    [PSCustomObject]@{ 
                        Name = Sanitize-JsonString $_.Name
                        Url = $_.Url # URLs are preserved as-is
                    } 
                })
            }
        }

        # Serialize with minimal formatting
        $updatedRaw = $updatedJson | ConvertTo-Json -Depth 100 -Compress:$false
        # Fix escaping and format with exact 2-space indentation
        $formattedRaw = $updatedRaw -replace '\\u0026', '&' -replace '\\u003d', '=' -replace '\\u0027', "'"
        $formattedRaw = Format-Json -JsonString $formattedRaw
        Write-Log "Generated JSON content: $($formattedRaw.Substring(0, [Math]::Min(200, $formattedRaw.Length)))"
        
        # Validate JSON
        try {
            $null = $formattedRaw | ConvertFrom-Json
            Write-Log "Generated JSON is valid."
        }
        catch {
            Write-Log "Generated JSON is invalid: $($_.Exception.Message)"
            throw "Invalid JSON generated: $($_.Exception.Message)"
        }

        # Write to file without BOM
        Write-Log "Writing updated JSON to file."
        [System.IO.File]::WriteAllText($FilePath, $formattedRaw, [System.Text.UTF8Encoding]::new($false))
        Write-Log "JSON file updated locally."

        # Log raw file content for verification
        $fileContent = Get-Content -Path $FilePath -Raw
        Write-Log "Raw file content (first 200 chars): $($fileContent.Substring(0, [Math]::Min(200, $fileContent.Length)))"

        # Stage changes
        git add $FilePath 2>&1 | ForEach-Object { Write-Log "Git add: $_" }

        # Attempt commit
        $commitOutput = git commit -m $CommitMessage 2>&1
        $commitOutput | ForEach-Object { Write-Log "Git commit: $_" }
        if ($commitOutput -match "fatal:") {
            throw "Git commit failed: $commitOutput"
        }

        # Push changes
        git push origin main 2>&1 | ForEach-Object { Write-Log "Git push: $_" }
        
        # Return to neutral directory
        Set-Location $env:TEMP
        Write-Log "Changes saved and pushed successfully."
        return $true
    }
    catch {
        Write-Log "Failed to save changes: $($_.Exception.Message)"
        Set-Location $env:TEMP
        return $false
    }
}

# Function to safely remove directory with retries
function Remove-DirectorySafely {
    param ($Path, $MaxRetries = 3, $RetryDelaySeconds = 1)
    $attempt = 0
    while ($attempt -lt $MaxRetries) {
        try {
            if (Test-Path $Path) {
                Write-Log "Attempting to remove directory $Path (Attempt $($attempt + 1))"
                Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
                Write-Log "Directory $Path removed successfully."
            }
            return $true
        }
        catch {
            Write-Log "Attempt $($attempt + 1) to remove directory failed: $($_.Exception.Message)"
            $attempt++
            if ($attempt -lt $MaxRetries) {
                Start-Sleep -Seconds $RetryDelaySeconds
            }
        }
    }
    Write-Log "Failed to remove directory $Path after $MaxRetries attempts."
    return $false
}

# Function to convert JSON to DataGrid items for Links
function Convert-LinksToDataGridItems {
    param ($Links)
    try {
        Write-Log "Converting Links to DataGrid items: $($Links | ConvertTo-Json -Compress)"
        # Ensure output is always an array, even for empty or single-item inputs
        $result = @($Links | Where-Object { $_.Name -and $_.Url } | ForEach-Object { 
            [PSCustomObject]@{ Name = $_.Name; Url = $_.Url } 
        })
        Write-Log "Converted Links: $($result | ConvertTo-Json -Compress)"
        return $result
    }
    catch {
        Write-Log "Failed to convert links to DataGrid items: $($_.Exception.Message)"
        return @()
    }
}

# Function to convert DataGrid items back to Links
function Convert-DataGridItemsToLinks {
    param ($Items)
    try {
        Write-Log "Converting DataGrid items to Links: $($Items | ConvertTo-Json -Compress)"
        # Filter out incomplete rows and ensure array output
        $result = @($Items | Where-Object { $_.Name -and $_.Url } | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                Url = $_.Url
            }
        })
        Write-Log "Converted DataGrid items: $($result | ConvertTo-Json -Compress)"
        return $result
    }
    catch {
        Write-Log "Failed to convert DataGrid items to links: $($_.Exception.Message)"
        return @()
    }
}

# Function to validate DataGrid rows
function Validate-DataGridRows {
    param ($Grid)
    foreach ($item in $Grid.Items) {
        if (-not $item.Name -or -not $item.Url) {
            return $false
        }
    }
    return $true
}

# WPF XAML for GUI with colored sections
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="LLNOTIFY JSON Content Editor" Height="700" Width="800">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="5">
            <TextBlock Text="GitHub Token:" VerticalAlignment="Center" Margin="5"/>
            <PasswordBox x:Name="TokenBox" Width="200" Margin="5"/>
            <TextBlock x:Name="RepoTextBlock" Text="Repository: Not connected" VerticalAlignment="Center" Margin="5" FontStyle="Italic"/>
        </StackPanel>
        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="5">
            <Button x:Name="LoadButton" Content="Load JSON" Width="100" Margin="5"/>
            <Button x:Name="RefreshButton" Content="Refresh" Width="100" Margin="5"/>
            <Button x:Name="ResetButton" Content="Reset" Width="100" Margin="5"/>
            <Button x:Name="ToggleViewButton" Content="Show Raw JSON" Width="120" Margin="5"/>
        </StackPanel>
        <TextBlock x:Name="StatusTextBlock" Grid.Row="2" Margin="5" Text="Checking Git status..." FontStyle="Italic"/>
        <Grid Grid.Row="3" x:Name="ContentGrid" Margin="5">
            <ScrollViewer x:Name="SectionView" VerticalScrollBarVisibility="Auto" Visibility="Visible">
                <StackPanel>
                    <Expander Header="Announcements" IsExpanded="True" Margin="5">
                        <StackPanel Margin="10" Background="#E6F0FA">
                            <TextBlock Text="Text:" Margin="0,0,0,5"/>
                            <TextBox x:Name="AnnouncementsTextBox" AcceptsReturn="True" AcceptsTab="True" Height="80" TextWrapping="Wrap"/>
                            <TextBlock Text="Details:" Margin="0,10,0,5"/>
                            <TextBox x:Name="AnnouncementsDetailsBox" AcceptsReturn="True" AcceptsTab="True" Height="80" TextWrapping="Wrap"/>
                            <TextBlock Text="Links:" Margin="0,10,0,5"/>
                            <DataGrid x:Name="AnnouncementsLinksGrid" AutoGenerateColumns="False" CanUserAddRows="True" CanUserDeleteRows="True" Height="100">
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="*"/>
                                    <DataGridTextColumn Header="Url" Binding="{Binding Url}" Width="*"/>
                                </DataGrid.Columns>
                            </DataGrid>
                        </StackPanel>
                    </Expander>
                    <Expander Header="Support" IsExpanded="True" Margin="5">
                        <StackPanel Margin="10" Background="#E6FAE6">
                            <TextBlock Text="Text:" Margin="0,0,0,5"/>
                            <TextBox x:Name="SupportTextBox" AcceptsReturn="True" AcceptsTab="True" Height="80" TextWrapping="Wrap"/>
                            <TextBlock Text="Links:" Margin="0,10,0,5"/>
                            <DataGrid x:Name="SupportLinksGrid" AutoGenerateColumns="False" CanUserAddRows="True" CanUserDeleteRows="True" Height="100">
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="*"/>
                                    <DataGridTextColumn Header="Url" Binding="{Binding Url}" Width="*"/>
                                </DataGrid.Columns>
                            </DataGrid>
                        </StackPanel>
                    </Expander>
                </StackPanel>
            </ScrollViewer>
            <TextBox x:Name="JsonTextBox" AcceptsReturn="True" AcceptsTab="True" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" Visibility="Collapsed"/>
        </Grid>
        <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right" Margin="5">
            <Button x:Name="SaveButton" Content="Save Changes" Width="100" Margin="5"/>
            <Button x:Name="CancelButton" Content="Cancel" Width="100" Margin="5"/>
        </StackPanel>
    </Grid>
</Window>
"@

# Load XAML
try {
    Write-Log "Loading XAML"
    $reader = (New-Object System.Xml.XmlNodeReader ([xml]$xaml))
    $window = [Windows.Markup.XamlReader]::Load($reader)
    Write-Log "XAML loaded successfully."
}
catch {
    Write-Log "Failed to load XAML: $($_.Exception.Message)"
    [System.Windows.MessageBox]::Show("Failed to load GUI: $($_.Exception.Message). Check the log at $logFile for details.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    exit
}

# Get controls
$tokenBox = $window.FindName("TokenBox")
$loadButton = $window.FindName("LoadButton")
$refreshButton = $window.FindName("RefreshButton")
$resetButton = $window.FindName("ResetButton")
$toggleViewButton = $window.FindName("ToggleViewButton")
$statusTextBlock = $window.FindName("StatusTextBlock")
$repoTextBlock = $window.FindName("RepoTextBlock")
$sectionView = $window.FindName("SectionView")
$announcementsTextBox = $window.FindName("AnnouncementsTextBox")
$announcementsDetailsBox = $window.FindName("AnnouncementsDetailsBox")
$announcementsLinksGrid = $window.FindName("AnnouncementsLinksGrid")
$supportTextBox = $window.FindName("SupportTextBox")
$supportLinksGrid = $window.FindName("SupportLinksGrid")
$jsonTextBox = $window.FindName("JsonTextBox")
$saveButton = $window.FindName("SaveButton")
$cancelButton = $window.FindName("CancelButton")

# Repository details
$repoUrl = "https://github.com/burnoil/LLNOTIFY.git"
$jsonUrl = "https://raw.githubusercontent.com/burnoil/LLNOTIFY/main/ContentData.json"
$localPath = Join-Path $env:TEMP "LLNOTIFY"
$filePath = Join-Path $localPath "ContentData.json"

# State variables
$isRawView = $false
$lastLoadedJson = $null

# Check Git status on startup
if (Test-GitInstalled) {
    $statusTextBlock.Text = "Git is installed."
}
else {
    $statusTextBlock.Text = "Git not found. Will install when loading JSON."
}

# Event handlers
$loadButton.Add_Click({
    # Check for WinGet and Git
    if (-not (Ensure-WinGet -StatusTextBlock $statusTextBlock)) {
        Write-Log "WinGet installation failed."
        [System.Windows.MessageBox]::Show("WinGet installation failed. Cannot proceed. Check the log at $logFile for details.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return
    }

    if (-not (Test-GitInstalled)) {
        if (-not (Install-Git -StatusTextBlock $statusTextBlock)) {
            Write-Log "Git installation failed."
            [System.Windows.MessageBox]::Show("Git installation failed. Cannot proceed. Check the log at $logFile for details.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            return
        }
    }

    # Configure Git identity
    $statusTextBlock.Text = "Configuring Git user identity..."
    if (-not (Configure-GitIdentity)) {
        Write-Log "Git identity configuration failed."
        $statusTextBlock.Text = "Error: Failed to configure Git identity."
        [System.Windows.MessageBox]::Show("Failed to configure Git user identity. Check the log at $logFile for details.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return
    }

    # Set Git credential helper for token
    $token = $tokenBox.Password
    if (-not $token) {
        Write-Log "GitHub token not provided."
        $statusTextBlock.Text = "Error: Please enter a GitHub token."
        [System.Windows.MessageBox]::Show("Please enter a GitHub personal access token.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return
    }

    # Configure Git to use token
    Write-Log "Configuring Git credential helper."
    $statusTextBlock.Text = "Configuring Git authentication..."
    $env:GIT_ASKPASS = $null
    git config --global credential.helper "!f() { echo 'username=x-access-token'; echo 'password=$token'; }; f" 2>&1 | ForEach-Object { Write-Log "Git config: $_" }

    # Sync repository
    $statusTextBlock.Text = "Syncing repository..."
    if (-not (Sync-Repository -RepoUrl $repoUrl -LocalPath $localPath)) {
        Write-Log "Repository sync failed."
        $statusTextBlock.Text = "Error: Failed to sync repository."
        [System.Windows.MessageBox]::Show("Failed to sync repository. Check token permissions or repository access. See log at $logFile.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return
    }

    # Update repository display
    $repoTextBlock.Text = "Repository: $repoUrl"

    # Load JSON content
    $statusTextBlock.Text = "Loading JSON content..."
    $script:lastLoadedJson = Get-JsonContent -Url $jsonUrl -LocalFilePath $filePath
    if ($null -eq $script:lastLoadedJson) {
        Write-Log "JSON content load failed."
        $statusTextBlock.Text = "Error: Failed to load JSON content."
        [System.Windows.MessageBox]::Show("Failed to load JSON content. See log at $logFile.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return
    }

    # Populate section view
    try {
        Write-Log "Populating section view."
        $announcementsTextBox.Text = $script:lastLoadedJson.Announcements.Text
        $announcementsDetailsBox.Text = $script:lastLoadedJson.Announcements.Details
        $announcementsLinksGrid.ItemsSource = Convert-LinksToDataGridItems -Links $script:lastLoadedJson.Announcements.Links
        $supportTextBox.Text = $script:lastLoadedJson.Support.Text
        $supportLinksGrid.ItemsSource = Convert-LinksToDataGridItems -Links $script:lastLoadedJson.Support.Links
        # Populate TextBox for raw view
        $jsonTextBox.Text = $script:lastLoadedJson | ConvertTo-Json -Depth 100
        Write-Log "Section view populated successfully."
        $statusTextBlock.Text = "JSON content loaded successfully."
    }
    catch {
        Write-Log "Failed to populate section view: $($_.Exception.Message)"
        $statusTextBlock.Text = "Error: Failed to process JSON."
        [System.Windows.MessageBox]::Show("Failed to process JSON for section view. Try raw view. See log at $logFile.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return
    }

    $loadButton.IsEnabled = $false
    $refreshButton.IsEnabled = $true
    $resetButton.IsEnabled = $true
    $saveButton.IsEnabled = $true
    $toggleViewButton.IsEnabled = $true
})

$refreshButton.Add_Click({
    # Sync repository
    $statusTextBlock.Text = "Syncing repository..."
    if (-not (Sync-Repository -RepoUrl $repoUrl -LocalPath $localPath)) {
        Write-Log "Repository sync failed during refresh."
        $statusTextBlock.Text = "Error: Failed to sync repository."
        [System.Windows.MessageBox]::Show("Failed to sync repository. Check token permissions or repository access. See log at $logFile.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return
    }

    # Update repository display
    $repoTextBlock.Text = "Repository: $repoUrl"

    # Load JSON content
    $statusTextBlock.Text = "Refreshing JSON content..."
    $script:lastLoadedJson = Get-JsonContent -Url $jsonUrl -LocalFilePath $filePath
    if ($null -eq $script:lastLoadedJson) {
        Write-Log "JSON content refresh failed."
        $statusTextBlock.Text = "Error: Failed to refresh JSON content."
        [System.Windows.MessageBox]::Show("Failed to refresh JSON content. See log at $logFile.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return
    }

    # Populate section view
    try {
        Write-Log "Populating section view after refresh."
        $announcementsTextBox.Text = $script:lastLoadedJson.Announcements.Text
        $announcementsDetailsBox.Text = $script:lastLoadedJson.Announcements.Details
        $announcementsLinksGrid.ItemsSource = Convert-LinksToDataGridItems -Links $script:lastLoadedJson.Announcements.Links
        $supportTextBox.Text = $script:lastLoadedJson.Support.Text
        $supportLinksGrid.ItemsSource = Convert-LinksToDataGridItems -Links $script:lastLoadedJson.Support.Links
        # Populate TextBox for raw view
        $jsonTextBox.Text = $script:lastLoadedJson | ConvertTo-Json -Depth 100
        Write-Log "Section view refreshed successfully."
        $statusTextBlock.Text = "JSON content refreshed successfully."
    }
    catch {
        Write-Log "Failed to populate section view after refresh: $($_.Exception.Message)"
        $statusTextBlock.Text = "Error: Failed to process refreshed JSON."
        [System.Windows.MessageBox]::Show("Failed to process refreshed JSON for section view. Try raw view. See log at $logFile.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
})

$resetButton.Add_Click({
    try {
        Write-Log "Resetting GUI to last loaded JSON."
        $statusTextBlock.Text = "Resetting fields..."
        if ($null -eq $script:lastLoadedJson) {
            Write-Log "No JSON loaded to reset to."
            $statusTextBlock.Text = "Error: No JSON loaded."
            [System.Windows.MessageBox]::Show("No JSON has been loaded to reset to.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            return
        }

        # Populate section view
        $announcementsTextBox.Text = $script:lastLoadedJson.Announcements.Text
        $announcementsDetailsBox.Text = $script:lastLoadedJson.Announcements.Details
        $announcementsLinksGrid.ItemsSource = Convert-LinksToDataGridItems -Links $script:lastLoadedJson.Announcements.Links
        $supportTextBox.Text = $script:lastLoadedJson.Support.Text
        $supportLinksGrid.ItemsSource = Convert-LinksToDataGridItems -Links $script:lastLoadedJson.Support.Links
        # Populate TextBox for raw view
        $jsonTextBox.Text = $script:lastLoadedJson | ConvertTo-Json -Depth 100
        Write-Log "GUI reset successfully."
        $statusTextBlock.Text = "Fields reset to last loaded JSON."
    }
    catch {
        Write-Log "Failed to reset GUI: $($_.Exception.Message)"
        $statusTextBlock.Text = "Error: Failed to reset fields."
        [System.Windows.MessageBox]::Show("Failed to reset fields: $($_.Exception.Message). See log at $logFile.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
})

$toggleViewButton.Add_Click({
    $script:isRawView = -not $script:isRawView
    if ($script:isRawView) {
        Write-Log "Switching to raw JSON view."
        $sectionView.Visibility = [System.Windows.Visibility]::Collapsed
        $jsonTextBox.Visibility = [System.Windows.Visibility]::Visible
        $toggleViewButton.Content = "Show Section View"
        # Sync section view to TextBox
        try {
            $jsonContent = [PSCustomObject]@{
                Announcements = [PSCustomObject]@{
                    Text = $announcementsTextBox.Text
                    Details = $announcementsDetailsBox.Text
                    Links = Convert-DataGridItemsToLinks -Items $announcementsLinksGrid.Items
                }
                Support = [PSCustomObject]@{
                    Text = $supportTextBox.Text
                    Links = Convert-DataGridItemsToLinks -Items $supportLinksGrid.Items
                }
            }
            $jsonTextBox.Text = $jsonContent | ConvertTo-Json -Depth 100
            Write-Log "Synced section view to raw JSON."
            $statusTextBlock.Text = "Switched to raw JSON view."
        }
        catch {
            Write-Log "Failed to sync section view to raw JSON: $($_.Exception.Message)"
            $statusTextBlock.Text = "Error: Failed to sync to raw JSON."
            [System.Windows.MessageBox]::Show("Failed to sync section data to JSON. See log at $logFile.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    }
    else {
        Write-Log "Switching to section view."
        $sectionView.Visibility = [System.Windows.Visibility]::Visible
        $jsonTextBox.Visibility = [System.Windows.Visibility]::Collapsed
        $toggleViewButton.Content = "Show Raw JSON"
        # Sync TextBox to section view
        try {
            $jsonContent = $jsonTextBox.Text | ConvertFrom-Json
            $announcementsTextBox.Text = $jsonContent.Announcements.Text
            $announcementsDetailsBox.Text = $jsonContent.Announcements.Details
            $announcementsLinksGrid.ItemsSource = Convert-LinksToDataGridItems -Links $jsonContent.Announcements.Links
            $supportTextBox.Text = $jsonContent.Support.Text
            $supportLinksGrid.ItemsSource = Convert-LinksToDataGridItems -Links $jsonContent.Support.Links
            Write-Log "Synced raw JSON to section view."
            $statusTextBlock.Text = "Switched to section view."
        }
        catch {
            Write-Log "Failed to sync raw JSON to section view: $($_.Exception.Message)"
            $statusTextBlock.Text = "Error: Invalid JSON in raw view."
            [System.Windows.MessageBox]::Show("Invalid JSON in raw view. Cannot switch to section view. See log at $logFile.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            $script:isRawView = $true
            $sectionView.Visibility = [System.Windows.Visibility]::Collapsed
            $jsonTextBox.Visibility = [System.Windows.Visibility]::Visible
            $toggleViewButton.Content = "Show Section View"
        }
    }
})

$saveButton.Add_Click({
    try {
        Write-Log "Attempting to save changes."
        $statusTextBlock.Text = "Validating changes..."

        # Validate DataGrid rows
        if (-not (Validate-DataGridRows -Grid $announcementsLinksGrid)) {
            Write-Log "Validation failed: Incomplete rows in Announcements Links."
            $statusTextBlock.Text = "Error: Incomplete rows in Announcements Links."
            [System.Windows.MessageBox]::Show("Please ensure all rows in Announcements Links have a Name and URL.", "Validation Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            return
        }
        if (-not (Validate-DataGridRows -Grid $supportLinksGrid)) {
            Write-Log "Validation failed: Incomplete rows in Support Links."
            $statusTextBlock.Text = "Error: Incomplete rows in Support Links."
            [System.Windows.MessageBox]::Show("Please ensure all rows in Support Links have a Name and URL.", "Validation Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            return
        }

        $statusTextBlock.Text = "Saving changes..."
        $jsonContent = if ($script:isRawView) {
            $jsonTextBox.Text | ConvertFrom-Json
        }
        else {
            [PSCustomObject]@{
                Announcements = [PSCustomObject]@{
                    Text = $announcementsTextBox.Text
                    Details = $announcementsDetailsBox.Text
                    Links = Convert-DataGridItemsToLinks -Items $announcementsLinksGrid.Items
                }
                Support = [PSCustomObject]@{
                    Text = $supportTextBox.Text
                    Links = Convert-DataGridItemsToLinks -Items $supportLinksGrid.Items
                }
            }
        }

        if (Save-Changes -LocalPath $localPath -FilePath $filePath -JsonContent $jsonContent -CommitMessage "Updated ContentData.json via GUI") {
            Write-Log "Changes saved successfully."
            $statusTextBlock.Text = "Changes saved successfully."
            # Update last loaded JSON to reflect saved changes
            $script:lastLoadedJson = $jsonContent
            [System.Windows.MessageBox]::Show("Changes saved successfully.", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        }
        else {
            Write-Log "Save operation returned false."
            $statusTextBlock.Text = "Error: Failed to save changes."
            [System.Windows.MessageBox]::Show("Failed to save changes. Check the log at $logFile for details.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    }
    catch {
        Write-Log "Failed to process JSON for saving: $($_.Exception.Message)"
        $statusTextBlock.Text = "Error: Invalid JSON or data error."
        [System.Windows.MessageBox]::Show("Invalid JSON format or data error: $($_.Exception.Message). Check the log at $logFile for details.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
})

$cancelButton.Add_Click({
    Write-Log "User cancelled operation."
    $statusTextBlock.Text = "Operation cancelled."
    $window.Close()
})

# Initialize UI
$saveButton.IsEnabled = $false
$resetButton.IsEnabled = $false
$toggleViewButton.IsEnabled = $false
$refreshButton.IsEnabled = $false

# Show the window
try {
    Write-Log "Showing GUI window."
    $window.ShowDialog() | Out-Null
    Write-Log "GUI window closed."
}
catch {
    Write-Log "Failed to show GUI: $($_.Exception.Message)"
}

# Cleanup
Write-Log "Performing cleanup."
Set-Location $env:TEMP
if (Test-Path $localPath) {
    Remove-DirectorySafely -Path $localPath
}
if (Test-GitInstalled) {
    git config --global --unset credential.helper 2>&1 | ForEach-Object { Write-Log "Git config unset: $_" }
}
Write-Log "Script execution completed."