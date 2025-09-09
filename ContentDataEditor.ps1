# ContentData JSON Editor for Lincoln Laboratory Endpoint Advisor
# Version 3.5 - Modified for GitHub Enterprise Server support with configurable API base URL
# Built for editing JSON from a user-specified repository (default: https://raw.servername/EndpointEngineering/EndpointAdvisor/main/ContentData.json)

# Ensure script directory
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Default JSON template from LLEA.ps1
$DefaultJson = @'
{
  "Announcements": {
    "Default": {
      "Text": "No announcements at this time.",
      "Details": "",
      "Links": []
    },
    "Targeted": []
  },
  "Support": {
    "Text": "Contact IT Support.",
    "Links": []
  }
}
'@

# Configuration file for persisting repository URL, API base URL, and PAT
$ConfigPath = Join-Path $ScriptDir "ContentDataEditor.config.json"
$PatPath = Join-Path $ScriptDir "ContentDataEditor.cred.xml"
$LogPath = Join-Path $ScriptDir "ContentDataEditor.log"

# Help content as a separate string variable
$helpContent = @'
**Welcome to the ContentData JSON Editor**

This tool allows you to edit the JSON content for the Lincoln Laboratory Endpoint Advisor app, which displays announcements and support information on user systems. The content is stored in a JSON file, typically hosted on a Git repository (e.g., GitHub Enterprise Server), and this editor simplifies updating that content.

**1. Setting Up the Editor**

- **GitHub Personal Access Token (PAT)**:
  - To save changes directly to the Git repository, you need a Personal Access Token with `repo` scope or `Contents: Read & Write` permissions.
  - Enter the PAT in the **GitHub PAT** field at the top of the editor (it displays as asterisks for security).
  - Press **Enter** or click the **Save PAT** button to save it securely to `ContentDataEditor.cred.xml`.
  - A "GitHub PAT saved securely!" message confirms success. If you see an error, ensure the PAT is valid and not empty.

- **Repository URL**:
  - The default URL is `https://raw.servername/EndpointEngineering/EndpointAdvisor/main/ContentData.json`.
  - To use a different repository, update the **Repository URL** field with the URL to your `ContentData.json` file and press **Enter**.
  - The editor will fetch and load the JSON content, displaying a "Updated repository URL and reloaded JSON!" message on success.
  - If the URL is invalid, you'll see an error message prompting for a valid URL starting with `http://` or `https://`.

- **API Base URL**:
  - Specify the API base URL for your Git server (e.g., `https://servername/api/v3` for GitHub Enterprise Server).
  - Update the **API Base URL** field and press **Enter** to save it.
  - This is required for saving changes to the repository.

**2. Editing Content**

The editor has three main tabs for content management, plus a preview panel on the right to see how the content will appear in the Endpoint Advisor app.

- **Default Announcement Tab**:
  - Edit the **Main Text** and **Details** for the announcement shown to all users by default.
  - Add hyperlinks in the **Links** section by clicking **Add Link** (enter a display name and URL, e.g., "Support Page" and `https://example.com`). Click **Remove Last Link** to delete the last link.
  - Changes are reflected in the preview panel instantly.

- **Targeted Announcements Tab**:
  - Create announcements for specific systems based on registry conditions (e.g., `HKLM:\SOFTWARE\MITLL\Targeting`, `Group`, `VPN-Pilot`).
  - Click **Add Targeted Announcement** to create a new announcement.
  - For each targeted announcement:
    - **Title**: The main message text.
    - **Message**: Additional details.
    - **Announcement Enabled**: Check to make the announcement active in the Endpoint Advisor app if its registry condition is met. Unchecked announcements (e.g., "VPN-Pilot" with `"Enabled": false`) are excluded from the app and preview.
    - **Append to Default**: Check to append the announcement to the default; uncheck to replace the default if the condition is met.
    - **Registry Condition**: Specify the registry path, key name, and value to target specific systems (e.g., `HKLM:\SOFTWARE\MITLL\Targeting`, `Group`, `VPN-Pilot`).
    - **Links**: Add hyperlinks similar to the Default Announcement.
    - Click **Remove This Announcement** to delete the targeted announcement.
  - Only enabled announcements with "Append to Default" checked appear in the preview if conditions are met.

- **Support Tab**:
  - Edit the **Support Text** for the support message shown in the Endpoint Advisor app.
  - Add hyperlinks in the **Links** section, similar to the Default Announcement.
  - Changes are reflected in the preview panel instantly.

- **Markdown Formatting**:
  - All text fields (Main Text, Details, Support Text, Title, Message) support Markdown for styling:
    - **Bold**: Use `**text**` (e.g., `**Important**` renders as **Important**).
    - *Italic*: Use `*text*` (e.g., `*Note*` renders as *Note*).
    - __Underline__: Use `__text__` (e.g., `__Details__` renders as __Details__).
    - [color]Colored Text[/color]: Use `[color]text[/color]` with `green`, `red`, `yellow`, or `blue` (e.g., `[blue]Click here[/blue]` renders as <span style="color:blue">Click here</span>).
  - The preview panel shows the formatted text as it will appear in the Endpoint Advisor app.

**3. Managing JSON Content**

The buttons at the bottom of the editor help you manage the JSON content:

- **Reload from Repo**: Fetches the latest JSON from the repository URL specified in the **Repository URL** field. A "Reloaded from [URL]!" message confirms success.
- **Validate JSON**: Checks if the current content is valid for the Endpoint Advisor app. A "JSON is valid!" message confirms success, or an error message indicates issues (e.g., missing required fields).
- **Save to File**: Saves the JSON to a local file (e.g., `ContentData_edited.json`) for manual upload to the repository. A file dialog lets you choose the location.
- **Save to GitHub**: Pushes changes directly to the Git repository using the saved PAT and API base URL. A "Successfully saved to Git repository!" message confirms success. If it fails, an error message and log entry provide details.
- **Copy to Clipboard**: Copies the JSON to the clipboard for external use.
- **Close**: Exits the editor, saving any configuration changes (e.g., repository URL, API base URL).

**4. Preview Panel**

- The right panel shows a real-time preview of how the content will appear in the Endpoint Advisor app.
- The **Announcements** section displays:
  - The Default Announcement's Main Text and Details.
  - Targeted announcements that are enabled (`"Announcement Enabled": true`) and set to append (`"Append to Default": true`), if their registry conditions are met.
  - Links from the Default Announcement and enabled Targeted Announcements.
- The **Support** section displays the Support Text and Links.
- Disabled announcements (e.g., "VPN-Pilot" with `"Enabled": false`) are excluded from the preview.

**5. Troubleshooting**

- **Log Files**:
  - Check `ContentDataEditor.log` in the script directory for detailed logs of actions (e.g., PAT saving, JSON fetching, Git saves).
  - For issues with the Endpoint Advisor app, check `C:\Scripts\JEdit\LLEndpointAdvisor.log`.
- **Common Issues**:
  - **PAT Errors**: Ensure the PAT has `repo` scope or `Contents: Read & Write` permissions. Regenerate the PAT in your Git server if needed (Settings > Developer settings > Personal access tokens).
  - **Invalid Repository URL**: Verify the URL points to a valid `ContentData.json` file (e.g., `https://raw.servername/owner/repo/branch/ContentData.json`).
  - **Invalid API Base URL**: Ensure the API base URL matches your Git server (e.g., `https://servername/api/v3` for GitHub Enterprise Server).
  - **JSON Save Failures**: Check the log for Git API errors (e.g., `401 Unauthorized` for invalid PAT, `403 Forbidden` for access issues, `422 Unprocessable Entity` for SHA mismatch).
  - **UI Issues**: If tabs or text appear oversized, check your system's DPI scaling (Display Settings > Scale and layout) and set to 100% if needed.
- **Contact**: For further assistance, contact your IT administrator or the system administrator responsible for the Endpoint Advisor deployment.

**6. Additional Notes**

- The editor saves configuration (e.g., repository URL, API base URL) to `ContentDataEditor.config.json`.
- The PAT is stored securely in `ContentDataEditor.cred.xml` and only needs to be re-entered if it expires or is invalid.
- Changes saved to the Git repository are reflected in the Endpoint Advisor app after its next update cycle (configured in `LLEA.ps1`, typically every 15 minutes).
- Ensure you have write permissions to the script directory for logs and configuration files.

For further assistance, consult the **Help** tab or contact your IT administrator.
'@

# Function to log messages
function Write-Log {
    param($Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Out-File -FilePath $LogPath -Append -Encoding utf8
}

# Load or initialize configuration
function Load-Configuration {
    $defaultConfig = @{
        RepositoryUrl = "https://raw.servername/EndpointEngineering/EndpointAdvisor/main/ContentData.json"
        ApiBaseUrl = "https://servername/api/v3"
    }
    if (Test-Path $ConfigPath) {
        try {
            $config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
            Write-Log "Loaded configuration from $ConfigPath"
            return $config
        } catch {
            Write-Log "Failed to load config: $($_.Exception.Message). Using default."
        }
    }
    $defaultConfig | ConvertTo-Json | Out-File $ConfigPath -Force
    Write-Log "Created default configuration at $ConfigPath"
    return $defaultConfig
}

function Save-Configuration {
    param($Config)
    try {
        $Config | ConvertTo-Json | Out-File $ConfigPath -Force
        Write-Log "Saved configuration to $ConfigPath"
    } catch {
        Write-Log "Failed to save config: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Failed to save configuration: $($_.Exception.Message)", "Error", "OK", "Error")
    }
}

# Load or save PAT securely
function Load-PAT {
    if (Test-Path $PatPath) {
        try {
            $cred = Import-Clixml -Path $PatPath
            $pat = $cred.GetNetworkCredential().Password
            Write-Log "Loaded PAT from $PatPath"
            return $pat
        } catch {
            Write-Log "Failed to load PAT: $($_.Exception.Message)"
            return $null
        }
    }
    Write-Log "No PAT found at $PatPath"
    return $null
}

function Save-PAT {
    param($PAT)
    try {
        if (-not $PAT) {
            throw "PAT is empty or null"
        }
        $secureString = ConvertTo-SecureString -String $PAT -AsPlainText -Force
        $cred = New-Object System.Management.Automation.PSCredential("GitHubPAT", $secureString)
        $cred | Export-Clixml -Path $PatPath -Force
        Write-Log "Saved PAT securely to $PatPath"
        return $true
    } catch {
        Write-Log "Failed to save PAT: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Failed to save PAT: $($_.Exception.Message)", "Error", "OK", "Error")
        return $false
    }
}

# Load initial config
$global:Config = Load-Configuration
$global:JsonUrl = $global:Config.RepositoryUrl
$global:ApiBaseUrl = $global:Config.ApiBaseUrl
$global:GitHubPAT = Load-PAT

# Load required assemblies
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to fetch JSON
function Fetch-Json {
    try {
        $headers = if ($global:GitHubPAT) { @{ Authorization = "Bearer $global:GitHubPAT" } } else { @{} }
        $response = Invoke-WebRequest -Uri $global:JsonUrl -UseBasicParsing -TimeoutSec 30 -Headers $headers
        $json = $response.Content | ConvertFrom-Json
        Write-Log "Fetched JSON from $global:JsonUrl"
        return $json
    } catch {
        Write-Log "Failed to fetch JSON from $global:JsonUrl: $($_.Exception.Message)"
        return $DefaultJson | ConvertFrom-Json
    }
}

# Function to validate JSON
function Validate-Json {
    param($JsonObject)
    try {
        if (-not $JsonObject.PSObject.Properties.Match('Announcements') -or -not $JsonObject.PSObject.Properties.Match('Support')) {
            throw "JSON missing 'Announcements' or 'Support' properties."
        }
        if (-not $JsonObject.Announcements.PSObject.Properties.Match('Default') -or -not $JsonObject.Announcements.PSObject.Properties.Match('Targeted')) {
            throw "Announcements missing 'Default' or 'Targeted' properties."
        }
        Write-Log "JSON validation successful"
        return $true
    } catch {
        Write-Log "JSON validation failed: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Invalid JSON: $($_.Exception.Message)", "Validation Error", "OK", "Error")
        return $false
    }
}

# Function to save JSON to GitHub Enterprise Server
function Save-ToGitHub {
    param($JsonObject)
    try {
        if (-not $global:GitHubPAT) {
            $newPat = [System.Windows.MessageBox]::Show("No PAT is set. Please enter it in the UI and press Enter or click Save PAT, then try again.", "PAT Missing", "OK", "Warning")
            Write-Log "PAT not set during save attempt"
            return $false
        }
        $repoUrlParts = $global:JsonUrl -split "/"
        if ($repoUrlParts.Length -lt 7) {
            throw "Invalid repository URL. Expected format: https://raw.servername/owner/repo/branch/file"
        }
        $owner = $repoUrlParts[3] # EndpointEngineering
        $repo = $repoUrlParts[4]  # EndpointAdvisor
        $branch = $repoUrlParts[5] # main
        $filePath = $repoUrlParts[6..($repoUrlParts.Length-1)] -join "/" # ContentData.json
        $apiUrl = "$global:ApiBaseUrl/repos/$owner/$repo/contents/$filePath"
        Write-Log "Parsed URL components: owner=$owner, repo=$repo, branch=$branch, filePath=$filePath"
        Write-Log "Preparing to save to Git at $apiUrl"

        # Get the current file's SHA
        $headers = @{
            Authorization = "Bearer $global:GitHubPAT"
            Accept = "application/vnd.github+json"
            "User-Agent" = "ContentDataEditor"
        }
        try {
            $response = Invoke-RestMethod -Uri $apiUrl -Headers $headers -Method Get
            $sha = $response.sha
            Write-Log "Retrieved SHA for ${filePath}: ${sha}"
        } catch {
            if ($_.Exception.Response.StatusCode.Value__ -eq 404) {
                Write-Log "File not found at $apiUrl. Creating new file."
                $sha = $null
            } else {
                throw "Failed to retrieve file SHA: $($_.Exception.Message)"
            }
        }

        # Prepare the new content
        $content = $JsonObject | ConvertTo-Json -Depth 10 | ForEach-Object { [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($_)) }
        $body = @{
            message = "Update ContentData.json via ContentDataEditor"
            content = $content
            branch = $branch
        }
        if ($sha) {
            $body.sha = $sha
        }
        $bodyJson = $body | ConvertTo-Json -Compress

        # Update or create the file
        $response = Invoke-RestMethod -Uri $apiUrl -Headers $headers -Method Put -Body $bodyJson
        Write-Log "Successfully saved to Git repository: $apiUrl"
        [System.Windows.MessageBox]::Show("Successfully saved to Git repository!", "Success")
        return $true
    } catch {
        $errorMessage = $_.Exception.Message
        if ($_.Exception.Response) {
            try {
                $errorResponse = $_.Exception.Response.GetResponseStream()
                $reader = New-Object System.IO.StreamReader($errorResponse)
                $errorBody = $reader.ReadToEnd()
                $errorMessage += "`nGit API Response: $errorBody"
            } catch {
                $errorMessage += "`nFailed to read Git API response: $($_.Exception.Message)"
            }
            Write-Log "Git save failed: $errorMessage"
            [System.Windows.MessageBox]::Show("Failed to save to Git: $errorMessage", "Error", "OK", "Error")
        } else {
            Write-Log "Git save failed: $errorMessage"
            [System.Windows.MessageBox]::Show("Failed to save to Git: $errorMessage", "Error", "OK", "Error")
        }
        return $false
    }
}

# Function to convert Markdown to TextBlock (based on LLEA.ps1's Convert-MarkdownToTextBlock)
function Convert-MarkdownToTextBlock {
    param(
        [string]$Text,
        [System.Windows.Controls.TextBlock]$TargetTextBlock
    )
    try {
        $TargetTextBlock.Inlines.Clear()
        if (-not $Text) {
            Write-Log "No text provided for Markdown conversion"
            return
        }

        $regexColor = "\[(green|red|yellow|blue)\](.*?)\[/\1\]"
        $regexBold = "\*\*(.*?)\*\*"
        $regexItalic = "\*(.*?)\*"
        $regexUnderline = "__(.*?)__"

        $currentText = $Text
        $colorPlaceholders = @{}
        $placeholderCounter = 0
        $colorMatches = [regex]::Matches($Text, $regexColor) | Sort-Object Index -Descending

        foreach ($match in $colorMatches) {
            $placeholder = "{COLORPH$placeholderCounter}"
            $leftOk = ($match.Index -ge 2) -and ($Text.Substring($match.Index - 2, 2) -eq "**")
            $rightOk = (($match.Index + $match.Length + 2) -le $Text.Length) -and ($Text.Substring($match.Index + $match.Length, 2) -eq "**")
            $isBold = $leftOk -and $rightOk
            $colorPlaceholders[$placeholder] = @{
                Text = $match.Groups[2].Value
                Color = $match.Groups[1].Value
                IsBold = $isBold
            }
            $currentText = $currentText.Remove($match.Index, $match.Length).Insert($match.Index, $placeholder)
            $placeholderCounter++
        }

        $matches = @()
        $boldMatches = [regex]::Matches($currentText, $regexBold)
        $italicMatches = [regex]::Matches($currentText, $regexItalic)
        $underlineMatches = [regex]::Matches($currentText, $regexUnderline)

        foreach ($match in $boldMatches) {
            $matches += [PSCustomObject]@{ Index = $match.Index; Length = $match.Length; Text = $match.Groups[1].Value; Type = "Bold"; FullMatch = $match.Value }
        }
        foreach ($match in $italicMatches) {
            if ([string]::IsNullOrWhiteSpace($match.Groups[1].Value)) { continue }
            $matches += [PSCustomObject]@{ Index = $match.Index; Length = $match.Length; Text = $match.Groups[1].Value; Type = "Italic"; FullMatch = $match.Value }
        }
        foreach ($match in $underlineMatches) {
            $matches += [PSCustomObject]@{ Index = $match.Index; Length = $match.Length; Text = $match.Groups[1].Value; Type = "Underline"; FullMatch = $match.Value }
        }

        $matches = $matches | Sort-Object Index
        $lastIndex = 0
        $runs = @()

        foreach ($match in $matches) {
            if ($match.Index -gt $lastIndex) {
                $plainText = $currentText.Substring($lastIndex, $match.Index - $lastIndex)
                $runs += Process-TextSegment -Text $plainText -ColorPlaceholders $colorPlaceholders
            }

            $text = $match.Text
            if ($colorPlaceholders.ContainsKey($text)) {
                $colorInfo = $colorPlaceholders[$text]
                $innerRuns = Process-InnerMarkdown -Text $colorInfo.Text -Color $colorInfo.Color -IsBold $colorInfo.IsBold
                $runs += $innerRuns
            } else {
                $run = New-Object System.Windows.Documents.Run($text)
                if ($match.Type -eq "Bold") {
                    $run.FontWeight = [System.Windows.FontWeights]::Bold
                } elseif ($match.Type -eq "Italic") {
                    $run.FontStyle = [System.Windows.FontStyles]::Italic
                } elseif ($match.Type -eq "Underline") {
                    $run.TextDecorations = [System.Windows.TextDecorations]::Underline
                }
                $runs += $run
            }
            $lastIndex = $match.Index + $match.Length
        }

        if ($lastIndex -lt $currentText.Length) {
            $plainText = $currentText.Substring($lastIndex)
            $runs += Process-TextSegment -Text $plainText -ColorPlaceholders $colorPlaceholders
        }

        foreach ($run in $runs) {
            $TargetTextBlock.Inlines.Add($run)
        }
        Write-Log "Successfully converted Markdown for TextBlock"
    } catch {
        Write-Log "Failed to convert Markdown: $($_.Exception.Message)"
        $TargetTextBlock.Inlines.Clear()
        $TargetTextBlock.Inlines.Add((New-Object System.Windows.Documents.Run($Text)))
        Write-Log "Fallback: Set raw text for TextBlock"
    }
}

function Process-TextSegment {
    param(
        [string]$Text,
        [hashtable]$ColorPlaceholders
    )
    $runs = @()
    $currentPos = 0
    $placeholderRegex = [regex] "{COLORPH\d+}"
    $placeholderMatches = $placeholderRegex.Matches($Text) | Sort-Object Index

    foreach ($match in $placeholderMatches) {
        if ($match.Index -gt $currentPos) {
            $plainText = $Text.Substring($currentPos, $match.Index - $currentPos)
            $runs += New-Object System.Windows.Documents.Run($plainText)
        }
        $placeholder = $match.Value
        if ($ColorPlaceholders.ContainsKey($placeholder)) {
            $colorInfo = $ColorPlaceholders[$placeholder]
            $innerRuns = Process-InnerMarkdown -Text $colorInfo.Text -Color $colorInfo.Color -IsBold $colorInfo.IsBold
            $runs += $innerRuns
        }
        $currentPos = $match.Index + $match.Length
    }

    if ($currentPos -lt $Text.Length) {
        $plainText = $Text.Substring($currentPos)
        $runs += New-Object System.Windows.Documents.Run($plainText)
    }
    return $runs
}

function Process-InnerMarkdown {
    param(
        [string]$Text,
        [string]$Color,
        [bool]$IsBold
    )
    $runs = @()
    if (-not $Text) { return $runs }

    $regexBold = "\*\*(.*?)\*\*"
    $regexItalic = "\*(.*?)\*"
    $regexUnderline = "__(.*?)__"

    $matches = @()
    $boldMatches = [regex]::Matches($Text, $regexBold)
    $italicMatches = [regex]::Matches($Text, $regexItalic)
    $underlineMatches = [regex]::Matches($Text, $regexUnderline)

    foreach ($match in $boldMatches) {
        $matches += [PSCustomObject]@{ Index = $match.Index; Length = $match.Length; Text = $match.Groups[1].Value; Type = "Bold"; FullMatch = $match.Value }
    }
    foreach ($match in $italicMatches) {
        if ([string]::IsNullOrWhiteSpace($match.Groups[1].Value)) { continue }
        $matches += [PSCustomObject]@{ Index = $match.Index; Length = $match.Length; Text = $match.Groups[1].Value; Type = "Italic"; FullMatch = $match.Value }
    }
    foreach ($match in $underlineMatches) {
        $matches += [PSCustomObject]@{ Index = $match.Index; Length = $match.Length; Text = $match.Groups[1].Value; Type = "Underline"; FullMatch = $match.Value }
    }

    $matches = $matches | Sort-Object Index
    $lastIndex = 0

    foreach ($match in $matches) {
        if ($match.Index -gt $lastIndex) {
            $plainText = $Text.Substring($lastIndex, $match.Index - $lastIndex)
            $run = New-Object System.Windows.Documents.Run($plainText)
            if ($Color) {
                $colorBrush = [System.Windows.Media.Brushes]::($Color.Substring(0,1).ToUpper() + $Color.Substring(1))
                $run.Foreground = $colorBrush
            }
            if ($IsBold) {
                $run.FontWeight = [System.Windows.FontWeights]::Bold
            }
            $runs += $run
        }
        $run = New-Object System.Windows.Documents.Run($match.Text)
        if ($match.Type -eq "Bold") {
            $run.FontWeight = [System.Windows.FontWeights]::Bold
        } elseif ($match.Type -eq "Italic") {
            $run.FontStyle = [System.Windows.FontStyles]::Italic
        } elseif ($match.Type -eq "Underline") {
            $run.TextDecorations = [System.Windows.TextDecorations]::Underline
        }
        if ($Color) {
            $colorBrush = [System.Windows.Media.Brushes]::($Color.Substring(0,1).ToUpper() + $Color.Substring(1))
            $run.Foreground = $colorBrush
        }
        if ($IsBold) {
            $run.FontWeight = [System.Windows.FontWeights]::Bold
        }
        $runs += $run
        $lastIndex = $match.Index + $match.Length
    }

    if ($lastIndex -lt $Text.Length) {
        $plainText = $Text.Substring($lastIndex)
        $run = New-Object System.Windows.Documents.Run($plainText)
        if ($Color) {
            $colorBrush = [System.Windows.Media.Brushes]::($Color.Substring(0,1).ToUpper() + $Color.Substring(1))
            $run.Foreground = $colorBrush
        }
        if ($IsBold) {
            $run.FontWeight = [System.Windows.FontWeights]::Bold
        }
        $runs += $run
    }
    return $runs
}

# XAML for the GUI
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="ContentData JSON Editor" Height="700" Width="1000" ResizeMode="CanResize">
    <Window.Resources>
        <Style TargetType="TabItem">
            <Setter Property="Height" Value="25"/>
            <Setter Property="Padding" Value="10,2"/>
            <Setter Property="Margin" Value="0"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border x:Name="Border" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1,1,1,0" Margin="0">
                            <ContentPresenter x:Name="ContentSite" ContentSource="Header" VerticalAlignment="Center" HorizontalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="#0078D7"/>
                                <Setter TargetName="Border" Property="BorderBrush" Value="#0055A4"/>
                                <Setter Property="Foreground" Value="White"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="False">
                                <Setter TargetName="Border" Property="Background" Value="#F0F0F0"/>
                                <Setter TargetName="Border" Property="BorderBrush" Value="#D3D3D3"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
            <TextBlock Text="GitHub PAT:" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,5,0" ToolTip="Enter your GitHub Personal Access Token with repo scope. Press Enter or click Save PAT to save securely."/>
            <PasswordBox x:Name="GitHubPatTextBox" Width="300" ToolTip="Enter your GitHub Personal Access Token with repo scope. Press Enter or click Save PAT to save securely." Margin="0,0,5,0"/>
            <Button x:Name="SavePatButton" Content="Save PAT" Width="80" ToolTip="Save the entered GitHub Personal Access Token securely."/>
        </StackPanel>
        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,10">
            <TextBlock Text="Repository URL:" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,5,0" ToolTip="Enter the URL to the ContentData.json file in your Git repository."/>
            <TextBox x:Name="RepoUrlTextBox" Width="500" Text="$($global:Config.RepositoryUrl)" ToolTip="Enter the URL to the ContentData.json file. Press Enter to update."/>
        </StackPanel>
        <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,0,0,10">
            <TextBlock Text="API Base URL:" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,5,0" ToolTip="Enter the API base URL for your Git server (e.g., https://servername/api/v3)."/>
            <TextBox x:Name="ApiBaseUrlTextBox" Width="500" Text="$($global:Config.ApiBaseUrl)" ToolTip="Enter the API base URL. Press Enter to update."/>
        </StackPanel>
        <TextBlock Grid.Row="3" Text="Edit the Announcements and Support content for Lincoln Laboratory Endpoint Advisor. Use the tabs to edit sections, and preview the formatted output on the right. Markdown supported: **bold**, *italic*, __underline__, [color]text[/color] (green, red, yellow, blue). Note: Announcements must have 'Announcement Enabled' checked to appear in the Endpoint Advisor app." FontSize="12" TextWrapping="Wrap" Margin="0,0,0,10"/>
        <Grid Grid.Row="4">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <TabControl x:Name="EditorTabs" Grid.Column="0">
                <TabItem Header="Default Announcement">
                    <Grid Margin="5">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <TextBlock Grid.Row="0" Text="Main Text" FontWeight="Bold" Margin="0,0,0,5" ToolTip="Main announcement text. Use Markdown for formatting."/>
                        <TextBox x:Name="DefaultText" Grid.Row="1" AcceptsReturn="True" TextWrapping="Wrap" Height="100" Margin="0,0,0,5" ToolTip="Main announcement text. Use Markdown: **bold**, *italic*, __underline__, [color]text[/color]"/>
                        <TextBlock Grid.Row="2" Text="Details" FontWeight="Bold" Margin="0,0,0,5" ToolTip="Additional details for the announcement."/>
                        <TextBox x:Name="DefaultDetails" Grid.Row="3" AcceptsReturn="True" TextWrapping="Wrap" Height="100" Margin="0,0,0,5" ToolTip="Details text. Use Markdown for formatting."/>
                        <StackPanel Grid.Row="4" Orientation="Vertical">
                            <TextBlock Text="Links" FontWeight="Bold" Margin="0,0,0,5" ToolTip="Add hyperlinks to external resources."/>
                            <StackPanel x:Name="DefaultLinksPanel" Orientation="Vertical" Margin="0,0,0,5"/>
                            <StackPanel Orientation="Horizontal">
                                <Button x:Name="AddDefaultLinkButton" Content="Add Link" Width="80" Margin="0,0,5,0" ToolTip="Add a new hyperlink to the Default Announcement."/>
                                <Button x:Name="RemoveDefaultLinkButton" Content="Remove Last Link" Width="120" ToolTip="Remove the last hyperlink from the Default Announcement."/>
                            </StackPanel>
                        </StackPanel>
                    </Grid>
                </TabItem>
                <TabItem Header="Targeted Announcements">
                    <Grid Margin="5">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <ScrollViewer Grid.Row="0" VerticalScrollBarVisibility="Auto">
                            <StackPanel x:Name="TargetedAnnouncementsPanel"/>
                        </ScrollViewer>
                        <Button x:Name="AddTargetedButton" Grid.Row="1" Content="Add Targeted Announcement" Width="150" Margin="0,5,0,0" ToolTip="Add a new targeted announcement based on registry conditions."/>
                    </Grid>
                </TabItem>
                <TabItem Header="Support">
                    <Grid Margin="5">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <TextBlock Grid.Row="0" Text="Support Text" FontWeight="Bold" Margin="0,0,0,5" ToolTip="Main support message. Use Markdown for formatting."/>
                        <TextBox x:Name="SupportText" Grid.Row="1" AcceptsReturn="True" TextWrapping="Wrap" Height="100" Margin="0,0,0,5" ToolTip="Support message. Use Markdown: **bold**, *italic*, __underline__, [color]text[/color]"/>
                        <TextBlock Grid.Row="2" Text="Links" FontWeight="Bold" Margin="0,0,0,5" ToolTip="Add hyperlinks to support resources."/>
                        <StackPanel x:Name="SupportLinksPanel" Grid.Row="3" Orientation="Vertical" Margin="0,0,0,5"/>
                        <StackPanel Grid.Row="4" Orientation="Horizontal">
                            <Button x:Name="AddSupportLinkButton" Content="Add Link" Width="80" Margin="0,0,5,0" ToolTip="Add a new hyperlink to the Support section."/>
                            <Button x:Name="RemoveSupportLinkButton" Content="Remove Last Link" Width="120" ToolTip="Remove the last hyperlink from the Support section."/>
                        </StackPanel>
                    </Grid>
                </TabItem>
                <TabItem Header="Help">
                    <Grid Margin="5">
                        <RichTextBox x:Name="HelpRichTextBox" IsReadOnly="True" VerticalScrollBarVisibility="Auto" Background="Transparent" BorderThickness="0" FontSize="12" Margin="5"/>
                    </Grid>
                </TabItem>
            </TabControl>
            <GridSplitter Grid.Column="1" Width="10" HorizontalAlignment="Stretch"/>
            <ScrollViewer Grid.Column="2" VerticalScrollBarVisibility="Auto">
                <StackPanel x:Name="PreviewPanel">
                    <TextBlock Text="Preview" FontWeight="Bold" FontSize="14" Margin="0,0,0,10"/>
                    <Border BorderBrush="#00008B" BorderThickness="1" Padding="5" CornerRadius="3" Background="White" Margin="0,0,0,5">
                        <StackPanel>
                            <TextBlock Text="Announcements" FontWeight="Bold" FontSize="12" Margin="0,0,0,5"/>
                            <TextBlock x:Name="PreviewAnnouncementsText" FontSize="11" TextWrapping="Wrap"/>
                            <TextBlock x:Name="PreviewAnnouncementsDetails" FontSize="11" TextWrapping="Wrap" Margin="0,5,0,0"/>
                            <StackPanel x:Name="PreviewAppendedAnnouncements" Orientation="Vertical" Margin="0,5,0,0"/>
                            <StackPanel x:Name="PreviewAnnouncementsLinks" Orientation="Vertical" Margin="0,5,0,0"/>
                        </StackPanel>
                    </Border>
                    <Border BorderBrush="#00008B" BorderThickness="1" Padding="5" CornerRadius="3" Background="White" Margin="0,5,0,0">
                        <StackPanel>
                            <TextBlock Text="Support" FontWeight="Bold" FontSize="12" Margin="0,0,0,5"/>
                            <TextBlock x:Name="PreviewSupportText" FontSize="11" TextWrapping="Wrap"/>
                            <StackPanel x:Name="PreviewSupportLinks" Orientation="Vertical" Margin="0,5,0,0"/>
                        </StackPanel>
                    </Border>
                </StackPanel>
            </ScrollViewer>
        </Grid>
        <StackPanel Grid.Row="5" Orientation="Horizontal" Margin="0,10,0,0">
            <Button x:Name="ReloadButton" Content="Reload from Repo" Width="120" Margin="0,0,10,0" ToolTip="Reload the JSON from the specified repository URL."/>
            <Button x:Name="ValidateButton" Content="Validate JSON" Width="120" Margin="0,0,10,0" ToolTip="Check if the JSON is valid for use in Endpoint Advisor."/>
            <Button x:Name="SaveButton" Content="Save to File" Width="120" Margin="0,0,10,0" ToolTip="Save the JSON to a local file for manual upload to GitHub."/>
            <Button x:Name="SaveToGitHubButton" Content="Save to GitHub" Width="120" Margin="0,0,10,0" ToolTip="Save the JSON directly to the Git repository."/>
            <Button x:Name="CopyButton" Content="Copy to Clipboard" Width="120" Margin="0,0,10,0" ToolTip="Copy the JSON to the clipboard."/>
            <Button x:Name="CloseButton" Content="Close" Width="120" ToolTip="Close the editor."/>
        </StackPanel>
    </Grid>
</Window>
"@

# Load XAML
try {
    $xmlDoc = New-Object System.Xml.XmlDocument
    $xmlDoc.LoadXml($xaml)
    $reader = New-Object System.Xml.XmlNodeReader $xmlDoc
    $window = [Windows.Markup.XamlReader]::Load($reader)
    Write-Log "Successfully loaded XAML UI"
} catch {
    Write-Log "Failed to load XAML: $($_.Exception.Message)"
    [System.Windows.MessageBox]::Show("Failed to load UI: $($_.Exception.Message)", "Error", "OK", "Error")
    exit
}

# Find controls
$GitHubPatTextBox = $window.FindName("GitHubPatTextBox")
$RepoUrlTextBox = $window.FindName("RepoUrlTextBox")
$ApiBaseUrlTextBox = $window.FindName("ApiBaseUrlTextBox")
$SavePatButton = $window.FindName("SavePatButton")
$EditorTabs = $window.FindName("EditorTabs")
$HelpRichTextBox = $window.FindName("HelpRichTextBox")
$DefaultText = $window.FindName("DefaultText")
$DefaultDetails = $window.FindName("DefaultDetails")
$DefaultLinksPanel = $window.FindName("DefaultLinksPanel")
$AddDefaultLinkButton = $window.FindName("AddDefaultLinkButton")
$RemoveDefaultLinkButton = $window.FindName("RemoveDefaultLinkButton")
$TargetedAnnouncementsPanel = $window.FindName("TargetedAnnouncementsPanel")
$AddTargetedButton = $window.FindName("AddTargetedButton")
$SupportText = $window.FindName("SupportText")
$SupportLinksPanel = $window.FindName("SupportLinksPanel")
$AddSupportLinkButton = $window.FindName("AddSupportLinkButton")
$RemoveSupportLinkButton = $window.FindName("RemoveSupportLinkButton")
$PreviewPanel = $window.FindName("PreviewPanel")
$PreviewAnnouncementsText = $window.FindName("PreviewAnnouncementsText")
$PreviewAnnouncementsDetails = $window.FindName("PreviewAnnouncementsDetails")
$PreviewAppendedAnnouncements = $window.FindName("PreviewAppendedAnnouncements")
$PreviewAnnouncementsLinks = $window.FindName("PreviewAnnouncementsLinks")
$PreviewSupportText = $window.FindName("PreviewSupportText")
$PreviewSupportLinks = $window.FindName("PreviewSupportLinks")
$ReloadButton = $window.FindName("ReloadButton")
$ValidateButton = $window.FindName("ValidateButton")
$SaveButton = $window.FindName("SaveButton")
$SaveToGitHubButton = $window.FindName("SaveToGitHubButton")
$CopyButton = $window.FindName("CopyButton")
$CloseButton = $window.FindName("CloseButton")

# Add the live preview event handlers for the main text boxes
$DefaultText.Add_TextChanged({ if (-not $script:IsLoading) { Update-Preview } })
$DefaultDetails.Add_TextChanged({ if (-not $script:IsLoading) { Update-Preview } })
$SupportText.Add_TextChanged({ if (-not $script:IsLoading) { Update-Preview } })

# Global JSON object
$global:JsonData = Fetch-Json
$global:TargetedEditors = @()

# Function to handle PAT saving (shared for Enter key and Save PAT button)
function Save-PatHandler {
    $newPat = $GitHubPatTextBox.Password.Trim()
    Write-Log "PAT submission triggered. PAT length: $($newPat.Length)"
    if ($newPat) {
        Write-Log "Attempting to save PAT from PasswordBox"
        $global:GitHubPAT = $newPat
        if (Save-PAT -PAT $newPat) {
            [System.Windows.MessageBox]::Show("GitHub PAT saved securely! Please try saving to Git again.", "Success")
        }
    } else {
        Write-Log "Invalid PAT entered: empty or whitespace"
        [System.Windows.MessageBox]::Show("Please enter a valid Personal Access Token.", "Error", "OK", "Error")
    }
}

# Function to create a link editor
function New-LinkEditor {
    param($Panel, $Name = "", $Url = "")
    $stackPanel = New-Object System.Windows.Controls.StackPanel
    $stackPanel.Orientation = "Horizontal"
    $stackPanel.Margin = [System.Windows.Thickness]::new(0,0,0,5)
    $nameBox = New-Object System.Windows.Controls.TextBox
    $nameBox.Width = 150
    $nameBox.Text = $Name
    $nameBox.Margin = [System.Windows.Thickness]::new(0,0,5,0)
    $nameBox.ToolTip = "Name of the hyperlink (display text)"
    $nameBox.Add_TextChanged({ if (-not $script:IsLoading) { Update-Preview } })
    $urlBox = New-Object System.Windows.Controls.TextBox
    $urlBox.Width = 300
    $urlBox.Text = $Url
    $urlBox.ToolTip = "URL for the hyperlink (e.g., https://example.com)"
    $urlBox.Add_TextChanged({ if (-not $script:IsLoading) { Update-Preview } })
    $stackPanel.Children.Add($nameBox)
    $stackPanel.Children.Add($urlBox)
    $Panel.Children.Add($stackPanel)
    return @{ NameBox = $nameBox; UrlBox = $urlBox }
}

# Function to create a targeted announcement editor
function New-TargetedAnnouncementEditor {
    param($Message = @{ Text = ""; Details = ""; Links = @(); AppendToDefault = $true; DisableAppended = $false; Enabled = $true; Condition = @{ Type = "Registry"; Path = ""; Name = ""; Value = "" } })
    $border = New-Object System.Windows.Controls.Border
    $border.BorderBrush = [System.Windows.Media.Brushes]::Gray
    $border.BorderThickness = 1
    $border.Margin = [System.Windows.Thickness]::new(0,0,0,5)
    $border.Padding = 5
    $stackPanel = New-Object System.Windows.Controls.StackPanel
    $stackPanel.Orientation = "Vertical"

    $titleLabel = New-Object System.Windows.Controls.TextBlock
    $titleLabel.Text = "Title"
    $titleLabel.FontWeight = "Bold"
    $titleBox = New-Object System.Windows.Controls.TextBox
    $titleBox.AcceptsReturn = $true
    $titleBox.TextWrapping = "Wrap"
    $titleBox.Height = 80
    $titleBox.Text = $Message.Text
    $titleBox.ToolTip = "Title of the announcement. Use Markdown for formatting."
    $titleBox.Add_TextChanged({ if (-not $script:IsLoading) { Update-Preview } })

    $messageLabel = New-Object System.Windows.Controls.TextBlock
    $messageLabel.Text = "Message"
    $messageLabel.FontWeight = "Bold"
    $messageLabel.Margin = [System.Windows.Thickness]::new(0,5,0,0)
    $messageBox = New-Object System.Windows.Controls.TextBox
    $messageBox.AcceptsReturn = $true
    $messageBox.TextWrapping = "Wrap"
    $messageBox.Height = 80
    $messageBox.Text = $Message.Details
    $messageBox.ToolTip = "Message content for the announcement. Use Markdown for formatting."
    $messageBox.Add_TextChanged({ if (-not $script:IsLoading) { Update-Preview } })

    $enabledCheckBox = New-Object System.Windows.Controls.CheckBox
    $enabledCheckBox.Content = "Announcement Enabled"
    $enabledCheckBox.IsChecked = $Message.Enabled
    $enabledCheckBox.Margin = [System.Windows.Thickness]::new(0,5,0,0)
    $enabledCheckBox.ToolTip = "If checked, this announcement is active and may be shown in the Endpoint Advisor app if conditions are met."
    $enabledCheckBox.Add_Checked({ if (-not $global:IsLoading) { Update-Preview } })
    $enabledCheckBox.Add_Unchecked({ if (-not $script:IsLoading) { Update-Preview } })

    $appendCheckBox = New-Object System.Windows.Controls.CheckBox
    $appendCheckBox.Content = "Append to Default"
    $appendCheckBox.IsChecked = $Message.AppendToDefault
    $appendCheckBox.Margin = [System.Windows.Thickness]::new(0,5,0,0)
    $appendCheckBox.ToolTip = "If checked, this announcement appends to the default announcement; if unchecked, it may replace it on matching systems."
    $appendCheckBox.Add_Checked({ if (-not $global:IsLoading) { Update-Preview } })
    $appendCheckBox.Add_Unchecked({ if (-not $script:IsLoading) { Update-Preview } })

    $conditionPanel = New-Object System.Windows.Controls.StackPanel
    $conditionPanel.Orientation = "Vertical"
    $conditionPanel.Margin = [System.Windows.Thickness]::new(0,5,0,0)
    $conditionLabel = New-Object System.Windows.Controls.TextBlock
    $conditionLabel.Text = "Registry Condition set by BigFix on endpoint"
    $conditionLabel.FontWeight = "Bold"
    $pathBox = New-Object System.Windows.Controls.TextBox
    $pathBox.Width = 300
    $pathBox.Text = $Message.Condition.Path
    $pathBox.Margin = [System.Windows.Thickness]::new(0,0,0,5)
    $pathBox.ToolTip = "Registry path (e.g., HKLM:\SOFTWARE\Example)"
    $pathBox.Add_TextChanged({ if (-not $script:IsLoading) { Update-Preview } })
    $nameBox = New-Object System.Windows.Controls.TextBox
    $nameBox.Width = 150
    $nameBox.Text = $Message.Condition.Name
    $nameBox.Margin = [System.Windows.Thickness]::new(0,0,0,5)
    $nameBox.ToolTip = "Registry key name"
    $nameBox.Add_TextChanged({ if (-not $script:IsLoading) { Update-Preview } })
    $valueBox = New-Object System.Windows.Controls.TextBox
    $valueBox.Width = 150
    $valueBox.Text = $Message.Condition.Value
    $valueBox.ToolTip = "Registry key value to match"
    $valueBox.Add_TextChanged({ if (-not $script:IsLoading) { Update-Preview } })
    $conditionPanel.Children.Add($conditionLabel)
    $conditionPanel.Children.Add($pathBox)
    $conditionPanel.Children.Add($nameBox)
    $conditionPanel.Children.Add($valueBox)

    $linksPanel = New-Object System.Windows.Controls.StackPanel
    $linksLabel = New-Object System.Windows.Controls.TextBlock
    $linksLabel.Text = "Links"
    $linksLabel.FontWeight = "Bold"
    $linksLabel.Margin = [System.Windows.Thickness]::new(0,5,0,0)
    $linksSubPanel = New-Object System.Windows.Controls.StackPanel
    $linksSubPanel.Orientation = "Vertical"
    $linksSubPanel.Margin = [System.Windows.Thickness]::new(0,0,0,5)
    $linksButtonPanel = New-Object System.Windows.Controls.StackPanel
    $linksButtonPanel.Orientation = "Horizontal"
    $addLinkButton = New-Object System.Windows.Controls.Button
    $addLinkButton.Content = "Add Link"
    $addLinkButton.Width = 80
    $addLinkButton.ToolTip = "Add a new hyperlink to this targeted announcement."
    $removeLinkButton = New-Object System.Windows.Controls.Button
    $removeLinkButton.Content = "Remove Last Link"
    $removeLinkButton.Width = 120
    $removeLinkButton.ToolTip = "Remove the last hyperlink from this targeted announcement."
    $linksButtonPanel.Children.Add($addLinkButton)
    $linksButtonPanel.Children.Add($removeLinkButton)
    $linksPanel.Children.Add($linksLabel)
    $linksPanel.Children.Add($linksSubPanel)
    $linksPanel.Children.Add($linksButtonPanel)

    $removeButton = New-Object System.Windows.Controls.Button
    $removeButton.Content = "Remove This Announcement"
    $removeButton.Width = 150
    $removeButton.Margin = [System.Windows.Thickness]::new(0,5,0,0)
    $removeButton.ToolTip = "Remove this targeted announcement."

    $stackPanel.Children.Add($titleLabel)
    $stackPanel.Children.Add($titleBox)
    $stackPanel.Children.Add($messageLabel)
    $stackPanel.Children.Add($messageBox)
    $stackPanel.Children.Add($enabledCheckBox)
    $stackPanel.Children.Add($appendCheckBox)
    $stackPanel.Children.Add($conditionPanel)
    $stackPanel.Children.Add($linksPanel)
    $stackPanel.Children.Add($removeButton)
    $border.Child = $stackPanel
    $TargetedAnnouncementsPanel.Children.Add($border)

    $linkEditors = @()
    foreach ($link in $Message.Links) {
        $linkEditors += New-LinkEditor -Panel $linksSubPanel -Name $link.Name -Url $link.Url
    }

    $addLinkButton.Add_Click({
        $linkEditors += New-LinkEditor -Panel $linksSubPanel
        Update-Preview
    })
    $removeLinkButton.Add_Click({
        if ($linksSubPanel.Children.Count -gt 0) {
            $linksSubPanel.Children.RemoveAt($linksSubPanel.Children.Count - 1)
            $linkEditors = $linkEditors | Select-Object -First ($linkEditors.Count - 1)
            Update-Preview
        }
    })
    $removeButton.Add_Click({
        $TargetedAnnouncementsPanel.Children.Remove($border)
        $global:TargetedEditors = $global:TargetedEditors | Where-Object { $_.Border -ne $border }
        Update-Preview
    })

    $editor = @{
        Border = $border
        TitleBox = $titleBox
        MessageBox = $messageBox
        EnabledCheckBox = $enabledCheckBox
        AppendCheckBox = $appendCheckBox
        PathBox = $pathBox
        NameBox = $nameBox
        ValueBox = $valueBox
        LinksPanel = $linksSubPanel
        LinkEditors = [ref]$linkEditors
    }
    $global:TargetedEditors += $editor
    return $editor
}

# Function to update preview
function Update-Preview {
    $PreviewAnnouncementsText.Inlines.Clear()
    $PreviewAnnouncementsDetails.Inlines.Clear()
    $PreviewAppendedAnnouncements.Children.Clear()
    $PreviewAnnouncementsLinks.Children.Clear()
    $PreviewSupportText.Inlines.Clear()
    $PreviewSupportLinks.Children.Clear()

    $json = Get-CurrentJson
    if (-not $json) { return }

    # Always show Default Announcement
    Convert-MarkdownToTextBlock -Text $json.Announcements.Default.Text -TargetTextBlock $PreviewAnnouncementsText
    Convert-MarkdownToTextBlock -Text $json.Announcements.Default.Details -TargetTextBlock $PreviewAnnouncementsDetails

    # Append Targeted Announcements only if Enabled and AppendToDefault
    foreach ($targeted in $json.Announcements.Targeted) {
        if ($targeted.Enabled -and $targeted.AppendToDefault) {
            if (-not [string]::IsNullOrEmpty($targeted.Text)) {
                $sep = New-Object System.Windows.Controls.Separator
                $sep.Margin = [System.Windows.Thickness]::new(0,10,0,10)
                $PreviewAppendedAnnouncements.Children.Add($sep)
                $textBlock = New-Object System.Windows.Controls.TextBlock
                $textBlock.FontSize = 11
                $textBlock.TextWrapping = "Wrap"
                Convert-MarkdownToTextBlock -Text $targeted.Text -TargetTextBlock $textBlock
                $PreviewAppendedAnnouncements.Children.Add($textBlock)
            }
            if (-not [string]::IsNullOrEmpty($targeted.Details)) {
                $detailsBlock = New-Object System.Windows.Controls.TextBlock
                $detailsBlock.FontSize = 11
                $detailsBlock.TextWrapping = "Wrap"
                $detailsBlock.Margin = [System.Windows.Thickness]::new(0,5,0,0)
                Convert-MarkdownToTextBlock -Text $targeted.Details -TargetTextBlock $detailsBlock
                $PreviewAppendedAnnouncements.Children.Add($detailsBlock)
            }
        }
    }

    # Combine all links from enabled Targeted Announcements and Default
    $allLinks = [System.Collections.Generic.List[object]]::new()
    $allLinks.AddRange($json.Announcements.Default.Links)
    foreach ($targeted in $json.Announcements.Targeted) {
        if ($targeted.Enabled) {
            $allLinks.AddRange($targeted.Links)
        }
    }
    foreach ($link in $allLinks) {
        $tb = New-Object System.Windows.Controls.TextBlock
        $hp = New-Object System.Windows.Documents.Hyperlink
        $hp.NavigateUri = [Uri]$link.Url
        $hp.Inlines.Add($link.Name)
        $tb.Inlines.Add($hp)
        $PreviewAnnouncementsLinks.Children.Add($tb)
    }

    # Preview Support
    Convert-MarkdownToTextBlock -Text $json.Support.Text -TargetTextBlock $PreviewSupportText
    foreach ($link in $json.Support.Links) {
        $tb = New-Object System.Windows.Controls.TextBlock
        $hp = New-Object System.Windows.Documents.Hyperlink
        $hp.NavigateUri = [Uri]$link.Url
        $hp.Inlines.Add($link.Name)
        $tb.Inlines.Add($hp)
        $PreviewSupportLinks.Children.Add($tb)
    }
}

# Function to get current JSON from UI
function Get-CurrentJson {
    try {
        $json = [PSCustomObject]@{
            Announcements = [PSCustomObject]@{
                Default = [PSCustomObject]@{
                    Text = $DefaultText.Text
                    Details = $DefaultDetails.Text
                    Links = @()
                }
                Targeted = @()
            }
            Support = [PSCustomObject]@{
                Text = $SupportText.Text
                Links = @()
            }
        }

        foreach ($child in $DefaultLinksPanel.Children) {
            if ($child.Children[0].Text -and $child.Children[1].Text) {
                $json.Announcements.Default.Links += [PSCustomObject]@{ Name = $child.Children[0].Text; Url = $child.Children[1].Text }
            }
        }

        foreach ($editor in $global:TargetedEditors) {
            $targeted = [PSCustomObject]@{
                Text = $editor.TitleBox.Text
                Details = $editor.MessageBox.Text
                AppendToDefault = $editor.AppendCheckBox.IsChecked
                DisableAppended = -not $editor.AppendCheckBox.IsChecked
                Enabled = $editor.EnabledCheckBox.IsChecked
                Condition = [PSCustomObject]@{
                    Type = "Registry"
                    Path = $editor.PathBox.Text
                    Name = $editor.NameBox.Text
                    Value = $editor.ValueBox.Text
                }
                Links = @()
                Message = [PSCustomObject]@{
                    Text = $editor.TitleBox.Text
                    Details = $editor.MessageBox.Text
                    Links = @()
                }
            }
            foreach ($linkChild in $editor.LinksPanel.Children) {
                if ($linkChild.Children[0].Text -and $linkChild.Children[1].Text) {
                    $link = [PSCustomObject]@{ Name = $linkChild.Children[0].Text; Url = $linkChild.Children[1].Text }
                    $targeted.Links += $link
                    $targeted.Message.Links += $link
                }
            }
            $json.Announcements.Targeted += $targeted
        }

        foreach ($child in $SupportLinksPanel.Children) {
            if ($child.Children[0].Text -and $child.Children[1].Text) {
                $json.Support.Links += [PSCustomObject]@{ Name = $child.Children[0].Text; Url = $child.Children[1].Text }
            }
        }

        Write-Log "Successfully built JSON from UI"
        return $json
    } catch {
        Write-Log "Error building JSON: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Error building JSON: $($_.Exception.Message)", "Error", "OK", "Error")
        return $null
    }
}

# Load initial JSON into UI
function Load-JsonToUI {
    try {
        # Load JSON data
        $DefaultText.Text = $global:JsonData.Announcements.Default.Text
        $DefaultDetails.Text = $global:JsonData.Announcements.Default.Details
        $DefaultLinksPanel.Children.Clear()
        foreach ($link in $global:JsonData.Announcements.Default.Links) {
            New-LinkEditor -Panel $DefaultLinksPanel -Name $link.Name -Url $link.Url
        }
        $TargetedAnnouncementsPanel.Children.Clear()
        $global:TargetedEditors = @()
        foreach ($targeted in $global:JsonData.Announcements.Targeted) {
            $message = @{
                Text = if ($targeted.Message) { $targeted.Message.Text } else { $targeted.Text }
                Details = if ($targeted.Message) { $targeted.Message.Details } else { $targeted.Details }
                Links = if ($targeted.Message -and $targeted.Message.Links) { $targeted.Message.Links } elseif ($targeted.Links) { $targeted.Links } else { @() }
                AppendToDefault = if ($null -ne $targeted.AppendToDefault) { $targeted.AppendToDefault } else { $true }
                DisableAppended = if ($null -ne $targeted.DisableAppended) { $targeted.DisableAppended } else { $false }
                Enabled = if ($null -ne $targeted.Enabled) { $targeted.Enabled } else { $true }
                Condition = if ($targeted.Condition) { $targeted.Condition } else { @{ Type = "Registry"; Path = ""; Name = ""; Value = "" } }
            }
            New-TargetedAnnouncementEditor -Message $message
        }
        $SupportText.Text = $global:JsonData.Support.Text
        $SupportLinksPanel.Children.Clear()
        foreach ($link in $global:JsonData.Support.Links) {
            New-LinkEditor -Panel $SupportLinksPanel -Name $link.Name -Url $link.Url
        }

        # Render Help tab content
        if ($HelpRichTextBox) {
            Write-Log "Assigning help content to HelpRichTextBox."
            $HelpRichTextBox.Document.Blocks.Clear()
            $paragraph = New-Object System.Windows.Documents.Paragraph
            $paragraph.Inlines.Add((New-Object System.Windows.Documents.Run($helpContent)))
            $HelpRichTextBox.Document.Blocks.Add($paragraph)
            Write-Log "Help content assigned to RichTextBox document."
        } else {
            Write-Log "HelpRichTextBox control could not be found."
        }
        # Force a full UI layout update, then switch to the Help tab to force it to render.
        $window.UpdateLayout()
        $EditorTabs.SelectedIndex = 3 # Index of the "Help" tab
        $EditorTabs.SelectedIndex = 0 # Switch back to the first tab

        Write-Log "Successfully loaded JSON into UI"
    } catch {
        Write-Log "Failed to load JSON into UI: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Failed to load JSON into UI: $($_.Exception.Message)", "Error", "OK", "Error")
    }
}

# Event handlers
$GitHubPatTextBox.Add_KeyDown({
    if ($_.Key -eq "Return") {
        Save-P
