# ContentDataEditor.ps1
# Version 1.0 - Phase 1: Basic Load/Display

$ScriptVersion = "1.0.0"
$GitHubUrl = "https://raw.githubusercontent.com/burnoil/EndpointAdvisor/refs/heads/main/ContentData2.json"

# Import required assemblies
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# XAML Layout
$xamlString = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Endpoint Advisor Content Editor v$ScriptVersion"
    Width="900" Height="700"
    WindowStartupLocation="CenterScreen"
    ResizeMode="CanResize">
  <Grid Margin="10">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    
    <!-- Header -->
    <Border Grid.Row="0" Background="#0078D7" Padding="10" CornerRadius="3" Margin="0,0,0,10">
      <TextBlock Text="Endpoint Advisor Content Editor" FontSize="16" FontWeight="Bold" Foreground="White"/>
    </Border>
    
    <!-- Load Controls -->
    <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,10">
      <TextBlock Text="GitHub URL:" VerticalAlignment="Center" Margin="0,0,10,0"/>
      <TextBox x:Name="GitHubUrlBox" Width="500" VerticalAlignment="Center" Text="$GitHubUrl"/>
      <Button x:Name="LoadButton" Content="Load from GitHub" Margin="10,0,0,0" Padding="10,5"/>
      <Button x:Name="LoadFileButton" Content="Load from File" Margin="10,0,0,0" Padding="10,5"/>
    </StackPanel>
    
    <!-- Tab Control for Editing -->
    <TabControl x:Name="MainTabControl" Grid.Row="2">
      <TabItem Header="Raw JSON">
        <Grid>
          <TextBox x:Name="RawJsonBox" AcceptsReturn="True" VerticalScrollBarVisibility="Auto" FontFamily="Consolas" FontSize="11"/>
        </Grid>
      </TabItem>
      <TabItem Header="Announcements" x:Name="AnnouncementsTab" IsEnabled="False">
  <Grid>
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="*"/>
      <ColumnDefinition Width="5"/>
      <ColumnDefinition Width="*"/>
    </Grid.ColumnDefinitions>
    
    <!-- LEFT: Editor -->
    <ScrollViewer Grid.Column="0" VerticalScrollBarVisibility="Auto">
      <StackPanel Margin="10">
        <TextBlock Text="Edit Default Announcement" FontSize="14" FontWeight="Bold" Margin="0,0,0,10"/>
        
        <!-- Main Text -->
        <TextBlock Text="Main Text:" FontWeight="Bold" Margin="0,5,0,2"/>
        <TextBox x:Name="AnnouncementText" TextWrapping="Wrap" AcceptsReturn="True" Height="80" VerticalScrollBarVisibility="Auto"/>
        <TextBlock Text="Supports markdown: **bold**, *italic*, __underline__, [color]text[/color]" FontSize="9" Foreground="Gray" Margin="0,2,0,10"/>
        
        <!-- Details -->
        <TextBlock Text="Details (optional):" FontWeight="Bold" Margin="0,5,0,2"/>
        <TextBox x:Name="AnnouncementDetails" TextWrapping="Wrap" AcceptsReturn="True" Height="80" VerticalScrollBarVisibility="Auto"/>
        
        <!-- Links -->
        <TextBlock Text="Links:" FontWeight="Bold" Margin="0,10,0,5"/>
        <StackPanel x:Name="AnnouncementLinksPanel" Margin="0,0,0,10"/>
        <Button x:Name="AddAnnouncementLinkButton" Content="+ Add Link" Padding="5,2" Width="100" HorizontalAlignment="Left"/>
        
        <Separator Margin="0,20,0,20"/>
        
        <!-- Targeted Announcements -->
        <TextBlock Text="Targeted Announcements" FontSize="14" FontWeight="Bold" Margin="0,0,0,5"/>
        <TextBlock Text="Show different announcements based on registry conditions" FontSize="10" Foreground="Gray" Margin="0,0,0,5"/>

        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="200"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>
          
          <Border Grid.Column="0" BorderBrush="Gray" BorderThickness="0,0,1,0" Padding="5" Margin="0,0,10,0">
            <StackPanel>
              <TextBlock Text="Targeted Messages" FontWeight="Bold" Margin="0,0,0,5"/>
              <ListBox x:Name="TargetedAnnouncementsList" Height="200" Margin="0,0,0,5"/>
              <Button x:Name="AddTargetedButton" Content="+ Add Targeted" Padding="5,2" Margin="0,0,0,5"/>
              <Button x:Name="RemoveTargetedButton" Content="Remove Selected" Padding="5,2"/>
            </StackPanel>
          </Border>
          
          <StackPanel Grid.Column="1" x:Name="TargetedEditorPanel" IsEnabled="False">
            <CheckBox x:Name="TargetedEnabledCheck" Content="Enabled" Margin="0,0,0,5"/>
            <CheckBox x:Name="TargetedAppendCheck" Content="Append to Default" Margin="0,0,0,10"/>
            
            <TextBlock Text="Registry Path:" Margin="0,5,0,2"/>
            <TextBox x:Name="TargetedConditionPathBox" Margin="0,0,0,5"/>
            <TextBlock Text="Registry Name:" Margin="0,5,0,2"/>
            <TextBox x:Name="TargetedConditionNameBox" Margin="0,0,0,5"/>
            <TextBlock Text="Expected Value:" Margin="0,5,0,2"/>
            <TextBox x:Name="TargetedConditionValueBox" Margin="0,0,0,10"/>
            
            <TextBlock Text="Text:" FontWeight="Bold" Margin="0,5,0,2"/>
            <TextBox x:Name="TargetedTextBox" TextWrapping="Wrap" AcceptsReturn="True" Height="60" Margin="0,0,0,5"/>
            
            <TextBlock Text="Details:" FontWeight="Bold" Margin="0,5,0,2"/>
            <TextBox x:Name="TargetedDetailsBox" TextWrapping="Wrap" AcceptsReturn="True" Height="60" Margin="0,0,0,5"/>
            
            <TextBlock Text="Links:" FontWeight="Bold" Margin="0,5,0,5"/>
            <StackPanel x:Name="TargetedLinksPanel" Margin="0,0,0,5"/>
            <Button x:Name="AddTargetedLinkButton" Content="+ Add Link" Padding="5,2" Width="100" HorizontalAlignment="Left" Margin="0,0,0,10"/>
            
            <Button x:Name="SaveTargetedButton" Content="Save Targeted" Padding="10,5"/>
          </StackPanel>
        </Grid>
        
        <Separator Margin="0,20,0,20"/>
        
        <StackPanel Orientation="Horizontal">
          <Button x:Name="SaveAnnouncementsButton" Content="Save Changes" Padding="10,5" Margin="0,0,10,0"/>
          <Button x:Name="RevertAnnouncementsButton" Content="Revert Changes" Padding="10,5"/>
        </StackPanel>
      </StackPanel>
    </ScrollViewer>
    
    <!-- SPLITTER -->
    <GridSplitter Grid.Column="1" Width="5" HorizontalAlignment="Stretch" Background="#CCCCCC"/>
    
    <!-- RIGHT: Live Preview -->
    <Border Grid.Column="2" Background="#F5F5F5" BorderBrush="#CCCCCC" BorderThickness="1,0,0,0">
      <ScrollViewer VerticalScrollBarVisibility="Auto">
        <StackPanel Margin="10">
          <TextBlock Text="LIVE PREVIEW" FontSize="12" FontWeight="Bold" Foreground="#0078D7" Margin="0,0,0,10"/>
          
          <Border Background="White" Padding="15" BorderBrush="#DDDDDD" BorderThickness="1" CornerRadius="3">
            <StackPanel x:Name="AnnouncementsPreviewPanel">
              <TextBlock Text="Edit content on the left to see preview" FontStyle="Italic" Foreground="Gray"/>
            </StackPanel>
          </Border>
        </StackPanel>
      </ScrollViewer>
    </Border>
  </Grid>
</TabItem>
      <TabItem Header="Support" x:Name="SupportTab" IsEnabled="False">
  <Grid>
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="*"/>
      <ColumnDefinition Width="5"/>
      <ColumnDefinition Width="*"/>
    </Grid.ColumnDefinitions>
    
    <!-- LEFT: Editor -->
    <ScrollViewer Grid.Column="0" VerticalScrollBarVisibility="Auto">
      <StackPanel Margin="10">
        <TextBlock Text="Support Information" FontSize="14" FontWeight="Bold" Margin="0,0,0,10"/>
        
        <!-- Main Text -->
        <TextBlock Text="Support Text:" FontWeight="Bold" Margin="0,5,0,2"/>
        <TextBox x:Name="SupportText" TextWrapping="Wrap" AcceptsReturn="True" Height="100" VerticalScrollBarVisibility="Auto"/>
        <TextBlock Text="Supports markdown: **bold**, *italic*, __underline__, [color]text[/color]" FontSize="9" Foreground="Gray" Margin="0,2,0,10"/>
        
        <!-- Links -->
        <TextBlock Text="Links:" FontWeight="Bold" Margin="0,10,0,5"/>
        <StackPanel x:Name="SupportLinksPanel" Margin="0,0,0,10"/>
        <Button x:Name="AddSupportLinkButton" Content="+ Add Link" Padding="5,2" Width="100" HorizontalAlignment="Left"/>
        
        <Separator Margin="0,20,0,20"/>
        
        <!-- Action Buttons -->
        <StackPanel Orientation="Horizontal">
          <Button x:Name="SaveSupportButton" Content="Save Changes" Padding="10,5" Margin="0,0,10,0"/>
          <Button x:Name="RevertSupportButton" Content="Revert Changes" Padding="10,5"/>
        </StackPanel>
      </StackPanel>
    </ScrollViewer>
    
    <!-- SPLITTER -->
    <GridSplitter Grid.Column="1" Width="5" HorizontalAlignment="Stretch" Background="#CCCCCC"/>
    
    <!-- RIGHT: Live Preview -->
    <Border Grid.Column="2" Background="#F5F5F5" BorderBrush="#CCCCCC" BorderThickness="1,0,0,0">
      <ScrollViewer VerticalScrollBarVisibility="Auto">
        <StackPanel Margin="10">
          <TextBlock Text="LIVE PREVIEW" FontSize="12" FontWeight="Bold" Foreground="#0078D7" Margin="0,0,0,10"/>
          
          <Border Background="White" Padding="15" BorderBrush="#DDDDDD" BorderThickness="1" CornerRadius="3">
            <StackPanel x:Name="SupportPreviewPanel">
              <TextBlock Text="Edit content on the left to see preview" FontStyle="Italic" Foreground="Gray"/>
            </StackPanel>
          </Border>
        </StackPanel>
      </ScrollViewer>
    </Border>
  </Grid>
</TabItem>
      <TabItem Header="Additional Tabs" x:Name="AdditionalTabsTab" IsEnabled="False">
  <Grid>
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="*"/>
      <ColumnDefinition Width="5"/>
      <ColumnDefinition Width="*"/>
    </Grid.ColumnDefinitions>
    
    <!-- LEFT: Editor -->
    <Grid Grid.Column="0">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="200"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>
      
      <!-- Tab List -->
      <Border Grid.Column="0" BorderBrush="Gray" BorderThickness="0,0,1,0" Padding="5">
        <StackPanel>
          <TextBlock Text="Dynamic Tabs" FontWeight="Bold" Margin="0,0,0,5"/>
          <ListBox x:Name="TabsList" Height="500" Margin="0,0,0,5"/>
          <Button x:Name="AddTabButton" Content="+ Add New Tab" Padding="5" Margin="0,5,0,0"/>
          <Button x:Name="RemoveTabButton" Content="Remove Selected" Padding="5" Margin="0,5,0,0"/>
        </StackPanel>
      </Border>
      
      <!-- Tab Editor -->
      <ScrollViewer Grid.Column="1" VerticalScrollBarVisibility="Auto">
        <StackPanel x:Name="TabEditorPanel" Margin="10" IsEnabled="False">
          <TextBlock Text="Tab Settings" FontSize="14" FontWeight="Bold" Margin="0,0,0,10"/>
          
          <TextBlock Text="Tab Header (Name):" FontWeight="Bold" Margin="0,5,0,2"/>
          <TextBox x:Name="TabHeaderBox" Margin="0,0,0,10"/>
          
          <CheckBox x:Name="TabEnabledCheck" Content="Enabled (visible to users)" Margin="0,0,0,10"/>
          
          <TextBlock Text="Content Text:" FontWeight="Bold" Margin="0,5,0,2"/>
          <TextBox x:Name="TabContentTextBox" TextWrapping="Wrap" AcceptsReturn="True" Height="80" VerticalScrollBarVisibility="Auto" Margin="0,0,0,10"/>
          
          <TextBlock Text="Content Links:" FontWeight="Bold" Margin="0,10,0,5"/>
          <StackPanel x:Name="TabContentLinksPanel" Margin="0,0,0,5"/>
          <Button x:Name="AddTabContentLinkButton" Content="+ Add Link" Padding="5,2" Width="100" HorizontalAlignment="Left" Margin="0,0,0,10"/>
          
          <Separator Margin="0,10,0,10"/>
          
          <TextBlock Text="Sections (optional expandable content)" FontSize="12" FontWeight="Bold" Margin="0,0,0,5"/>
          <TextBlock Text="Coming soon..." FontSize="10" Foreground="Gray" Margin="0,0,0,10"/>
          
          <Separator Margin="0,20,0,20"/>
          
          <StackPanel Orientation="Horizontal">
            <Button x:Name="SaveTabButton" Content="Save Tab Changes" Padding="10,5" Margin="0,0,10,0"/>
            <Button x:Name="RevertTabButton" Content="Revert Tab" Padding="10,5"/>
          </StackPanel>
        </StackPanel>
      </ScrollViewer>
    </Grid>
    
    <!-- SPLITTER -->
    <GridSplitter Grid.Column="1" Width="5" HorizontalAlignment="Stretch" Background="#CCCCCC"/>
    
    <!-- RIGHT: Live Preview -->
    <Border Grid.Column="2" Background="#F5F5F5" BorderBrush="#CCCCCC" BorderThickness="1,0,0,0">
      <ScrollViewer VerticalScrollBarVisibility="Auto">
        <StackPanel Margin="10">
          <TextBlock Text="LIVE PREVIEW" FontSize="12" FontWeight="Bold" Foreground="#0078D7" Margin="0,0,0,10"/>
          
          <Border Background="White" Padding="15" BorderBrush="#DDDDDD" BorderThickness="1" CornerRadius="3">
            <StackPanel x:Name="AdditionalTabsPreviewPanel">
              <TextBlock Text="Select a tab to preview" FontStyle="Italic" Foreground="Gray"/>
            </StackPanel>
          </Border>
        </StackPanel>
      </ScrollViewer>
    </Border>
  </Grid>
</TabItem>

<TabItem Header="GitHub Settings">
  <ScrollViewer VerticalScrollBarVisibility="Auto">
    <StackPanel Margin="10">
      <TextBlock Text="GitHub Configuration" FontSize="14" FontWeight="Bold" Margin="0,0,0,10"/>
      
      <TextBlock Text="Repository Information" FontWeight="Bold" Margin="0,10,0,5"/>
      <TextBlock Text="Owner/Organization:" Margin="0,5,0,2"/>
      <TextBox x:Name="GitHubOwnerBox" Text="burnoil" Margin="0,0,0,5"/>
      
      <TextBlock Text="Repository Name:" Margin="0,5,0,2"/>
      <TextBox x:Name="GitHubRepoBox" Text="EndpointAdvisor" Margin="0,0,0,5"/>
      
      <TextBlock Text="Branch:" Margin="0,5,0,2"/>
      <TextBox x:Name="GitHubBranchBox" Text="main" Margin="0,0,0,5"/>
      
      <TextBlock Text="File Path:" Margin="0,5,0,2"/>
      <TextBox x:Name="GitHubFilePathBox" Text="ContentData2.json" Margin="0,0,0,5"/>
      
      <Separator Margin="0,20,0,20"/>
      
      <TextBlock Text="Personal Access Token (PAT)" FontWeight="Bold" Margin="0,10,0,5"/>
      <TextBlock TextWrapping="Wrap" FontSize="10" Foreground="Gray" Margin="0,0,0,5">
        Required for committing changes. Create at: https://github.com/settings/tokens
        <LineBreak/>Needs 'repo' scope (read/write repository access)
      </TextBlock>
      
      <TextBox x:Name="GitHubTokenBox" Margin="0,5,0,5"/>
      <StackPanel Orientation="Horizontal">
        <Button x:Name="SaveGitHubConfigButton" Content="Save Configuration" Padding="10,5" Margin="0,0,10,0"/>
        <Button x:Name="TestGitHubConnectionButton" Content="Test Connection" Padding="10,5"/>
      </StackPanel>
    </StackPanel>
  </ScrollViewer>
</TabItem>
    </TabControl>
    
    <!-- Status Bar -->
    <!-- Action Bar -->
<Grid Grid.Row="3" Margin="0,10,0,0">
  <Grid.RowDefinitions>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="Auto"/>
  </Grid.RowDefinitions>
  
  <!-- Commit Section -->
  <Border Grid.Row="0" Background="#E8F4F8" Padding="10" BorderBrush="#0078D7" BorderThickness="1" Margin="0,0,0,5">
    <StackPanel>
      <TextBlock Text="Commit Message:" FontWeight="Bold" Margin="0,0,0,5"/>
      <TextBox x:Name="CommitMessageBox" TextWrapping="Wrap" Height="50" Margin="0,0,0,5"/>
      <StackPanel Orientation="Horizontal">
        <Button x:Name="CommitToGitHubButton" Content="Commit to GitHub" Padding="10,5" Background="#0078D7" Foreground="White" Margin="0,0,10,0" IsEnabled="False"/>
        <Button x:Name="SaveToFileButton" Content="Save to Local File" Padding="10,5" Margin="0,0,10,0"/>
        <TextBlock x:Name="CommitStatusText" VerticalAlignment="Center" Margin="10,0,0,0" Foreground="Gray"/>
      </StackPanel>
    </StackPanel>
  </Border>
  
  <!-- Status Bar -->
  <Border Grid.Row="1" Background="#F0F0F0" Padding="5">
    <TextBlock x:Name="StatusText" Text="Ready" FontSize="10"/>
  </Border>
</Grid>
  </Grid>
</Window>
"@

# Load XAML
try {
    $xmlDoc = New-Object System.Xml.XmlDocument
    $xmlDoc.LoadXml($xamlString)
    $reader = New-Object System.Xml.XmlNodeReader $xmlDoc
    [System.Windows.Window]$window = [Windows.Markup.XamlReader]::Load($reader)
    
    # Get UI elements
    $GitHubUrlBox = $window.FindName("GitHubUrlBox")
    $LoadButton = $window.FindName("LoadButton")
    $LoadFileButton = $window.FindName("LoadFileButton")
    $RawJsonBox = $window.FindName("RawJsonBox")
    $StatusText = $window.FindName("StatusText")
    $MainTabControl = $window.FindName("MainTabControl")
    
} catch {
    Write-Host "Failed to load XAML: $($_.Exception.Message)"
    exit
}

# Global variable to store loaded JSON
$global:ContentData = $null

# Function: Load JSON from GitHub
function Load-FromGitHub {
    try {
        $StatusText.Text = "Loading from GitHub..."
        $url = $GitHubUrlBox.Text
        
        $response = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 30
        $jsonText = $response.Content
        
        # Validate JSON
        $jsonObject = $jsonText | ConvertFrom-Json
        
        # Store and display
        $global:ContentData = $jsonObject
        $RawJsonBox.Text = $jsonText | ConvertFrom-Json | ConvertTo-Json -Depth 100
        
        $StatusText.Text = "Successfully loaded from GitHub"
        $StatusText.Foreground = [System.Windows.Media.Brushes]::Green
        
        # Enable editing tabs
        for ($i = 1; $i -lt $MainTabControl.Items.Count; $i++) {
            $MainTabControl.Items[$i].IsEnabled = $true
        }
        # Populate the announcements editor
Populate-AnnouncementsEditor
Populate-SupportEditor
Populate-AdditionalTabsList
Update-AnnouncementsPreview
Update-SupportPreview
# Populate targeted announcements list
Populate-TargetedAnnouncementsList
    } catch {
        $StatusText.Text = "Error loading from GitHub: $($_.Exception.Message)"
        $StatusText.Foreground = [System.Windows.Media.Brushes]::Red
    }
}

# Function: Load JSON from local file
function Load-FromFile {
    try {
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
        $openFileDialog.Title = "Select ContentData2.json file"
        
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $StatusText.Text = "Loading from file..."
            
            $jsonText = Get-Content $openFileDialog.FileName -Raw
            
            # Validate JSON
            $jsonObject = $jsonText | ConvertFrom-Json
            
            # Store and display
            $global:ContentData = $jsonObject
            $RawJsonBox.Text = $jsonText | ConvertFrom-Json | ConvertTo-Json -Depth 100
            
            $StatusText.Text = "Successfully loaded from file: $($openFileDialog.FileName)"
            $StatusText.Foreground = [System.Windows.Media.Brushes]::Green
            
            # Enable editing tabs
            for ($i = 1; $i -lt $MainTabControl.Items.Count; $i++) {
                $MainTabControl.Items[$i].IsEnabled = $true
            }
        }
        # Populate the announcements editor
Populate-AnnouncementsEditor
Populate-SupportEditor
Populate-AdditionalTabsList
# Populate targeted announcements list
Populate-TargetedAnnouncementsList
    } catch {
        $StatusText.Text = "Error loading from file: $($_.Exception.Message)"
        $StatusText.Foreground = [System.Windows.Media.Brushes]::Red
    }
}

# Get new UI elements
$AnnouncementText = $window.FindName("AnnouncementText")
$AnnouncementDetails = $window.FindName("AnnouncementDetails")
$AnnouncementLinksPanel = $window.FindName("AnnouncementLinksPanel")
$AddAnnouncementLinkButton = $window.FindName("AddAnnouncementLinkButton")
$SaveAnnouncementsButton = $window.FindName("SaveAnnouncementsButton")
$RevertAnnouncementsButton = $window.FindName("RevertAnnouncementsButton")
$AnnouncementsPreviewPanel = $window.FindName("AnnouncementsPreviewPanel")
# Update preview when text changes
$AnnouncementText.Add_TextChanged({ Update-AnnouncementsPreview })
$AnnouncementDetails.Add_TextChanged({ Update-AnnouncementsPreview })

$SupportText = $window.FindName("SupportText")
$SupportLinksPanel = $window.FindName("SupportLinksPanel")
$AddSupportLinkButton = $window.FindName("AddSupportLinkButton")
$SaveSupportButton = $window.FindName("SaveSupportButton")
$RevertSupportButton = $window.FindName("RevertSupportButton")

$TabsList = $window.FindName("TabsList")
$AddTabButton = $window.FindName("AddTabButton")
$RemoveTabButton = $window.FindName("RemoveTabButton")
$TabEditorPanel = $window.FindName("TabEditorPanel")
$TabHeaderBox = $window.FindName("TabHeaderBox")
$TabEnabledCheck = $window.FindName("TabEnabledCheck")
$TabContentTextBox = $window.FindName("TabContentTextBox")
$TabContentLinksPanel = $window.FindName("TabContentLinksPanel")
$AddTabContentLinkButton = $window.FindName("AddTabContentLinkButton")
$SaveTabButton = $window.FindName("SaveTabButton")
$RevertTabButton = $window.FindName("RevertTabButton")

$TargetedAnnouncementsList = $window.FindName("TargetedAnnouncementsList")
$AddTargetedButton = $window.FindName("AddTargetedButton")
$RemoveTargetedButton = $window.FindName("RemoveTargetedButton")
$TargetedEditorPanel = $window.FindName("TargetedEditorPanel")
$TargetedEnabledCheck = $window.FindName("TargetedEnabledCheck")
$TargetedAppendCheck = $window.FindName("TargetedAppendCheck")
$TargetedConditionPathBox = $window.FindName("TargetedConditionPathBox")
$TargetedConditionNameBox = $window.FindName("TargetedConditionNameBox")
$TargetedConditionValueBox = $window.FindName("TargetedConditionValueBox")
$TargetedTextBox = $window.FindName("TargetedTextBox")
$TargetedDetailsBox = $window.FindName("TargetedDetailsBox")
$TargetedLinksPanel = $window.FindName("TargetedLinksPanel")
$AddTargetedLinkButton = $window.FindName("AddTargetedLinkButton")
$SaveTargetedButton = $window.FindName("SaveTargetedButton")

$GitHubOwnerBox = $window.FindName("GitHubOwnerBox")
$GitHubRepoBox = $window.FindName("GitHubRepoBox")
$GitHubBranchBox = $window.FindName("GitHubBranchBox")
$GitHubFilePathBox = $window.FindName("GitHubFilePathBox")
$GitHubTokenBox = $window.FindName("GitHubTokenBox")
$SaveGitHubConfigButton = $window.FindName("SaveGitHubConfigButton")
$TestGitHubConnectionButton = $window.FindName("TestGitHubConnectionButton")
$CommitMessageBox = $window.FindName("CommitMessageBox")
$CommitToGitHubButton = $window.FindName("CommitToGitHubButton")
$SaveToFileButton = $window.FindName("SaveToFileButton")
$CommitStatusText = $window.FindName("CommitStatusText")

$SupportPreviewPanel = $window.FindName("SupportPreviewPanel")
$AdditionalTabsPreviewPanel = $window.FindName("AdditionalTabsPreviewPanel")

# Function: Update Support Live Preview
function Update-SupportPreview {
    try {
        $SupportPreviewPanel.Children.Clear()
        
        # Title
        $title = New-Object System.Windows.Controls.TextBlock
        $title.Text = "Support"
        $title.FontSize = 12
        $title.FontWeight = "Bold"
        $title.Margin = [System.Windows.Thickness]::new(0,0,0,10)
        $SupportPreviewPanel.Children.Add($title)
        
        # Border
        $border = New-Object System.Windows.Controls.Border
        $border.BorderBrush = [System.Windows.Media.Brushes]::DarkBlue
        $border.BorderThickness = [System.Windows.Thickness]::new(2)
        $border.Padding = [System.Windows.Thickness]::new(8)
        $border.CornerRadius = 3
        
        $contentPanel = New-Object System.Windows.Controls.StackPanel
        
        # Text with markdown
        if ($SupportText.Text) {
            $textBlock = New-Object System.Windows.Controls.TextBlock
            $textBlock.TextWrapping = "Wrap"
            $textBlock.Margin = [System.Windows.Thickness]::new(0,0,0,10)
            $renderedText = Render-SimpleMarkdown -Text $SupportText.Text
            $textBlock.Inlines.AddRange($renderedText)
            $contentPanel.Children.Add($textBlock)
        }
        
        # Links
        foreach ($grid in $SupportLinksPanel.Children) {
            $nameBox = $grid.Children | Where-Object { $_.Tag -eq "LinkName" }
            $urlBox = $grid.Children | Where-Object { $_.Tag -eq "LinkUrl" }
            
            if ($nameBox.Text -and $urlBox.Text) {
                $linkButton = New-Object System.Windows.Controls.Button
                $linkButton.Content = "[Link] $($nameBox.Text)"
                $linkButton.Margin = [System.Windows.Thickness]::new(0,2,0,2)
                $linkButton.Padding = [System.Windows.Thickness]::new(5,2,5,2)
                $linkButton.HorizontalAlignment = "Left"
                $linkButton.ToolTip = $urlBox.Text
                $contentPanel.Children.Add($linkButton)
            }
        }
        
        $border.Child = $contentPanel
        $SupportPreviewPanel.Children.Add($border)
        
    } catch {
        # Silently fail - preview is non-critical
    }
}

# Function: Update Additional Tabs Live Preview
function Update-AdditionalTabsPreview {
    try {
        $AdditionalTabsPreviewPanel.Children.Clear()
        
        if ($global:CurrentEditingTabIndex -lt 0) {
            $noSelection = New-Object System.Windows.Controls.TextBlock
            $noSelection.Text = "Select a tab from the list to preview"
            $noSelection.FontStyle = "Italic"
            $noSelection.Foreground = [System.Windows.Media.Brushes]::Gray
            $AdditionalTabsPreviewPanel.Children.Add($noSelection)
            return
        }
        
        # Tab header
        $header = New-Object System.Windows.Controls.TextBlock
        $header.Text = if ($TabHeaderBox.Text) { $TabHeaderBox.Text } else { "Tab Name" }
        $header.FontSize = 14
        $header.FontWeight = "Bold"
        $header.Margin = [System.Windows.Thickness]::new(0,0,0,5)
        $AdditionalTabsPreviewPanel.Children.Add($header)
        
        $separator = New-Object System.Windows.Controls.Separator
        $separator.Margin = [System.Windows.Thickness]::new(0,0,0,10)
        $AdditionalTabsPreviewPanel.Children.Add($separator)
        
        # Enabled status
        if (-not $TabEnabledCheck.IsChecked) {
            $disabledNote = New-Object System.Windows.Controls.TextBlock
            $disabledNote.Text = "[This tab is DISABLED and will not be visible to users]"
            $disabledNote.FontSize = 10
            $disabledNote.FontWeight = "Bold"
            $disabledNote.Foreground = [System.Windows.Media.Brushes]::Red
            $disabledNote.Margin = [System.Windows.Thickness]::new(0,0,0,10)
            $AdditionalTabsPreviewPanel.Children.Add($disabledNote)
        }
        
        # Content text
        if ($TabContentTextBox.Text) {
            $textBlock = New-Object System.Windows.Controls.TextBlock
            $textBlock.Text = $TabContentTextBox.Text
            $textBlock.TextWrapping = "Wrap"
            $textBlock.Margin = [System.Windows.Thickness]::new(0,0,0,10)
            $AdditionalTabsPreviewPanel.Children.Add($textBlock)
        }
        
        # Content links
        foreach ($grid in $TabContentLinksPanel.Children) {
            $nameBox = $grid.Children | Where-Object { $_.Tag -eq "LinkName" }
            $urlBox = $grid.Children | Where-Object { $_.Tag -eq "LinkUrl" }
            
            if ($nameBox.Text -and $urlBox.Text) {
                $linkButton = New-Object System.Windows.Controls.Button
                $linkButton.Content = "[Link] $($nameBox.Text)"
                $linkButton.Margin = [System.Windows.Thickness]::new(0,2,0,2)
                $linkButton.Padding = [System.Windows.Thickness]::new(5,2,5,2)
                $linkButton.HorizontalAlignment = "Left"
                $linkButton.ToolTip = $urlBox.Text
                $AdditionalTabsPreviewPanel.Children.Add($linkButton)
            }
        }
        
    } catch {
        # Silently fail - preview is non-critical
    }
}

# Function: Update Announcements Live Preview
function Update-AnnouncementsPreview {
    try {
        $AnnouncementsPreviewPanel.Children.Clear()
        
        # Title
        $title = New-Object System.Windows.Controls.TextBlock
        $title.Text = "Announcements"
        $title.FontSize = 12
        $title.FontWeight = "Bold"
        $title.Margin = [System.Windows.Thickness]::new(0,0,0,10)
        $AnnouncementsPreviewPanel.Children.Add($title)
        
        # Expander
        $expander = New-Object System.Windows.Controls.Expander
        $expander.Header = "System Announcements"
        $expander.IsExpanded = $true
        $expander.BorderBrush = [System.Windows.Media.Brushes]::LightGray
        $expander.BorderThickness = [System.Windows.Thickness]::new(1)
        $expander.Margin = [System.Windows.Thickness]::new(0,0,0,10)
        
        $expanderContent = New-Object System.Windows.Controls.StackPanel
        $expanderContent.Margin = [System.Windows.Thickness]::new(5)
        
        # Main text with markdown
        if ($AnnouncementText.Text) {
            $textBlock = New-Object System.Windows.Controls.TextBlock
            $textBlock.TextWrapping = "Wrap"
            $textBlock.Margin = [System.Windows.Thickness]::new(0,5,0,5)
            $renderedText = Render-SimpleMarkdown -Text $AnnouncementText.Text
            $textBlock.Inlines.AddRange($renderedText)
            $expanderContent.Children.Add($textBlock)
        }
        
        # Details
        # Details
# Details
if ($AnnouncementDetails.Text) {
    $detailsBlock = New-Object System.Windows.Controls.TextBlock
    $detailsBlock.TextWrapping = "Wrap"
    $detailsBlock.FontSize = 12  # Larger size
    $detailsBlock.Foreground = [System.Windows.Media.Brushes]::DarkGray  # Darker gray
    $detailsBlock.Margin = [System.Windows.Thickness]::new(0,5,0,5)
    
    # Apply markdown rendering
    $renderedDetails = Render-SimpleMarkdown -Text $AnnouncementDetails.Text
    $detailsBlock.Inlines.AddRange($renderedDetails)
    
    $expanderContent.Children.Add($detailsBlock)
}
        
        # Links
        $hasLinks = $false
        foreach ($grid in $AnnouncementLinksPanel.Children) {
            $nameBox = $grid.Children | Where-Object { $_.Tag -eq "LinkName" }
            $urlBox = $grid.Children | Where-Object { $_.Tag -eq "LinkUrl" }
            
            if ($nameBox.Text -and $urlBox.Text) {
                if (-not $hasLinks) {
                    $linksPanel = New-Object System.Windows.Controls.StackPanel
                    $linksPanel.Margin = [System.Windows.Thickness]::new(0,10,0,0)
                    $hasLinks = $true
                }
                
                $linkButton = New-Object System.Windows.Controls.Button
                $linkButton.Content = "[Link] $($nameBox.Text)"
                $linkButton.Margin = [System.Windows.Thickness]::new(0,2,0,2)
                $linkButton.Padding = [System.Windows.Thickness]::new(5,2,5,2)
                $linkButton.HorizontalAlignment = "Left"
                $linkButton.ToolTip = $urlBox.Text
                $linksPanel.Children.Add($linkButton)
            }
        }
        
        if ($hasLinks) {
            $expanderContent.Children.Add($linksPanel)
        }
        
        $expander.Content = $expanderContent
        $AnnouncementsPreviewPanel.Children.Add($expander)
        
    } catch {
        # Silently fail - preview is non-critical
    }
}

# Function: Simple markdown renderer (FIXED SYNTAX)
function Render-SimpleMarkdown {
    param([string]$Text)
    
    if (-not $Text) { 
        return @()
    }
    
    $inlines = New-Object System.Collections.ArrayList
    $lines = $Text -split "`r?`n"
    
    foreach ($line in $lines) {
        if (-not $line) {
            $lineBreak = New-Object System.Windows.Documents.LineBreak
            $inlines.Add($lineBreak) | Out-Null
            continue
        }
        
        # Find all formatting in order: colors first, then bold/italic/underline within
        $i = 0
        $currentColor = $null
        
        while ($i -lt $line.Length) {
            $remainder = $line.Substring($i)
            
            # Check for color start
            if ($remainder -match '^\[(\w+)\]') {
                $currentColor = $matches[1]
                $i += $matches[0].Length
                continue
            }
            
            # Check for color end
            if ($currentColor -and $remainder -match '^\[/' + [regex]::Escape($currentColor) + '\]') {
                $currentColor = $null
                $i += $matches[0].Length
                continue
            }
            
            # Check for bold **
            if (($remainder.Length -ge 2) -and ($remainder.Substring(0,2) -eq '**')) {
                $endIdx = $remainder.IndexOf('**', 2)
                if ($endIdx -gt 0) {
                    $text = $remainder.Substring(2, $endIdx - 2)
                    $run = New-Object System.Windows.Documents.Run
                    $run.Text = $text
                    $run.FontWeight = "Bold"
                    if ($currentColor) {
                        try {
                            $cName = $currentColor.Substring(0,1).ToUpper() + $currentColor.Substring(1).ToLower()
                            $run.Foreground = [System.Windows.Media.Brushes]::($cName)
                        } catch {}
                    }
                    $inlines.Add($run) | Out-Null
                    $i += $endIdx + 2
                    continue
                }
            }
            
            # Check for underline __
            if (($remainder.Length -ge 2) -and ($remainder.Substring(0,2) -eq '__')) {
                $endIdx = $remainder.IndexOf('__', 2)
                if ($endIdx -gt 0) {
                    $text = $remainder.Substring(2, $endIdx - 2)
                    $run = New-Object System.Windows.Documents.Run
                    $run.Text = $text
                    $run.TextDecorations = [System.Windows.TextDecorations]::Underline
                    if ($currentColor) {
                        try {
                            $cName = $currentColor.Substring(0,1).ToUpper() + $currentColor.Substring(1).ToLower()
                            $run.Foreground = [System.Windows.Media.Brushes]::($cName)
                        } catch {}
                    }
                    $inlines.Add($run) | Out-Null
                    $i += $endIdx + 2
                    continue
                }
            }
            
            # Check for italic *
            if (($remainder.Length -ge 1) -and ($remainder[0] -eq '*')) {
                $endIdx = $remainder.IndexOf('*', 1)
                if ($endIdx -gt 0) {
                    $text = $remainder.Substring(1, $endIdx - 1)
                    $run = New-Object System.Windows.Documents.Run
                    $run.Text = $text
                    $run.FontStyle = "Italic"
                    if ($currentColor) {
                        try {
                            $cName = $currentColor.Substring(0,1).ToUpper() + $currentColor.Substring(1).ToLower()
                            $run.Foreground = [System.Windows.Media.Brushes]::($cName)
                        } catch {}
                    }
                    $inlines.Add($run) | Out-Null
                    $i += $endIdx + 1
                    continue
                }
            }
            
            # Plain character
            $run = New-Object System.Windows.Documents.Run
            $run.Text = $line[$i].ToString()
            if ($currentColor) {
                try {
                    $cName = $currentColor.Substring(0,1).ToUpper() + $currentColor.Substring(1).ToLower()
                    $run.Foreground = [System.Windows.Media.Brushes]::($cName)
                } catch {}
            }
            $inlines.Add($run) | Out-Null
            $i++
        }
        
        # Line break
        $lineBreak = New-Object System.Windows.Documents.LineBreak
        $inlines.Add($lineBreak) | Out-Null
    }
    
    return $inlines
}

# Function: Populate Announcements editor from loaded JSON
function Populate-AnnouncementsEditor {
    try {
        $announcements = $global:ContentData.Dashboard.Announcements.Default
        
        $AnnouncementText.Text = $announcements.Text
        $AnnouncementDetails.Text = $announcements.Details
        
        # Clear and populate links
        $AnnouncementLinksPanel.Children.Clear()
        if ($announcements.Links) {
            foreach ($link in $announcements.Links) {
                Add-AnnouncementLinkControl -Name $link.Name -Url $link.Url
            }
        }
        
    } catch {
        $StatusText.Text = "Error populating announcements editor: $($_.Exception.Message)"
        $StatusText.Foreground = [System.Windows.Media.Brushes]::Red
    }
}

# Function: Add a link control to the panel
function Add-AnnouncementLinkControl {
    param([string]$Name = "", [string]$Url = "")
    
    $grid = New-Object System.Windows.Controls.Grid
    $grid.Margin = [System.Windows.Thickness]::new(0,2,0,2)
    
    $col1 = New-Object System.Windows.Controls.ColumnDefinition
    $col1.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
    $col2 = New-Object System.Windows.Controls.ColumnDefinition
    $col2.Width = [System.Windows.GridLength]::new(2, [System.Windows.GridUnitType]::Star)
    $col3 = New-Object System.Windows.Controls.ColumnDefinition
    $col3.Width = [System.Windows.GridLength]::Auto
    
    $grid.ColumnDefinitions.Add($col1)
    $grid.ColumnDefinitions.Add($col2)
    $grid.ColumnDefinitions.Add($col3)
    
    $nameBox = New-Object System.Windows.Controls.TextBox
    $nameBox.Text = $Name
    $nameBox.Margin = [System.Windows.Thickness]::new(0,0,5,0)
    $nameBox.Tag = "LinkName"
    [System.Windows.Controls.Grid]::SetColumn($nameBox, 0)
    
    $urlBox = New-Object System.Windows.Controls.TextBox
    $urlBox.Text = $Url
    $urlBox.Margin = [System.Windows.Thickness]::new(0,0,5,0)
    $urlBox.Tag = "LinkUrl"
    [System.Windows.Controls.Grid]::SetColumn($urlBox, 1)
    
    $removeBtn = New-Object System.Windows.Controls.Button
$removeBtn.Content = "Remove"
$removeBtn.Padding = [System.Windows.Thickness]::new(5,2,5,2)
$removeBtn.Tag = $grid  # Store reference to parent grid
$removeBtn.Add_Click({
    $gridToRemove = $this.Tag
    $AnnouncementLinksPanel.Children.Remove($gridToRemove)
})
[System.Windows.Controls.Grid]::SetColumn($removeBtn, 2)
    
    $grid.Children.Add($nameBox)
    $grid.Children.Add($urlBox)
    $grid.Children.Add($removeBtn)
    
    $AnnouncementLinksPanel.Children.Add($grid)
}

# Function: Save Announcements changes back to JSON object
function Save-AnnouncementsChanges {
    try {
        # Update text and details
        $global:ContentData.Dashboard.Announcements.Default.Text = $AnnouncementText.Text
        $global:ContentData.Dashboard.Announcements.Default.Details = $AnnouncementDetails.Text
        
        # Update links
        $links = @()
        foreach ($grid in $AnnouncementLinksPanel.Children) {
            $nameBox = $grid.Children | Where-Object { $_.Tag -eq "LinkName" }
            $urlBox = $grid.Children | Where-Object { $_.Tag -eq "LinkUrl" }
            
            if ($nameBox.Text -and $urlBox.Text) {
                $links += @{
                    Name = $nameBox.Text
                    Url = $urlBox.Text
                }
            }
        }
        $global:ContentData.Dashboard.Announcements.Default.Links = $links
        
        # Update raw JSON display
        $RawJsonBox.Text = $global:ContentData | ConvertTo-Json -Depth 100
        
        $StatusText.Text = "Announcements saved successfully"
        $StatusText.Foreground = [System.Windows.Media.Brushes]::Green
        
    } catch {
        $StatusText.Text = "Error saving announcements: $($_.Exception.Message)"
        $StatusText.Foreground = [System.Windows.Media.Brushes]::Red
    }
}

# Function: Populate Support editor from loaded JSON
function Populate-SupportEditor {
    try {
        $support = $global:ContentData.Dashboard.Support
        
        $SupportText.Text = $support.Text
        
        # Clear and populate links
        $SupportLinksPanel.Children.Clear()
        if ($support.Links) {
            foreach ($link in $support.Links) {
                Add-SupportLinkControl -Name $link.Name -Url $link.Url
            }
        }
        
    } catch {
        $StatusText.Text = "Error populating support editor: $($_.Exception.Message)"
        $StatusText.Foreground = [System.Windows.Media.Brushes]::Red
    }
}

# Function: Add a link control to the support panel
function Add-SupportLinkControl {
    param([string]$Name = "", [string]$Url = "")
    
    $grid = New-Object System.Windows.Controls.Grid
    $grid.Margin = [System.Windows.Thickness]::new(0,2,0,2)
    
    $col1 = New-Object System.Windows.Controls.ColumnDefinition
    $col1.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
    $col2 = New-Object System.Windows.Controls.ColumnDefinition
    $col2.Width = [System.Windows.GridLength]::new(2, [System.Windows.GridUnitType]::Star)
    $col3 = New-Object System.Windows.Controls.ColumnDefinition
    $col3.Width = [System.Windows.GridLength]::Auto
    
    $grid.ColumnDefinitions.Add($col1)
    $grid.ColumnDefinitions.Add($col2)
    $grid.ColumnDefinitions.Add($col3)
    
    $nameBox = New-Object System.Windows.Controls.TextBox
    $nameBox.Text = $Name
    $nameBox.Margin = [System.Windows.Thickness]::new(0,0,5,0)
    $nameBox.Tag = "LinkName"
    [System.Windows.Controls.Grid]::SetColumn($nameBox, 0)
    
    $urlBox = New-Object System.Windows.Controls.TextBox
    $urlBox.Text = $Url
    $urlBox.Margin = [System.Windows.Thickness]::new(0,0,5,0)
    $urlBox.Tag = "LinkUrl"
    [System.Windows.Controls.Grid]::SetColumn($urlBox, 1)
    
    $removeBtn = New-Object System.Windows.Controls.Button
    $removeBtn.Content = "Remove"
    $removeBtn.Padding = [System.Windows.Thickness]::new(5,2,5,2)
    $removeBtn.Tag = $grid
    $removeBtn.Add_Click({
        $gridToRemove = $this.Tag
        $SupportLinksPanel.Children.Remove($gridToRemove)
    })
    [System.Windows.Controls.Grid]::SetColumn($removeBtn, 2)
    
    $grid.Children.Add($nameBox)
    $grid.Children.Add($urlBox)
    $grid.Children.Add($removeBtn)
    
    $SupportLinksPanel.Children.Add($grid)
}

# Function: Save Support changes back to JSON object
function Save-SupportChanges {
    try {
        # Update text
        $global:ContentData.Dashboard.Support.Text = $SupportText.Text
        
        # Update links
        $links = @()
        foreach ($grid in $SupportLinksPanel.Children) {
            $nameBox = $grid.Children | Where-Object { $_.Tag -eq "LinkName" }
            $urlBox = $grid.Children | Where-Object { $_.Tag -eq "LinkUrl" }
            
            if ($nameBox.Text -and $urlBox.Text) {
                $links += @{
                    Name = $nameBox.Text
                    Url = $urlBox.Text
                }
            }
        }
        $global:ContentData.Dashboard.Support.Links = $links
        
        # Update raw JSON display
        $RawJsonBox.Text = $global:ContentData | ConvertTo-Json -Depth 100
        
        $StatusText.Text = "Support information saved successfully"
        $StatusText.Foreground = [System.Windows.Media.Brushes]::Green
        
    } catch {
        $StatusText.Text = "Error saving support: $($_.Exception.Message)"
        $StatusText.Foreground = [System.Windows.Media.Brushes]::Red
    }
}

# Global to track current editing tab
$global:CurrentEditingTabIndex = -1

# Function: Populate Additional Tabs list
function Populate-AdditionalTabsList {
    try {
        $TabsList.Items.Clear()
        
        if ($global:ContentData.AdditionalTabs) {
            for ($i = 0; $i -lt $global:ContentData.AdditionalTabs.Count; $i++) {
                $tab = $global:ContentData.AdditionalTabs[$i]
                $enabledText = if ($tab.Enabled) { "" } else { " (Disabled)" }
                $TabsList.Items.Add("$($tab.TabHeader)$enabledText")
            }
        }
        
    } catch {
        $StatusText.Text = "Error populating tabs list: $($_.Exception.Message)"
        $StatusText.Foreground = [System.Windows.Media.Brushes]::Red
    }
}

# Function: Load selected tab into editor
function Load-TabIntoEditor {
    param([int]$Index)
    
    if ($Index -lt 0 -or $Index -ge $global:ContentData.AdditionalTabs.Count) {
        $TabEditorPanel.IsEnabled = $false
        return
    }
    
    $global:CurrentEditingTabIndex = $Index
    $tab = $global:ContentData.AdditionalTabs[$Index]
    
    $TabHeaderBox.Text = $tab.TabHeader
    $TabEnabledCheck.IsChecked = $tab.Enabled
    $TabContentTextBox.Text = $tab.Content.Text
    
    # Load content links
    $TabContentLinksPanel.Children.Clear()
    if ($tab.Content.Links) {
        foreach ($link in $tab.Content.Links) {
            Add-TabContentLinkControl -Name $link.Name -Url $link.Url
        }
    }
    
    $TabEditorPanel.IsEnabled = $true
	Update-AdditionalTabsPreview
}

# Function: Add tab content link control
function Add-TabContentLinkControl {
    param([string]$Name = "", [string]$Url = "")
    
    $grid = New-Object System.Windows.Controls.Grid
    $grid.Margin = [System.Windows.Thickness]::new(0,2,0,2)
    
    $col1 = New-Object System.Windows.Controls.ColumnDefinition
    $col1.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
    $col2 = New-Object System.Windows.Controls.ColumnDefinition
    $col2.Width = [System.Windows.GridLength]::new(2, [System.Windows.GridUnitType]::Star)
    $col3 = New-Object System.Windows.Controls.ColumnDefinition
    $col3.Width = [System.Windows.GridLength]::Auto
    
    $grid.ColumnDefinitions.Add($col1)
    $grid.ColumnDefinitions.Add($col2)
    $grid.ColumnDefinitions.Add($col3)
    
    $nameBox = New-Object System.Windows.Controls.TextBox
    $nameBox.Text = $Name
    $nameBox.Margin = [System.Windows.Thickness]::new(0,0,5,0)
    $nameBox.Tag = "LinkName"
    [System.Windows.Controls.Grid]::SetColumn($nameBox, 0)
    
    $urlBox = New-Object System.Windows.Controls.TextBox
    $urlBox.Text = $Url
    $urlBox.Margin = [System.Windows.Thickness]::new(0,0,5,0)
    $urlBox.Tag = "LinkUrl"
    [System.Windows.Controls.Grid]::SetColumn($urlBox, 1)
    
    $removeBtn = New-Object System.Windows.Controls.Button
    $removeBtn.Content = "Remove"
    $removeBtn.Padding = [System.Windows.Thickness]::new(5,2,5,2)
    $removeBtn.Tag = $grid
    $removeBtn.Add_Click({
        $gridToRemove = $this.Tag
        $TabContentLinksPanel.Children.Remove($gridToRemove)
    })
    [System.Windows.Controls.Grid]::SetColumn($removeBtn, 2)
    
    $grid.Children.Add($nameBox)
    $grid.Children.Add($urlBox)
    $grid.Children.Add($removeBtn)
    
    $TabContentLinksPanel.Children.Add($grid)
}

# Function: Save current tab changes
function Save-CurrentTabChanges {
    if ($global:CurrentEditingTabIndex -lt 0) { return }
    
    try {
        $tab = $global:ContentData.AdditionalTabs[$global:CurrentEditingTabIndex]
        
        $tab.TabHeader = $TabHeaderBox.Text
        $tab.Enabled = $TabEnabledCheck.IsChecked
        $tab.Content.Text = $TabContentTextBox.Text
        
        # Save links
        $links = @()
        foreach ($grid in $TabContentLinksPanel.Children) {
            $nameBox = $grid.Children | Where-Object { $_.Tag -eq "LinkName" }
            $urlBox = $grid.Children | Where-Object { $_.Tag -eq "LinkUrl" }
            
            if ($nameBox.Text -and $urlBox.Text) {
                $links += @{
                    Name = $nameBox.Text
                    Url = $urlBox.Text
                }
            }
        }
        $tab.Content.Links = $links
        
        # Refresh list and raw JSON
        Populate-AdditionalTabsList
        $TabsList.SelectedIndex = $global:CurrentEditingTabIndex
        $RawJsonBox.Text = $global:ContentData | ConvertTo-Json -Depth 100
        
        $StatusText.Text = "Tab '$($tab.TabHeader)' saved successfully"
        $StatusText.Foreground = [System.Windows.Media.Brushes]::Green
        
    } catch {
        $StatusText.Text = "Error saving tab: $($_.Exception.Message)"
        $StatusText.Foreground = [System.Windows.Media.Brushes]::Red
    }
}

# Function: Add new tab
function Add-NewTab {
    try {
        if (-not $global:ContentData.AdditionalTabs) {
            $global:ContentData | Add-Member -NotePropertyName "AdditionalTabs" -NotePropertyValue @() -Force
        }
        
        $newTab = @{
            TabHeader = "New Tab"
            Enabled = $true
            Content = @{
                Text = "Enter tab content here..."
                Links = @()
            }
        }
        
        $global:ContentData.AdditionalTabs += $newTab
        Populate-AdditionalTabsList
        $TabsList.SelectedIndex = $global:ContentData.AdditionalTabs.Count - 1
        
    } catch {
        $StatusText.Text = "Error adding new tab: $($_.Exception.Message)"
        $StatusText.Foreground = [System.Windows.Media.Brushes]::Red
    }
}

# Function: Remove selected tab
function Remove-SelectedTab {
    $index = $TabsList.SelectedIndex
    if ($index -lt 0) { return }
    
    try {
        $result = [System.Windows.MessageBox]::Show(
            "Are you sure you want to remove this tab?",
            "Confirm Removal",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning
        )
        
        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            $tempList = New-Object System.Collections.ArrayList
            $tempList.AddRange($global:ContentData.AdditionalTabs)
            $tempList.RemoveAt($index)
            $global:ContentData.AdditionalTabs = $tempList.ToArray()
            
            Populate-AdditionalTabsList
            $TabEditorPanel.IsEnabled = $false
            $RawJsonBox.Text = $global:ContentData | ConvertTo-Json -Depth 100
            
            $StatusText.Text = "Tab removed successfully"
            $StatusText.Foreground = [System.Windows.Media.Brushes]::Green
        }
        
    } catch {
        $StatusText.Text = "Error removing tab: $($_.Exception.Message)"
        $StatusText.Foreground = [System.Windows.Media.Brushes]::Red
    }
}

# Global to track current editing targeted announcement
$global:CurrentEditingTargetedIndex = -1

# Function: Populate targeted announcements list
function Populate-TargetedAnnouncementsList {
    try {
        $TargetedAnnouncementsList.Items.Clear()
        
        $targeted = $global:ContentData.Dashboard.Announcements.Targeted
        if ($targeted) {
            for ($i = 0; $i -lt $targeted.Count; $i++) {
                $item = $targeted[$i]
                $enabledText = if ($item.Enabled) { "" } else { " (Disabled)" }
                $appendText = if ($item.AppendToDefault) { " [Append]" } else { " [Replace]" }
                $displayName = if ($item.Text) { 
                    $preview = $item.Text.Substring(0, [Math]::Min(30, $item.Text.Length))
                    "$preview...$enabledText$appendText"
                } else {
                    "Targeted $($i + 1)$enabledText$appendText"
                }
                $TargetedAnnouncementsList.Items.Add($displayName)
            }
        }
        
    } catch {
        $StatusText.Text = "Error populating targeted list: $($_.Exception.Message)"
        $StatusText.Foreground = [System.Windows.Media.Brushes]::Red
    }
}

# Function: Load targeted announcement into editor
function Load-TargetedIntoEditor {
    param([int]$Index)
    
    $targeted = $global:ContentData.Dashboard.Announcements.Targeted
    if ($Index -lt 0 -or $Index -ge $targeted.Count) {
        $TargetedEditorPanel.IsEnabled = $false
        return
    }
    
    $global:CurrentEditingTargetedIndex = $Index
    $item = $targeted[$Index]
    
    $TargetedEnabledCheck.IsChecked = $item.Enabled
    $TargetedAppendCheck.IsChecked = $item.AppendToDefault
    $TargetedConditionPathBox.Text = $item.Condition.Path
    $TargetedConditionNameBox.Text = $item.Condition.Name
    $TargetedConditionValueBox.Text = $item.Condition.Value
    $TargetedTextBox.Text = $item.Text
    $TargetedDetailsBox.Text = $item.Details
    
    # Load links
    $TargetedLinksPanel.Children.Clear()
    if ($item.Links) {
        foreach ($link in $item.Links) {
            Add-TargetedLinkControl -Name $link.Name -Url $link.Url
        }
    }
    
    $TargetedEditorPanel.IsEnabled = $true
}

# Function: Add targeted link control
function Add-TargetedLinkControl {
    param([string]$Name = "", [string]$Url = "")
    
    $grid = New-Object System.Windows.Controls.Grid
    $grid.Margin = [System.Windows.Thickness]::new(0,2,0,2)
    
    $col1 = New-Object System.Windows.Controls.ColumnDefinition
    $col1.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
    $col2 = New-Object System.Windows.Controls.ColumnDefinition
    $col2.Width = [System.Windows.GridLength]::new(2, [System.Windows.GridUnitType]::Star)
    $col3 = New-Object System.Windows.Controls.ColumnDefinition
    $col3.Width = [System.Windows.GridLength]::Auto
    
    $grid.ColumnDefinitions.Add($col1)
    $grid.ColumnDefinitions.Add($col2)
    $grid.ColumnDefinitions.Add($col3)
    
    $nameBox = New-Object System.Windows.Controls.TextBox
    $nameBox.Text = $Name
    $nameBox.Margin = [System.Windows.Thickness]::new(0,0,5,0)
    $nameBox.Tag = "LinkName"
    [System.Windows.Controls.Grid]::SetColumn($nameBox, 0)
    
    $urlBox = New-Object System.Windows.Controls.TextBox
    $urlBox.Text = $Url
    $urlBox.Margin = [System.Windows.Thickness]::new(0,0,5,0)
    $urlBox.Tag = "LinkUrl"
    [System.Windows.Controls.Grid]::SetColumn($urlBox, 1)
    
    $removeBtn = New-Object System.Windows.Controls.Button
    $removeBtn.Content = "Remove"
    $removeBtn.Padding = [System.Windows.Thickness]::new(5,2,5,2)
    $removeBtn.Tag = $grid
    $removeBtn.Add_Click({
        $gridToRemove = $this.Tag
        $TargetedLinksPanel.Children.Remove($gridToRemove)
    })
    [System.Windows.Controls.Grid]::SetColumn($removeBtn, 2)
    
    $grid.Children.Add($nameBox)
    $grid.Children.Add($urlBox)
    $grid.Children.Add($removeBtn)
    
    $TargetedLinksPanel.Children.Add($grid)
}

# Function: Save targeted announcement changes
function Save-TargetedChanges {
    if ($global:CurrentEditingTargetedIndex -lt 0) { return }
    
    try {
        $targeted = $global:ContentData.Dashboard.Announcements.Targeted[$global:CurrentEditingTargetedIndex]
        
        $targeted.Enabled = $TargetedEnabledCheck.IsChecked
        $targeted.AppendToDefault = $TargetedAppendCheck.IsChecked
        $targeted.Condition.Path = $TargetedConditionPathBox.Text
        $targeted.Condition.Name = $TargetedConditionNameBox.Text
        $targeted.Condition.Value = $TargetedConditionValueBox.Text
        $targeted.Text = $TargetedTextBox.Text
        $targeted.Details = $TargetedDetailsBox.Text
        
        # Save links
        $links = @()
        foreach ($grid in $TargetedLinksPanel.Children) {
            $nameBox = $grid.Children | Where-Object { $_.Tag -eq "LinkName" }
            $urlBox = $grid.Children | Where-Object { $_.Tag -eq "LinkUrl" }
            
            if ($nameBox.Text -and $urlBox.Text) {
                $links += @{
                    Name = $nameBox.Text
                    Url = $urlBox.Text
                }
            }
        }
        $targeted.Links = $links
        
        # Refresh list and raw JSON
        Populate-TargetedAnnouncementsList
        $TargetedAnnouncementsList.SelectedIndex = $global:CurrentEditingTargetedIndex
        $RawJsonBox.Text = $global:ContentData | ConvertTo-Json -Depth 100
        
        $StatusText.Text = "Targeted announcement saved successfully"
        $StatusText.Foreground = [System.Windows.Media.Brushes]::Green
        
    } catch {
        $StatusText.Text = "Error saving targeted announcement: $($_.Exception.Message)"
        $StatusText.Foreground = [System.Windows.Media.Brushes]::Red
    }
}

# Function: Add new targeted announcement
function Add-NewTargeted {
    try {
        $newTargeted = @{
            Enabled = $true
            AppendToDefault = $true
            Condition = @{
                Type = "Registry"
                Path = "HKLM:\SOFTWARE\YourKey"
                Name = "ValueName"
                Value = "ExpectedValue"
            }
            Text = "Targeted announcement text"
            Details = ""
            Links = @()
        }
        
        if (-not $global:ContentData.Dashboard.Announcements.Targeted) {
            $global:ContentData.Dashboard.Announcements | Add-Member -NotePropertyName "Targeted" -NotePropertyValue @() -Force
        }
        
        $tempList = New-Object System.Collections.ArrayList
        $tempList.AddRange($global:ContentData.Dashboard.Announcements.Targeted)
        $tempList.Add($newTargeted) | Out-Null
        $global:ContentData.Dashboard.Announcements.Targeted = $tempList.ToArray()
        
        Populate-TargetedAnnouncementsList
        $TargetedAnnouncementsList.SelectedIndex = $global:ContentData.Dashboard.Announcements.Targeted.Count - 1
        
    } catch {
        $StatusText.Text = "Error adding targeted announcement: $($_.Exception.Message)"
        $StatusText.Foreground = [System.Windows.Media.Brushes]::Red
    }
}

# Function: Remove targeted announcement
function Remove-SelectedTargeted {
    $index = $TargetedAnnouncementsList.SelectedIndex
    if ($index -lt 0) { return }
    
    try {
        $result = [System.Windows.MessageBox]::Show(
            "Are you sure you want to remove this targeted announcement?",
            "Confirm Removal",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning
        )
        
        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            $tempList = New-Object System.Collections.ArrayList
            $tempList.AddRange($global:ContentData.Dashboard.Announcements.Targeted)
            $tempList.RemoveAt($index)
            $global:ContentData.Dashboard.Announcements.Targeted = $tempList.ToArray()
            
            Populate-TargetedAnnouncementsList
            $TargetedEditorPanel.IsEnabled = $false
            $RawJsonBox.Text = $global:ContentData | ConvertTo-Json -Depth 100
            
            $StatusText.Text = "Targeted announcement removed successfully"
            $StatusText.Foreground = [System.Windows.Media.Brushes]::Green
        }
        
    } catch {
        $StatusText.Text = "Error removing targeted announcement: $($_.Exception.Message)"
        $StatusText.Foreground = [System.Windows.Media.Brushes]::Red
    }
}

# Configuration file path
$ConfigPath = Join-Path $PSScriptRoot "ContentEditorConfig.json"

# Function: Load GitHub configuration
function Load-GitHubConfig {
    try {
        if (Test-Path $ConfigPath) {
            $config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
            $GitHubOwnerBox.Text = $config.Owner
            $GitHubRepoBox.Text = $config.Repo
            $GitHubBranchBox.Text = $config.Branch
            $GitHubFilePathBox.Text = $config.FilePath
            
            # Token should be stored securely, but for simplicity we'll store in config
            # In production, use Windows Credential Manager
            if ($config.Token) {
                $GitHubTokenBox.Text = $config.Token
                $CommitToGitHubButton.IsEnabled = $true
            }
        }
    } catch {
        # Ignore errors on first run
    }
}

# Function: Save GitHub configuration
function Save-GitHubConfig {
    try {
        $config = @{
            Owner = $GitHubOwnerBox.Text
            Repo = $GitHubRepoBox.Text
            Branch = $GitHubBranchBox.Text
            FilePath = $GitHubFilePathBox.Text
            Token = $GitHubTokenBox.Text
        }
        
        $config | ConvertTo-Json | Out-File $ConfigPath -Force
        
        if ($GitHubTokenBox.Text) {
            $CommitToGitHubButton.IsEnabled = $true
        }
        
        $StatusText.Text = "GitHub configuration saved"
        $StatusText.Foreground = [System.Windows.Media.Brushes]::Green
        
    } catch {
        $StatusText.Text = "Error saving configuration: $($_.Exception.Message)"
        $StatusText.Foreground = [System.Windows.Media.Brushes]::Red
    }
}

# Function: Test GitHub connection
function Test-GitHubConnection {
    try {
        $CommitStatusText.Text = "Testing connection..."
        
        $owner = $GitHubOwnerBox.Text
        $repo = $GitHubRepoBox.Text
        $token = $GitHubTokenBox.Text
        
        $headers = @{
            "Authorization" = "token $token"
            "Accept" = "application/vnd.github.v3+json"
        }
        
        $url = "https://api.github.com/repos/$owner/$repo"
        $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get
        
        $CommitStatusText.Text = "[OK] Connected to: $($response.full_name)"
        $CommitStatusText.Foreground = [System.Windows.Media.Brushes]::Green
        $StatusText.Text = "GitHub connection successful"
        $StatusText.Foreground = [System.Windows.Media.Brushes]::Green
        
    } catch {
        $CommitStatusText.Text = "[ERROR] Connection failed"
        $CommitStatusText.Foreground = [System.Windows.Media.Brushes]::Red
        $StatusText.Text = "Error: $($_.Exception.Message)"
        $StatusText.Foreground = [System.Windows.Media.Brushes]::Red
    }
}

# Function: Commit to GitHub
function Commit-ToGitHub {
    try {
        if (-not $CommitMessageBox.Text) {
            [System.Windows.MessageBox]::Show("Please enter a commit message", "Commit Message Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        $CommitStatusText.Text = "Committing to GitHub..."
        $StatusText.Text = "Committing changes..."
        
        $owner = $GitHubOwnerBox.Text
        $repo = $GitHubRepoBox.Text
        $branch = $GitHubBranchBox.Text
        $filePath = $GitHubFilePathBox.Text
        $token = $GitHubTokenBox.Text
        $message = $CommitMessageBox.Text
        
        $headers = @{
            "Authorization" = "token $token"
            "Accept" = "application/vnd.github.v3+json"
        }
        
        # Step 1: Get current file SHA
        $fileUrl = "https://api.github.com/repos/$owner/$repo/contents/$filePath"
        $fileInfo = Invoke-RestMethod -Uri $fileUrl -Headers $headers -Method Get
        $currentSha = $fileInfo.sha
        
        # Step 2: Prepare new content
        $jsonContent = $global:ContentData | ConvertTo-Json -Depth 100
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($jsonContent)
        $base64Content = [Convert]::ToBase64String($bytes)
        
        # Step 3: Commit the change
        $body = @{
            message = $message
            content = $base64Content
            sha = $currentSha
            branch = $branch
        } | ConvertTo-Json
        
        $response = Invoke-RestMethod -Uri $fileUrl -Headers $headers -Method Put -Body $body -ContentType "application/json"
        
        $CommitStatusText.Text = "[OK] Committed successfully"
        $CommitStatusText.Foreground = [System.Windows.Media.Brushes]::Green
        $StatusText.Text = "Changes committed to GitHub successfully!"
        $StatusText.Foreground = [System.Windows.Media.Brushes]::Green
        $CommitMessageBox.Text = ""
        
        [System.Windows.MessageBox]::Show(
            "Changes have been committed to GitHub!`n`nChanges will appear on all endpoints within 15 minutes.",
            "Commit Successful",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        )
        
    } catch {
        $CommitStatusText.Text = "[ERROR] Commit failed"
        $CommitStatusText.Foreground = [System.Windows.Media.Brushes]::Red
        $StatusText.Text = "Error committing: $($_.Exception.Message)"
        $StatusText.Foreground = [System.Windows.Media.Brushes]::Red
        
        [System.Windows.MessageBox]::Show(
            "Failed to commit to GitHub:`n`n$($_.Exception.Message)",
            "Commit Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}

# Function: Save to local file
function Save-ToLocalFile {
    try {
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
        $saveFileDialog.Title = "Save ContentData2.json"
        $saveFileDialog.FileName = "ContentData2.json"
        
        if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $jsonContent = $global:ContentData | ConvertTo-Json -Depth 100
            $jsonContent | Out-File $saveFileDialog.FileName -Force
            
            $StatusText.Text = "Saved to: $($saveFileDialog.FileName)"
            $StatusText.Foreground = [System.Windows.Media.Brushes]::Green
            
            [System.Windows.MessageBox]::Show(
                "File saved successfully!`n`nYou can now manually commit this to Git.",
                "Save Successful",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            )
        }
        
    } catch {
        $StatusText.Text = "Error saving file: $($_.Exception.Message)"
        $StatusText.Foreground = [System.Windows.Media.Brushes]::Red
    }
}

# Function: Refresh Preview
function Refresh-Preview {
    try {
        $PreviewPanel.Children.Clear()
        
        $selectedIndex = $PreviewSectionCombo.SelectedIndex
        
        switch ($selectedIndex) {
            0 { Show-AnnouncementsPreview }
            1 { Show-SupportPreview }
            2 { Show-AdditionalTabsPreview }
        }
        
        $StatusText.Text = "Preview refreshed"
        $StatusText.Foreground = [System.Windows.Media.Brushes]::Green
        
    } catch {
        $StatusText.Text = "Error refreshing preview: $($_.Exception.Message)"
        $StatusText.Foreground = [System.Windows.Media.Brushes]::Red
    }
}

# Function: Show Announcements Preview
function Show-AnnouncementsPreview {
    $announcements = $global:ContentData.Dashboard.Announcements.Default
    
    # Title
    $title = New-Object System.Windows.Controls.TextBlock
    $title.Text = "Announcements"
    $title.FontSize = 14
    $title.FontWeight = "Bold"
    $title.Margin = [System.Windows.Thickness]::new(0,0,0,10)
    $PreviewPanel.Children.Add($title)
    
    # Expander (simulating the collapsible announcement)
    $expander = New-Object System.Windows.Controls.Expander
    $expander.Header = "System Announcements"
    $expander.IsExpanded = $true
    $expander.Margin = [System.Windows.Thickness]::new(0,0,0,10)
    
    $expanderContent = New-Object System.Windows.Controls.StackPanel
    
    # Main text with markdown rendering
    $textBlock = New-Object System.Windows.Controls.TextBlock
    $textBlock.TextWrapping = "Wrap"
    $textBlock.Margin = [System.Windows.Thickness]::new(0,5,0,5)
    
    # Simple markdown rendering
    $renderedText = Render-SimpleMarkdown -Text $announcements.Text
    $textBlock.Inlines.AddRange($renderedText)
    $expanderContent.Children.Add($textBlock)
    
    # Details
    if ($announcements.Details) {
        $detailsBlock = New-Object System.Windows.Controls.TextBlock
        $detailsBlock.Text = $announcements.Details
        $detailsBlock.TextWrapping = "Wrap"
        $detailsBlock.FontSize = 10
        $detailsBlock.Foreground = [System.Windows.Media.Brushes]::Gray
        $detailsBlock.Margin = [System.Windows.Thickness]::new(0,5,0,5)
        $expanderContent.Children.Add($detailsBlock)
    }
    
    # Links
    if ($announcements.Links) {
        $linksPanel = New-Object System.Windows.Controls.StackPanel
        $linksPanel.Margin = [System.Windows.Thickness]::new(0,10,0,0)
        
        foreach ($link in $announcements.Links) {
            $linkButton = New-Object System.Windows.Controls.Button
            $linkButton.Content = "[LINK] $($link.Name)"
            $linkButton.Margin = [System.Windows.Thickness]::new(0,2,0,2)
            $linkButton.Padding = [System.Windows.Thickness]::new(5,2,5,2)
            $linkButton.HorizontalAlignment = "Left"
            $linkButton.Cursor = [System.Windows.Input.Cursors]::Hand
            $linkButton.ToolTip = $link.Url
            $linksPanel.Children.Add($linkButton)
        }
        
        $expanderContent.Children.Add($linksPanel)
    }
    
    $expander.Content = $expanderContent
    $PreviewPanel.Children.Add($expander)
    
    # Note about targeted announcements
    $targetedNote = New-Object System.Windows.Controls.TextBlock
    $targetedNote.Text = "Note: Targeted announcements are conditional and will only appear when registry conditions are met."
    $targetedNote.FontSize = 9
    $targetedNote.FontStyle = "Italic"
    $targetedNote.Foreground = [System.Windows.Media.Brushes]::Gray
    $targetedNote.Margin = [System.Windows.Thickness]::new(0,10,0,0)
    $PreviewPanel.Children.Add($targetedNote)
}

# Function: Show Support Preview
function Show-SupportPreview {
    $support = $global:ContentData.Dashboard.Support
    
    # Title
    $title = New-Object System.Windows.Controls.TextBlock
    $title.Text = "Support"
    $title.FontSize = 14
    $title.FontWeight = "Bold"
    $title.Margin = [System.Windows.Thickness]::new(0,0,0,10)
    $PreviewPanel.Children.Add($title)
    
    # Border around content
    $border = New-Object System.Windows.Controls.Border
    $border.BorderBrush = [System.Windows.Media.Brushes]::DarkBlue
    $border.BorderThickness = [System.Windows.Thickness]::new(2)
    $border.Padding = [System.Windows.Thickness]::new(8)
    $border.CornerRadius = 3
    
    $contentPanel = New-Object System.Windows.Controls.StackPanel
    
    # Text with markdown
    $textBlock = New-Object System.Windows.Controls.TextBlock
    $textBlock.TextWrapping = "Wrap"
    $textBlock.Margin = [System.Windows.Thickness]::new(0,0,0,10)
    $renderedText = Render-SimpleMarkdown -Text $support.Text
    $textBlock.Inlines.AddRange($renderedText)
    $contentPanel.Children.Add($textBlock)
    
    # Links
    if ($support.Links) {
        foreach ($link in $support.Links) {
            $linkButton = New-Object System.Windows.Controls.Button
            $linkButton.Content = "[LINK] $($link.Name)"
            $linkButton.Margin = [System.Windows.Thickness]::new(0,2,0,2)
            $linkButton.Padding = [System.Windows.Thickness]::new(5,2,5,2)
            $linkButton.HorizontalAlignment = "Left"
            $linkButton.Cursor = [System.Windows.Input.Cursors]::Hand
            $linkButton.ToolTip = $link.Url
            $contentPanel.Children.Add($linkButton)
        }
    }
    
    $border.Child = $contentPanel
    $PreviewPanel.Children.Add($border)
}

# Function: Show Additional Tabs Preview
function Show-AdditionalTabsPreview {
    if (-not $global:ContentData.AdditionalTabs -or $global:ContentData.AdditionalTabs.Count -eq 0) {
        $noTabsText = New-Object System.Windows.Controls.TextBlock
        $noTabsText.Text = "No additional tabs defined"
        $noTabsText.FontStyle = "Italic"
        $noTabsText.Foreground = [System.Windows.Media.Brushes]::Gray
        $PreviewPanel.Children.Add($noTabsText)
        return
    }
    
    foreach ($tab in $global:ContentData.AdditionalTabs) {
        if (-not $tab.Enabled) { continue }
        
        # Tab header
        $header = New-Object System.Windows.Controls.TextBlock
        $header.Text = $tab.TabHeader
        $header.FontSize = 14
        $header.FontWeight = "Bold"
        $header.Margin = [System.Windows.Thickness]::new(0,10,0,5)
        $PreviewPanel.Children.Add($header)
        
        $separator = New-Object System.Windows.Controls.Separator
        $separator.Margin = [System.Windows.Thickness]::new(0,0,0,10)
        $PreviewPanel.Children.Add($separator)
        
        # Content text
        if ($tab.Content.Text) {
            $textBlock = New-Object System.Windows.Controls.TextBlock
            $textBlock.Text = $tab.Content.Text
            $textBlock.TextWrapping = "Wrap"
            $textBlock.Margin = [System.Windows.Thickness]::new(10,0,0,10)
            $PreviewPanel.Children.Add($textBlock)
        }
        
        # Content links
        if ($tab.Content.Links) {
            foreach ($link in $tab.Content.Links) {
                $linkButton = New-Object System.Windows.Controls.Button
                $linkButton.Content = "[LINK] $($link.Name)"
                $linkButton.Margin = [System.Windows.Thickness]::new(10,2,0,2)
                $linkButton.Padding = [System.Windows.Thickness]::new(5,2,5,2)
                $linkButton.HorizontalAlignment = "Left"
                $linkButton.Cursor = [System.Windows.Input.Cursors]::Hand
                $linkButton.ToolTip = $link.Url
                $PreviewPanel.Children.Add($linkButton)
            }
        }
    }
}

# Load GitHub configuration on startup
Load-GitHubConfig

# Wire up events
$LoadButton.Add_Click({ Load-FromGitHub })
$LoadFileButton.Add_Click({ Load-FromFile })
$AddAnnouncementLinkButton.Add_Click({ Add-AnnouncementLinkControl })
$SaveAnnouncementsButton.Add_Click({ Save-AnnouncementsChanges })
$RevertAnnouncementsButton.Add_Click({ Populate-AnnouncementsEditor })

$AddSupportLinkButton.Add_Click({ Add-SupportLinkControl })
$SaveSupportButton.Add_Click({ Save-SupportChanges })
$RevertSupportButton.Add_Click({ Populate-SupportEditor })

$TabsList.Add_SelectionChanged({
    if ($TabsList.SelectedIndex -ge 0) {
        Load-TabIntoEditor -Index $TabsList.SelectedIndex
    }
})
$AddTabButton.Add_Click({ Add-NewTab })
$RemoveTabButton.Add_Click({ Remove-SelectedTab })
$AddTabContentLinkButton.Add_Click({ Add-TabContentLinkControl })
$SaveTabButton.Add_Click({ Save-CurrentTabChanges })
$RevertTabButton.Add_Click({
    if ($global:CurrentEditingTabIndex -ge 0) {
        Load-TabIntoEditor -Index $global:CurrentEditingTabIndex
    }
})

$TargetedAnnouncementsList.Add_SelectionChanged({
    if ($TargetedAnnouncementsList.SelectedIndex -ge 0) {
        Load-TargetedIntoEditor -Index $TargetedAnnouncementsList.SelectedIndex
    }
})
$AddTargetedButton.Add_Click({ Add-NewTargeted })
$RemoveTargetedButton.Add_Click({ Remove-SelectedTargeted })
$AddTargetedLinkButton.Add_Click({ Add-TargetedLinkControl })
$SaveTargetedButton.Add_Click({ Save-TargetedChanges })

$SaveGitHubConfigButton.Add_Click({ Save-GitHubConfig })
$TestGitHubConnectionButton.Add_Click({ Test-GitHubConnection })

$CommitToGitHubButton.Add_Click({ Commit-ToGitHub })
$SaveToFileButton.Add_Click({ Save-ToLocalFile })

# Support preview updates
$SupportText.Add_TextChanged({ Update-SupportPreview })

# Additional Tabs preview updates
$TabHeaderBox.Add_TextChanged({ Update-AdditionalTabsPreview })
$TabContentTextBox.Add_TextChanged({ Update-AdditionalTabsPreview })
$TabEnabledCheck.Add_Checked({ Update-AdditionalTabsPreview })
$TabEnabledCheck.Add_Unchecked({ Update-AdditionalTabsPreview })

# Show window
$window.ShowDialog() | Out-Null
