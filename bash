Lincoln Laboratory Endpoint Advisor - Web Editor
A powerful, user-friendly web-based editor for managing JSON configuration files that drive the Lincoln Laboratory Endpoint Advisor PowerShell application. Edit announcements, support information, and targeted messages with a modern interface - no JSON knowledge required!
🌟 Features
Dual Editing Modes

📝 Form Editor (Easy Mode) - Visual interface with text fields, checkboxes, and buttons
💻 JSON Editor (Advanced Mode) - Direct JSON editing with syntax validation
Seamlessly switch between modes - changes sync automatically

Live Preview

Real-time preview showing exactly how content will appear on endpoints
Markdown rendering (bold, italic, colors, links)
Collapsible sections matching the actual application UI

Smart Content Management

Auto-Save - Automatic browser-based draft saving every 2 seconds
Diff Viewer - Side-by-side comparison before publishing to GitHub
Templates - 6 pre-built templates for common scenarios:

System Maintenance
Security Alerts
Patching Notices
Training Reminders
Holiday Schedules
Pilot Programs



Targeted Announcements

Create registry-based conditional messages
Show different content to different user groups
Built-in BigFix Action Script templates for easy deployment
One-click copy for all scripts

GitHub Integration

Direct editing of files in GitHub repositories (public or Enterprise)
Support for multiple GitHub instances via dropdown selector
Token permission checker to diagnose authentication issues
Handles both classic and Enterprise GitHub with SSO

Rich Text Support

Markdown toolbar for easy formatting
Color tags: [green], [red], [yellow], [blue]
Bold, italic, underline support
Visual preview while editing

Built-in Help System

Comprehensive documentation
5 tabbed help sections
Copy-paste ready examples
Troubleshooting guides

🚀 Quick Start
Prerequisites

A GitHub account with access to your ContentData.json repository
A Personal Access Token with "repo" scope
Modern web browser (Chrome, Firefox, Edge, Safari)

Installation
Option 1: Download and Open

Download endpoint-advisor-editor.html from this repository
Open the file in your web browser
Start editing!

Option 2: Host on Web Server

Place the HTML file on your internal web server
Access via browser at https://your-server/endpoint-advisor-editor.html
Users can bookmark and access easily

First Use

Select GitHub Instance

Choose from dropdown: Public GitHub, Enterprise, or Custom
For Enterprise, select your pre-configured instance


Enter Repository Details

Repository Owner: your-username-or-org
Repository Name: EndpointAdvisor
Branch: main
File Path: ContentData.json


Create Personal Access Token

Go to GitHub Settings → Developer settings → Personal access tokens
Generate new token (classic)
Check "repo" scope (full control)
For Enterprise with SSO: Authorize SSO after creating token
Copy token and paste into editor


Load and Edit

Click "Load File"
Edit using Form Editor or JSON Editor
View live preview on right side
Click "Save to GitHub" when done



🔧 Configuration
Customizing for Your Organization
The editor includes a dropdown for GitHub instances. To customize for your organization:

Open the HTML file in a text editor
Find the <select id="apiBase"> section (around line 560)
Add your organization's GitHub instances:

html<select id="apiBase" onchange="handleApiBaseChange()">
  <option value="https://api.github.com">Public GitHub</option>
  <option value="https://github.ll.mit.edu/api/v3" selected>Lincoln Lab Production</option>
  <option value="https://github-dev.ll.mit.edu/api/v3">Dev Environment</option>
  <option value="custom">Custom URL...</option>
</select>
Tips:

Add selected attribute to your default instance
API URL format: https://your-github-domain/api/v3
Use descriptive names (Production, Dev, Staging)
Keep "Custom URL..." option for flexibility

Default Values
You can pre-fill repository details by editing the HTML:
html<input type="text" id="repoOwner" value="your-org" placeholder="e.g., burnoil">
<input type="text" id="repoName" value="EndpointAdvisor" placeholder="e.g., EndpointAdvisor">
<input type="text" id="branch" value="main" placeholder="e.g., main">
<input type="text" id="filePath" value="ContentData.json" placeholder="e.g., ContentData.json">
📖 Usage Guide
Editing Announcements
Using Form Editor:

Switch to "Form Editor (Easy)" mode
Type announcement text in the text box
Use toolbar buttons for formatting (Bold, Italic, Colors)
Add optional details in the Details field
Add links using "+ Add Link" button
Preview updates automatically on the right

Markdown Formatting:

**bold text** → bold text
*italic text* → italic text
__underlined__ → underlined
[green]success[/green] → green text
[red]error[/red] → red text

Creating Targeted Announcements
Targeted announcements appear only on computers with specific registry keys.
Steps:

Scroll to "Targeted Announcements" section
Click "+ Add Targeted Announcement"
Check "Enabled" to activate
Enter text and details (supports markdown)
Configure registry condition:

Registry Path: HKLM:\SOFTWARE\MITLL\Targeting
Value Name: Group
Expected Value: Pilot


Choose "Append to Default" or replace entirely

Deploying Registry Keys (BigFix):
The editor includes ready-to-use BigFix Action Scripts. Click the Help button → Targeted Announcements tab → BigFix Action Script Templates section.
Example script:
actionscriptregset "[HKEY_LOCAL_MACHINE\SOFTWARE\MITLL\Targeting]" "Group"="Pilot"
```

Deploy this to target computers before the announcement will appear.

### Using Templates

1. Click "📋 Load Template" button
2. Choose from 6 pre-built templates
3. Customize the loaded content
4. Save to GitHub

Templates include:
- System Maintenance announcements
- Security alerts
- Patching notices
- Training reminders
- Holiday schedules
- Pilot program messages

### Saving Changes

1. Click "💾 Save to GitHub" button
2. Review changes in the diff viewer
   - Left: Current GitHub version
   - Right: Your changes
3. Click "Confirm & Save to GitHub"
4. Enter commit message
5. Changes are published!

## 🎯 Common Use Cases

### Monthly Patch Announcements
```
Template: Patching Notice
Text: **Monthly Patch Deployment** - Windows and Office updates available
Details: Updates will be installed overnight. Computers will reboot automatically.
```

### Emergency Security Alert
```
Template: Security Alert
Text: [red]**SECURITY ALERT**[/red] - Action Required
Details: Critical security update available. Install immediately.
```

### Pilot Program for Select Users
```
Targeted Announcement:
- Text: [green]**Pilot Program**[/green] - You've been selected!
- Enabled: ✓
- Append to Default: ✓
- Registry Condition: Group = "Pilot"
🔍 Troubleshooting
"GitHub API error: 401" when saving
Solutions:

Click "🔑 Check Token Permissions" to diagnose
Verify token has "repo" scope checked
For Enterprise GitHub: Authorize SSO

Settings → Personal Access Tokens → Configure SSO → Authorize


Try creating a new token
Check GitHub instance dropdown is correct

Changes not appearing on endpoints
Wait Time: The Endpoint Advisor app refreshes every 10-20 minutes. Wait a bit!
Verify:

File was saved to GitHub (check commit history)
Endpoint can reach GitHub (not blocked by firewall)
PowerShell script URL matches your repository

Targeted announcement not showing
Checklist:

✓ Announcement is marked "Enabled" in editor
✓ Registry key exists on target computer
✓ Registry path, name, and value match exactly (case-sensitive)
✓ BigFix Action was deployed successfully

Verify registry on endpoint:
powershellGet-ItemProperty -Path "HKLM:\SOFTWARE\MITLL\Targeting"
Lost work / browser crashed
Click "💾 Restore Auto-save" button! The editor automatically saves drafts every 2 seconds.
📚 Documentation
Help System
Click the ❓ Help button or floating help button (bottom right) to access:

Getting Started guide
Markdown formatting reference
Targeted announcements setup
Best practices
Troubleshooting tips

JSON Structure
The editor supports both flat and Dashboard structures:
Flat Structure (Legacy):
json{
  "Announcements": {
    "Default": { "Text": "...", "Details": "...", "Links": [] },
    "Targeted": []
  },
  "Support": { "Text": "...", "Links": [] }
}
Dashboard Structure (New):
json{
  "Dashboard": {
    "Announcements": { ... },
    "Support": { ... }
  }
}
🛡️ Security Considerations
GitHub Token Security

Never commit tokens to repositories
Tokens are stored in browser memory only (not saved to disk)
Use tokens with minimum required permissions ("repo" scope)
Rotate tokens regularly
For Enterprise: Enable SSO and MFA

Browser Storage

Auto-save uses localStorage (stays on local machine)
Clear auto-save data after publishing: Done automatically
Use HTTPS when hosting on web servers

🤝 Contributing
Contributions are welcome! Areas for improvement:

Additional templates
More markdown formatting options
Dark mode support
Keyboard shortcuts
Multi-language support

📄 License
This project is part of the Lincoln Laboratory Endpoint Advisor suite.
🙏 Acknowledgments

Built for Lincoln Laboratory IT operations
Integrates with BigFix endpoint management
Designed for PowerShell-based endpoint agents

📞 Support
For questions or issues:

Check the built-in Help documentation
Review troubleshooting section
Contact: endpointengineering@ll.mit.edu


Version: 1.0
Last Updated: 2025-01-15
Compatibility: Chrome 90+, Firefox 88+, Edge 90+, Safari 14+
