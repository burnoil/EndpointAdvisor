<p align="center">
  <h1 align="center">🎯 Lincoln Laboratory Endpoint Advisor - Web Editor</h1>
  <p align="center">
    A powerful, user-friendly web-based editor for managing endpoint announcements
    <br />
    <strong>No JSON knowledge required!</strong>
  </p>
</p>
<p align="center">
  <img alt="GitHub release" src="https://img.shields.io/badge/version-1.0-blue.svg">
  <img alt="License" src="https://img.shields.io/badge/license-MIT-green.svg">
  <img alt="Status" src="https://img.shields.io/badge/status-active-success.svg">
</p>

📋 Table of Contents

Overview
Features
Quick Start
Configuration
Usage Guide
Common Use Cases
Troubleshooting
Security
Support


🎯 Overview
The Endpoint Advisor Web Editor is a single-page HTML application that provides a modern interface for managing JSON configuration files that drive the Lincoln Laboratory Endpoint Advisor PowerShell application.
Why Use This Editor?
BeforeAfter❌ Manual JSON editing✅ Visual form interface❌ Syntax errors✅ Automatic validation❌ No preview✅ Live preview as you type❌ Lost work on crash✅ Auto-save to browser❌ Blind commits✅ Diff viewer before saving

✨ Features
<table>
<tr>
<td width="50%">
📝 Dual Editing Modes

Form Editor for non-technical users
JSON Editor for power users
Instant sync between modes

👁️ Live Preview

Real-time rendering
Markdown support
Exact endpoint simulation

💾 Smart Saving

Auto-save every 2 seconds
Diff viewer before commits
Restore lost work instantly

</td>
<td width="50%">
🎯 Targeted Announcements

Registry-based targeting
BigFix script templates
One-click copy buttons

📋 Templates

6 ready-to-use templates
Maintenance notices
Security alerts
Pilot programs

🔧 GitHub Integration

Public & Enterprise support
Multi-instance dropdown
Token diagnostics

</td>
</tr>
</table>

🚀 Quick Start
Prerequisites
bash✅ GitHub account with repository access
✅ Personal Access Token with "repo" scope
✅ Modern web browser (Chrome, Firefox, Edge, Safari)
Installation
Option 1: Download & Open
bash# 1. Download the HTML file
wget https://raw.githubusercontent.com/yourusername/endpoint-advisor-editor/main/editor.html

# 2. Open in browser
open editor.html
Option 2: Host on Server
bash# Place on your web server
cp editor.html /var/www/html/
# Access at: https://your-server/editor.html
```

### First-Time Setup

<details>
<summary><b>1️⃣ Create GitHub Personal Access Token</b></summary>

1. Go to [GitHub Settings → Tokens](https://github.com/settings/tokens)
2. Click **"Generate new token (classic)"**
3. Check the **`repo`** scope (full control)
4. For Enterprise with SSO: Click **"Authorize SSO"** after creation
5. Copy the token (`ghp_...`)

> **⚠️ Security Note:** Never commit tokens to repositories!

</details>

<details>
<summary><b>2️⃣ Configure Repository Connection</b></summary>

| Field | Example | Description |
|-------|---------|-------------|
| **GitHub Instance** | Internal GitHub Enterprise | Select from dropdown |
| **Repository Owner** | `your-org` | Username or organization |
| **Repository Name** | `EndpointAdvisor` | Repository name |
| **Branch** | `main` | Target branch |
| **File Path** | `ContentData.json` | Path to JSON file |
| **Token** | `ghp_xxxxx...` | Your PAT |

</details>

<details>
<summary><b>3️⃣ Load and Edit</b></summary>
```
1. Click "Load File" → File appears in editor
2. Switch to "Form Editor (Easy)" mode
3. Edit announcements using visual interface
4. View live preview on right side
5. Click "Save to GitHub" → Review diff → Confirm
</details>

🔧 Configuration
Customize GitHub Instance Dropdown
Edit the HTML file to add your organization's GitHub instances:
html<select id="apiBase" onchange="handleApiBaseChange()">
  <option value="https://api.github.com">Public GitHub</option>
  <option value="https://github.ll.mit.edu/api/v3" selected>🏢 Production</option>
  <option value="https://github-dev.ll.mit.edu/api/v3">🧪 Development</option>
  <option value="https://github-stage.ll.mit.edu/api/v3">🚀 Staging</option>
  <option value="custom">✏️ Custom URL...</option>
</select>

💡 Tip: Add selected to your most-used instance. API format: https://your-domain/api/v3

Pre-fill Default Values
html<input type="text" id="repoOwner" value="your-org">
<input type="text" id="repoName" value="EndpointAdvisor">
<input type="text" id="branch" value="main">
<input type="text" id="filePath" value="ContentData.json">
```

---

## 📖 Usage Guide

### Editing Announcements

#### Using Form Editor (Recommended)
```
1. Switch to "📝 Form Editor (Easy)" mode
2. Type announcement text
3. Use toolbar for formatting: [B] [I] [🟢] [🔴] [🟡]
4. Add optional details
5. Add links with "+ Add Link" button
6. Preview updates automatically
```

#### Markdown Formatting

| Syntax | Result |
|--------|--------|
| `**bold**` | **bold** |
| `*italic*` | *italic* |
| `__underline__` | <u>underline</u> |
| `[green]text[/green]` | <span style="color:green">text</span> |
| `[red]warning[/red]` | <span style="color:red">warning</span> |
| `[yellow]caution[/yellow]` | <span style="color:gold">caution</span> |

### Creating Targeted Announcements

> **Targeted announcements** appear only on computers with specific registry keys

<details>
<summary><b>📊 Step-by-Step Setup</b></summary>

#### In the Editor:

1. Scroll to **"🎯 Targeted Announcements"** section
2. Click **"+ Add Targeted Announcement"**
3. Configure:
   - ✅ Check **"Enabled"**
   - ✅ Check **"Append to Default"** (or uncheck to replace)
   - 📝 Enter announcement text
   - 🔗 Add links if needed
   
4. Set **Registry Condition**:
```
   Registry Path:    HKLM:\SOFTWARE\MITLL\Targeting
   Value Name:       Group
   Expected Value:   Pilot
Deploy Registry Keys (BigFix):
Copy from Help → Targeted Announcements → BigFix Scripts:
actionscriptregset "[HKEY_LOCAL_MACHINE\SOFTWARE\MITLL\Targeting]" "Group"="Pilot"
Deploy to target computers using BigFix Actions.
Verify on Endpoint:
powershellGet-ItemProperty -Path "HKLM:\SOFTWARE\MITLL\Targeting"
</details>
Using Templates
mermaidgraph LR
    A[Click Load Template] --> B[Choose Template]
    B --> C[Edit Content]
    C --> D[Preview]
    D --> E[Save to GitHub]
Available Templates:
TemplateUse CaseColors🔧 System MaintenanceScheduled downtimeYellow cautions🔒 Security AlertCritical updatesRed warnings⚙️ Patching NoticeMonthly patchesBlue info📚 Training ReminderRequired trainingGreen highlights🎉 Holiday ScheduleHoliday hoursGreen/Yellow🚀 Pilot ProgramTargeted testingGreen success

🎯 Common Use Cases
Example 1: Monthly Patch Announcement
jsonTemplate: Patching Notice
---
Text: **Monthly Patch Deployment** - Windows and Office updates available

Details: 
**Starting today at 3pm** patches will be deployed.
Updates install overnight. [yellow]**Computers will reboot automatically.**[/yellow]

Daytime patching begins [red]tomorrow[/red] for computers still pending updates.
Example 2: Emergency Security Alert
jsonTemplate: Security Alert
---
Text: [red]**SECURITY ALERT**[/red] - Action Required

Details:
A critical security update has been released.

**Action Required:**
- Update your system immediately
- Restart when prompted
- Contact IT if issues occur
Example 3: Pilot Program (Targeted)
jsonTargeted Announcement
---
Text: [green]**Pilot Program**[/green] - You've been selected!

Details:
You're part of the VPN 2.0 pilot group.

**Next Steps:**
- Install pilot software by Friday
- Review documentation
- Submit feedback survey

Registry Condition:
- Path: HKLM:\SOFTWARE\MITLL\Targeting
- Name: Group
- Value: VPN-Pilot

Append to Default: ✅ Yes

🔍 Troubleshooting
Common Issues
<details>
<summary><b>❌ "GitHub API error: 401" when saving</b></summary>
Possible Causes:

Token missing "repo" scope
Token expired
SSO not authorized (Enterprise)
Wrong GitHub instance selected

Solutions:

Click "🔑 Check Token Permissions" button to diagnose
Verify token has "repo" scope checked
For Enterprise: Settings → Personal Access Tokens → Configure SSO → Authorize
Create fresh token
Verify GitHub instance dropdown selection

</details>
<details>
<summary><b>❌ Changes not appearing on endpoints</b></summary>
Timeline:
The Endpoint Advisor refreshes every 10-20 minutes
Verify:
bash# 1. Check GitHub commit history
git log --oneline -5

# 2. Verify endpoint can reach GitHub
Test-NetConnection github.yourcompany.com -Port 443

# 3. Check PowerShell script URL matches
Get-Content C:\Path\To\LLEA_tabs.ps1 | Select-String "ContentDataUrl"
</details>
<details>
<summary><b>❌ Targeted announcement not showing</b></summary>
Checklist:

 Announcement marked "Enabled" in editor
 Registry key exists on target computer
 Registry path, name, value match exactly (case-sensitive)
 BigFix Action deployed successfully

Verify on endpoint:
powershellGet-ItemProperty -Path "HKLM:\SOFTWARE\MITLL\Targeting" -ErrorAction SilentlyContinue
Debug:
powershell# Check all properties
Get-ItemProperty -Path "HKLM:\SOFTWARE\MITLL\Targeting" | Format-List

# Expected output:
# Group      : Pilot
# Department : Engineering
```

</details>

<details>
<summary><b>💾 Lost work / browser crashed</b></summary>

#### Recovery:
1. Reopen the editor
2. Click **"💾 Restore Auto-save"** button
3. Editor saves drafts every 2 seconds automatically!

> **Note:** Auto-save is browser-specific and computer-specific

</details>

---

## 🛡️ Security

### Best Practices

| ✅ Do | ❌ Don't |
|-------|----------|
| Use tokens with minimum permissions | Share tokens with others |
| Enable SSO and MFA | Store tokens in repositories |
| Rotate tokens regularly | Use admin tokens for routine edits |
| Use HTTPS for hosted editors | Expose editor on public internet |
| Clear auto-save after publishing | Leave sensitive data in browser |

### Token Storage
```
✅ Stored in browser memory only (not saved to disk)
✅ Automatically cleared when browser closes
✅ Never transmitted except to GitHub API
```

### Enterprise Security
```
✅ Supports SSO authorization
✅ Respects branch protection rules
✅ Audit trail via GitHub commits
✅ No external dependencies (single HTML file)
```

---

## 📚 Additional Resources

### Built-in Help System

The editor includes comprehensive documentation:
```
❓ Help Button → 5 Tabbed Sections:
  📖 Getting Started
  📝 Markdown Formatting  
  🎯 Targeted Announcements
  ✨ Best Practices
  🔍 Troubleshooting
JSON Structure Reference
<details>
<summary><b>View JSON Schema</b></summary>
````json
{
  "Announcements": {
    "Default": {
      "Text": "Main announcement text",
      "Details": "Optional detailed information",
      "Links": [
        { "Name": "Link text", "Url": "https://..." }
      ]
    },
    "Targeted": [
      {
        "Text": "Targeted message",
        "Details": "Details for specific group",
        "Links": [],
        "Enabled": true,
        "AppendToDefault": true,
        "Condition": {
          "Type": "Registry",
          "Path": "HKLM:\\SOFTWARE\\MITLL\\Targeting",
          "Name": "Group",
          "Value": "Pilot"
        }
      }
    ]
  },
  "Support": {
    "Text": "Contact IT Support at...",
    "Links": [
      { "Name": "Knowledge Base", "Url": "https://..." }
    ]
  }
}
````
</details>

🤝 Contributing
We welcome contributions! Areas for improvement:

 Additional templates
 Dark mode support
 Keyboard shortcuts (Ctrl+S to save, etc.)
 Multi-language support
 Export/import JSON files
 Version history browser

Development Setup
bash# Clone repository
git clone https://github.com/yourusername/endpoint-advisor-editor.git

# Edit HTML file
code editor.html

# Test locally
open editor.html
```

---

## 📞 Support

### Getting Help

| Resource | Link |
|----------|------|
| 📖 Documentation | Click **❓ Help** in editor |
| 🐛 Issues | [GitHub Issues](https://github.com/yourusername/endpoint-advisor-editor/issues) |
| 📧 Email | endpointengineering@ll.mit.edu |

### System Requirements

| Component | Requirement |
|-----------|-------------|
| Browser | Chrome 90+, Firefox 88+, Edge 90+, Safari 14+ |
| JavaScript | Must be enabled |
| localStorage | Required for auto-save |
| Network | HTTPS access to GitHub |

---

## 📄 License

This project is part of the Lincoln Laboratory Endpoint Advisor suite.
```
Copyright (c) 2025 Lincoln Laboratory
Licensed under MIT License

🙏 Acknowledgments

Built for Lincoln Laboratory IT Operations
Integrates with BigFix endpoint management
Designed for PowerShell-based endpoint agents
Community feedback and contributions


<p align="center">
  <sub>Version 1.0 | Last Updated: January 2025</sub><br>
  <sub>Made with ❤️ for endpoint administrators</sub>
</p>
