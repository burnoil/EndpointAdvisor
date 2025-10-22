# 🎯 Lincoln Laboratory Endpoint Advisor - Web Editor

A powerful, user-friendly web-based editor for managing endpoint announcements. **No JSON knowledge required!**

![Version](https://img.shields.io/badge/version-1.0-blue.svg)
![Status](https://img.shields.io/badge/status-active-success.svg)

---

## 📋 Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Quick Start](#quick-start)
- [Configuration](#configuration)
- [Usage Guide](#usage-guide)
- [Troubleshooting](#troubleshooting)
- [Support](#support)

---

## Overview

The **Endpoint Advisor Web Editor** is a single-page HTML application for managing JSON configuration files that drive the [Lincoln Laboratory Endpoint Advisor](https://github.com/burnoil/EndpointAdvisor) PowerShell application.

### Why Use This Editor?

| Before | After |
|--------|-------|
| ❌ Manual JSON editing | ✅ Visual form interface |
| ❌ Syntax errors | ✅ Automatic validation |
| ❌ No preview | ✅ Live preview |
| ❌ Lost work on crash | ✅ Auto-save |
| ❌ Blind commits | ✅ Diff viewer |

---

## ✨ Features

### 📝 Dual Editing Modes
- **Form Editor** for non-technical users
- **JSON Editor** for power users
- Instant sync between modes

### 👁️ Live Preview
- Real-time rendering as you type
- Markdown support (bold, colors, links)
- Exact endpoint simulation

### 💾 Smart Saving
- Auto-save every 2 seconds to browser
- Diff viewer before commits
- Restore lost work instantly

### 🎯 Targeted Announcements
- Registry-based conditional messages
- Built-in BigFix Action Script templates
- One-click copy buttons

### 📋 Templates
6 ready-to-use templates:
- 🔧 System Maintenance
- 🔒 Security Alerts
- ⚙️ Patching Notices
- 📚 Training Reminders
- 🎉 Holiday Schedules
- 🚀 Pilot Programs

### 🔧 GitHub Integration
- Public & Enterprise GitHub support
- Multi-instance dropdown selector
- Token permission diagnostics

---

## 🚀 Quick Start

### Prerequisites

- GitHub account with repository access
- Personal Access Token with "repo" scope
- Modern web browser

### Installation

**Option 1: Download and Open**
```bash
# Download the HTML file
wget https://raw.githubusercontent.com/burnoil/EndpointAdvisor/main/editor.html

# Open in browser
open editor.html
```

**Option 2: Host on Server**
```bash
# Place on your web server
cp editor.html /var/www/html/
# Access at: https://your-server/editor.html
```

### First-Time Setup

#### 1. Create GitHub Personal Access Token

1. Go to [GitHub Settings → Tokens](https://github.com/settings/tokens)
2. Click **"Generate new token (classic)"**
3. Check the **`repo`** scope
4. For Enterprise with SSO: Click **"Authorize SSO"**
5. Copy the token

> ⚠️ **Security:** Never commit tokens to repositories!

#### 2. Configure Repository Connection

| Field | Example |
|-------|---------|
| GitHub Instance | Internal GitHub Enterprise |
| Repository Owner | `burnoil` |
| Repository Name | `EndpointAdvisor` |
| Branch | `main` |
| File Path | `ContentData.json` |
| Token | `ghp_xxxxx...` |

#### 3. Load and Edit

1. Click **"Load File"**
2. Switch to **"Form Editor (Easy)"** mode
3. Edit announcements
4. View live preview
5. Click **"Save to GitHub"** → Review diff → Confirm

---

## 🔧 Configuration

### Customize GitHub Instance Dropdown

Edit the HTML file to add your organization's GitHub instances:
```html
<select id="apiBase" onchange="handleApiBaseChange()">
  <option value="https://api.github.com">Public GitHub</option>
  <option value="https://github.ll.mit.edu/api/v3" selected>Production</option>
  <option value="https://github-dev.ll.mit.edu/api/v3">Development</option>
  <option value="custom">Custom URL...</option>
</select>
```

**API URL Format:** `https://your-github-domain/api/v3`

### Pre-fill Default Values
```html
<input type="text" id="repoOwner" value="burnoil">
<input type="text" id="repoName" value="EndpointAdvisor">
<input type="text" id="branch" value="main">
<input type="text" id="filePath" value="ContentData.json">
```

---

## 📖 Usage Guide

### Editing Announcements

#### Using Form Editor (Recommended)

1. Switch to **"📝 Form Editor (Easy)"** mode
2. Type announcement text
3. Use toolbar for formatting: **[B]** **[I]** 🟢 🔴 🟡
4. Add optional details
5. Add links with **"+ Add Link"** button
6. Preview updates automatically

#### Markdown Formatting

| Syntax | Result |
|--------|--------|
| `**bold**` | **bold** |
| `*italic*` | *italic* |
| `__underline__` | underlined |
| `[green]text[/green]` | green text |
| `[red]warning[/red]` | red text |
| `[yellow]caution[/yellow]` | yellow text |

### Creating Targeted Announcements

**Targeted announcements** appear only on computers with specific registry keys.

#### In the Editor:

1. Scroll to **"🎯 Targeted Announcements"** section
2. Click **"+ Add Targeted Announcement"**
3. Check **"Enabled"**
4. Check **"Append to Default"** (or uncheck to replace)
5. Enter text and details
6. Set Registry Condition:
   - **Path:** `HKLM:\SOFTWARE\MITLL\Targeting`
   - **Name:** `Group`
   - **Value:** `Pilot`

#### Deploy Registry Keys (BigFix):

Copy from Help → Targeted Announcements → BigFix Scripts:
```actionscript
regset "[HKEY_LOCAL_MACHINE\SOFTWARE\MITLL\Targeting]" "Group"="Pilot"
```

Deploy to target computers using BigFix Actions.

#### Verify on Endpoint:
```powershell
Get-ItemProperty -Path "HKLM:\SOFTWARE\MITLL\Targeting"
```

---

## 🎯 Common Use Cases

### Example 1: Monthly Patch Announcement
```
Template: Patching Notice

Text: **Monthly Patch Deployment** - Windows and Office updates available

Details: 
**Starting today at 3pm** patches will be deployed.
Updates install overnight. 
[yellow]**Computers will reboot automatically.**[/yellow]
```

### Example 2: Emergency Security Alert
```
Template: Security Alert

Text: [red]**SECURITY ALERT**[/red] - Action Required

Details:
Critical security update released.

**Action Required:**
- Update system immediately
- Restart when prompted
```

### Example 3: Pilot Program (Targeted)
```
Targeted Announcement

Text: [green]**Pilot Program**[/green] - You've been selected!

Details:
You're part of the VPN 2.0 pilot group.

**Next Steps:**
- Install pilot software by Friday
- Review documentation

Registry: Group = VPN-Pilot
Append to Default: ✅ Yes
```

---

## 🔍 Troubleshooting

### "GitHub API error: 401" when saving

**Possible Causes:**
- Token missing "repo" scope
- Token expired
- SSO not authorized (Enterprise)
- Wrong GitHub instance selected

**Solutions:**
1. Click **"🔑 Check Token Permissions"** to diagnose
2. Verify token has **"repo"** scope
3. For Enterprise: **Authorize SSO** in token settings
4. Create fresh token
5. Verify GitHub instance dropdown

### Changes not appearing on endpoints

**Timeline:** The app refreshes every 10-20 minutes

**Verify:**
```bash
# Check GitHub commit history
git log --oneline -5

# Verify endpoint can reach GitHub
Test-NetConnection github.yourcompany.com -Port 443
```

### Targeted announcement not showing

**Checklist:**
- ✅ Announcement marked "Enabled"
- ✅ Registry key exists on target computer
- ✅ Registry path/name/value match exactly (case-sensitive)
- ✅ BigFix Action deployed

**Verify on endpoint:**
```powershell
Get-ItemProperty -Path "HKLM:\SOFTWARE\MITLL\Targeting"
```

### Lost work / browser crashed

**Recovery:**
1. Reopen editor
2. Click **"💾 Restore Auto-save"**
3. Auto-save runs every 2 seconds!

---

## 🛡️ Security Best Practices

**✅ Do:**
- Use tokens with minimum permissions
- Enable SSO and MFA
- Rotate tokens regularly
- Use HTTPS for hosted editors
- Clear auto-save after publishing

**❌ Don't:**
- Share tokens with others
- Store tokens in repositories
- Use admin tokens for routine edits
- Expose editor on public internet

**Token Storage:**
- Stored in browser memory only
- Never saved to disk
- Cleared when browser closes

---

## 📚 Documentation

### Built-in Help System

Click **❓ Help** button for comprehensive docs:

- 📖 Getting Started
- 📝 Markdown Formatting
- 🎯 Targeted Announcements
- ✨ Best Practices
- 🔍 Troubleshooting

### JSON Structure
```json
{
  "Announcements": {
    "Default": {
      "Text": "Main announcement",
      "Details": "Optional details",
      "Links": [
        { "Name": "Link text", "Url": "https://..." }
      ]
    },
    "Targeted": [
      {
        "Text": "Targeted message",
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
    "Text": "Contact IT Support...",
    "Links": []
  }
}
```

---

## 🤝 Contributing

We welcome contributions!

**Ideas for improvement:**
- Additional templates
- Dark mode support
- Keyboard shortcuts
- Multi-language support
- Export/import JSON files

---

## 📞 Support

| Resource | Link |
|----------|------|
| Documentation | Click **❓ Help** in editor |
| Issues | [GitHub Issues](https://github.com/burnoil/EndpointAdvisor/issues) |
| Email | endpointengineering@ll.mit.edu |

**System Requirements:**
- Chrome 90+, Firefox 88+, Edge 90+, Safari 14+
- JavaScript enabled
- localStorage for auto-save
- HTTPS access to GitHub

---

## 📄 License

Copyright (c) 2025 Lincoln Laboratory  
Part of the Lincoln Laboratory Endpoint Advisor suite

---

## 🙏 Acknowledgments

- Built for **Lincoln Laboratory IT Operations**
- Integrates with **BigFix** endpoint management
- Designed for **PowerShell**-based endpoint agents

---

**Version 1.0** | Last Updated: January 2025  
Made with ❤️ for endpoint administrators
