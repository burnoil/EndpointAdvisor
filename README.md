# SHOT (System Health Observation Tool)
### SHOT (System Health Observation Tool) is a PowerShell-based application designed to monitor system health, compliance, and provide real-time alerts. Built with WPF, it runs as a system tray utility, offering a lightweight yet powerful way to keep tabs on critical system metrics, YubiKey certificate status, and organizational announcements. This tool was written for MIT Lincoln Laboratory.

## Features

- System Monitoring: Tracks logged-in user, machine type, OS version, uptime, disk space, and IP addresses.
- Compliance Checks: Monitors antivirus, BitLocker, BigFix, Code42, and FIPS status with a visual indicator.
- Real-Time Alerts: Notifies users of YubiKey certificate expirations (configurable threshold, default 7 days) via tray       balloon tips.
- Announcements: Displays updates with a red dot alert, fetched from a configurable JSON source.
- Tray Integration: Collapsible context menu with quick actions (Show Dashboard, Refresh, Export Logs, Exit).
- Logging: Detailed logs saved to SHOT.log with rotation support.
- Async Operations: Non-blocking YubiKey certificate checks for smooth UI performance.

## Tech

- Written in Powershell and Windows Presentation Foundation (WPF)
- Leverages .NET 4 or higher

## Prerequisites
- Windows OS: Tested on Windows 10/11. 
- PowerShell: Version 5.1 or later (pre-installed on Windows). 
- YubiKey Manager: Optional, for YubiKey certificate monitoring (default path: C:\Program Files\Yubico\Yubikey       Manager\ykman.exe).

## Running it
- Icons: icon.ico (main) and warning.ico (non-healthy state) in the script directory.
- Installation Place Icons: Copy icon.ico and warning.ico to the script directory (e.g., C:\SHOT). If missing, the app falls back to default system icons.
- Run the Script: Open PowerShell as Administrator (recommended for full system access). Execute: powershell Unwrap Copy .\SHOT.ps1
- Ensure Tray Visibility (Optional): Right-click the taskbar → "Taskbar settings" → "Notification area" → "Select which icons appear on the taskbar". Set "SHOT" to "On" for persistent visibility.
- Usage Tray Icon: Displays icon.ico when healthy, warning.ico if issues are detected. Left-click to toggle the dashboard; right-click for the context menu. 
- Dashboard: Expand sections (e.g., "Information", "Compliance") to view details. 
- Announcements: Red dot appears on new updates; expand to clear. YubiKey Alerts: Balloon tip shown when certificate nears expiry (default: ≤7 days). 
- Logs: View recent logs in the "Logs" section or export via the tray menu. Configuration The app uses SHOT.config.json for settings. Default configuration:

# SHOT.config.json
```json
{
    "LogRotationSizeMB":  5,
    "YubiKeyLastCheck":  {
                             "Date":  "2025-03-03 07:43:29",
                             "Result":  "YubiKey Certificate: Unable to determine expiry date - No certificate found in slots 9a, 9c, 9d, or 9e"
                         },
    "DefaultLogLevel":  "INFO",
    "IconPaths":  {
                      "Main":  "icon.ico",
                      "Warning":  "warning.ico"
                  },
    "Version":  "1.1.0",
    "AnnouncementsLastState":  {
                                   "Text":  "This is a test Announcement. Written on 3/3/2025.",
                                   "Details":  "This space for rent. Inquire within.",
                                   "Links":  {
                                                 "Link1":  {
                                                               "Name":  "MITLL Website",
                                                               "Url":  "https://www.ll.mit.edu/"
                                                           },
                                                 "Link2":  {
                                                               "Name":  "NIST Cyberframework Website",
                                                               "Url":  "https://www.nist.gov/cyberframework"
                                                           }
                                             }
                               },
    "YubiKeyAlertDays":  7,
    "RefreshInterval":  30,
    "ContentDataUrl":  "https://Some_Hosted_Repository_On_Git/ContentData.json"
}
```
# ContentData.json (Can be hosted in a GIT repository or some other share/URL)
```json
{
  "Announcements": {
    "Text": "System Monitor v1.1 released on 2025-03-01! This data is from the JSON on GIT.",
    "Links": {
      "Link1": { "Name": "MIT Lincoln Lab", "Url": "https://www.ll.mit.edu/" },
      "Link2": { "Name": "NIST Cyber Framework", "Url": "https://www.nist.gov/cyberframework" }
    }
  },
  "Support": {
    "Text": "For assistance, contact IT Support at support@company.com or call 1-800-555-1234.",
    "Links": {
      "Link1": { "Name": "Knowledge Base", "Url": "https://support.company.com/knowledge-base" },
      "Link2": { "Name": "Submit Ticket", "Url": "https://support.company.com/submit-ticket" }
    }
  },
  "EarlyAdopter": {
    "Text": "Join our Early Adopter Program to test upcoming features! Sign up now.",
    "Links": {
      "Link1": { "Name": "Register", "Url": "https://beta.company.com/register" },
      "Link2": { "Name": "More Info", "Url": "https://beta.company.com/details" }
    }
  }
}
```

## License

MIT
