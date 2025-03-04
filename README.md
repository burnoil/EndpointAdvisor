# SHOT - System Health Observation Tool

## SHOT (System Health Observation Tool) is a PowerShell-based application designed to monitor system health, compliance, and provide real-time alerts. Built with WPF, it runs as a system tray utility, offering a ## lightweight yet powerful way to keep tabs on critical system metrics, YubiKey certificate status, and organizational announcements. This tool was written for MIT Lincoln Laboratory.

## Features
System Monitoring: Tracks logged-in user, machine type, OS version, uptime, disk space, and IP addresses.
Compliance Checks: Monitors antivirus, BitLocker, BigFix, Code42, and FIPS status with a visual indicator.
Real-Time Alerts: Notifies users of YubiKey certificate expirations (configurable threshold, default 7 days) via tray balloon tips.

## Announcements: Displays updates with a red dot alert, fetched from a configurable JSON source.
Tray Integration: Collapsible context menu with quick actions (Show Dashboard, Refresh, Export Logs, Exit).

## Logging: Detailed logs saved to SHOT.log with rotation support.

## Async Operations: Non-blocking YubiKey certificate checks for smooth UI performance.

## Versioning: Includes a changelog in the "About" section (current version: 1.1.0).

Prerequisites
Windows OS: Tested on Windows 10/11.
PowerShell: Version 5.1 or later (pre-installed on Windows).
YubiKey Manager: Optional, for YubiKey certificate monitoring (default path: C:\Program Files\Yubico\Yubikey Manager\ykman.exe).

Icons: icon.ico (main) and warning.ico (non-healthy state) in the script directory.

Installation
Place Icons:
Copy icon.ico and warning.ico to the script directory (e.g., C:\SHOT\).
If missing, the app falls back to default system icons.

Run the Script:
Open PowerShell as Administrator (recommended for full system access).
Execute:
powershell
Unwrap
Copy
.\SHOT.ps1

Ensure Tray Visibility (Optional):
Right-click the taskbar → "Taskbar settings" → "Notification area" → "Select which icons appear on the taskbar".
Set "SHOT" to "On" for persistent visibility.

Usage
Tray Icon:
Displays icon.ico when healthy, warning.ico if issues are detected.
Left-click to toggle the dashboard; right-click for the context menu.
Dashboard: Expand sections (e.g., "Information", "Compliance") to view details.
Announcements: Red dot appears on new updates; expand to clear.
YubiKey Alerts: Balloon tip shown when certificate nears expiry (default: ≤7 days).
Logs: View recent logs in the "Logs" section or export via the tray menu.
Configuration
The app uses SHOT.config.json for settings. Default configuration:

json

{
  "RefreshInterval": 30,
  "LogRotationSizeMB": 5,
  "DefaultLogLevel": "INFO",
  "ContentDataUrl": "ContentData.json",
  "YubiKeyAlertDays": 7,
  "IconPaths": {
    "Main": "icon.ico",
    "Warning": "warning.ico"
  },
  "YubiKeyLastCheck": {
    "Date": "1970-01-01 00:00:00",
    "Result": "YubiKey Certificate: Not yet checked"
  },
  "AnnouncementsLastState": {},
  "Version": "1.1.0"
}

RefreshInterval: Update frequency in seconds (default: 30).
ContentDataUrl: Path or URL to JSON announcements (local, network, or HTTP).
YubiKeyAlertDays: Days before YubiKey expiry to trigger alerts (default: 7).

Example ContentData.json

{
  "Announcements": {
    "Text": "SHOT v1.1 released!",
    "Details": "New features added.",
    "Links": {
      "Link1": "https://example.com/news1",
      "Link2": "https://example.com/news2"
    }
  },
  "Support": {
    "Text": "Contact IT: support@example.com",
    "Links": {
      "Link1": "https://support.example.com",
      "Link2": "https://tickets.example.com"
    }
  }
}

Troubleshooting
Icon Missing: Ensure icon.ico and warning.ico are in the script directory.
YubiKey Not Detected: Verify ykman.exe path in Start-YubiKeyCertCheckAsync.
UI Freezes: Check SHOT.log for errors; async updates should prevent this.
Tray Hidden: Follow the visibility instructions above.

Changelog
v1.1.0: Added tooltips, collapsible tray menu, status indicators, YubiKey alerts, async updates, versioning.
v1.0.0: Initial release as "System Monitor".
License
© 2025 SHOT. All rights reserved. MIT License.
