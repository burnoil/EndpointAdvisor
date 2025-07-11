# LLNOTIFY - Application Overview
<img width="331" height="502" alt="image" src="https://github.com/user-attachments/assets/58f7d58d-ad06-4472-9290-abfbe7b8c516" />


## Purpose

LLNOTIFY is a PowerShell-based system monitoring application designed to provide real-time system health information, certificate status, and organizational announcements to Windows endpoint users. Running as a system tray application with a user-friendly WPF (Windows Presentation Foundation) interface, MITSI enhances endpoint visibility and user support by delivering critical system metrics and IT communications directly to the desktop.

## Key Features

1. **System Tray Integration**:
   - Displays a system tray icon (`healthy.ico` or `warning.ico`) indicating system status.
   - Provides a context menu with actions: Show Dashboard, Refresh Now, Export Logs, and Exit.
   - Configured to "Always Show" in the notification area via registry settings.

2. **Graphical User Interface**:
   - A WPF-based dashboard with expandable sections for:
     - **Information**: Displays system metrics (e.g., logged-in user, machine type, OS version, uptime, disk usage, IP addresses).
     - **Announcements**: Shows organizational news with hyperlinks and source indicators (Cache, Remote, Default).
     - **Patching and Updates**: Reports patch status from a local file (e.g., `C:\temp\patch_fixlets.txt`).
     - **Support**: Provides IT contact details and links.
     - **Early Adopter**: Promotes beta programs with links.
     - **Compliance**: Monitors YubiKey and Microsoft Virtual Smart Card certificate expiry.
     - **Logs**: Displays recent log entries with export functionality.
     - **About**: Shows version, changelog, and copyright.

3. **Content Fetching**:
   - Retrieves dynamic content (Announcements, Support, Early Adopter) from a configurable URL (e.g., GitHub raw JSON) or local/network path.
   - Caches content to reduce network load, with a configurable fetch interval (`ContentFetchInterval`, default 120 seconds).
   - Falls back to default content if fetching fails, ensuring continuity.

4. **Certificate Monitoring**:
   - Checks YubiKey certificate expiry using `ykman.exe`.
   - Monitors Microsoft Virtual Smart Card certificates in user and machine stores.
   - Displays combined expiry status in the Compliance section.

5. **Logging and Diagnostics**:
   - Logs all operations to `MITSI.log` with rotation at 5MB.
   - Supports log export via a GUI button for troubleshooting.
   - Includes detailed error handling and debugging information.

## Functionality

- **System Monitoring**: Collects and displays real-time system metrics using PowerShell cmdlets (e.g., `Get-CimInstance`, `Get-NetIPAddress`), ensuring accurate data like Windows 11 24H2 detection (fixed April 22, 2025).
- **Content Updates**: Fetches external JSON content periodically, with source indicators to show whether data is from cache, remote, or default (added April 22, 2025).
- **User Interaction**: Runs in the user context to display a system tray icon and GUI, with silent execution (`-WindowStyle Hidden`) for minimal disruption.
- **Configuration**: Uses `MITSI.config.json` for customizable settings (e.g., `ContentDataUrl`, `ContentFetchInterval`), with defaults in the script for reliability.
- **Deployment**: Designed for mass deployment via BigFix, with scripts to deploy files to `C:\ProgramData\MITSI` and run via a scheduled task at user logon.
- **Updates**: Supports version upgrades (e.g., from 1.1.0 to 1.2.0) with BigFix Tasks to replace files while preserving `MITSI.config.json`.

## Use Case

MITSI is ideal for enterprise environments with thousands of Windows endpoints, providing IT teams with a lightweight tool to:
- Monitor system health and certificate compliance.
- Communicate announcements and support details to users.
- Integrate with BigFix for deployment and patch reporting (aligned with your March 14, 2025, BigFix usage).
- Ensure reliable content delivery with caching and fallback mechanisms (addressing your April 22, 2025, URL issues).

## Technical Requirements

- **OS**: Windows 10/11.
- **PowerShell**: Version 5.1 (Windows PowerShell) for `System.Windows.Forms` and WPF.
- **Dependencies**: Icon files (`healthy.ico`, `warning.ico`), optional `MITSI.config.json`, and `ykman.exe` for YubiKey checks.
- **Deployment**: BigFix for mass deployment to `C:\ProgramData\MITSI`.

---

**End of Overview**
