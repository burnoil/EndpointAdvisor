# Endpoint Advisor - Application Overview
<img width="331" height="502" alt="image" src="https://github.com/user-attachments/assets/58f7d58d-ad06-4472-9290-abfbe7b8c516" />


## Purpose

EndPoint Advisor is a PowerShell-based messaging application designed to provide real-time organizational announcements to Windows endpoint users. Running as a system tray application with a user-friendly WPF (Windows Presentation Foundation) interface, EndPoint Advisor enhances endpoint visibility and user support by delivering important ISD communications directly to the desktop with minimal interruption.

## Key Features

1. **System Tray Integration**:
   - Displays a system tray icon (`LL_LOGO.ico` or `LL_LOGO_MSG.ico`) indicating system status.
   - Provides a context menu with actions: Show Dashboard, Refresh Now, and Exit.
   - Configured to "Always Show" in the notification area via registry settings (deployment-dependent).

2. **Graphical User Interface**:
   - A WPF-based dashboard with expandable sections for:
     - **Announcements**: Shows organizational news with hyperlinks, details, and source indicators (Remote or Default).
     - **Patching and Updates**: Reports patch status from a local file (e.g., `C:\temp\X-Fixlet-Source_Count.txt`), pending restart status, and a button to launch BigFix SSA.
     - **Support**: Provides IT contact details and links.
     - **Compliance**: Monitors YubiKey and Microsoft Virtual Smart Card certificate expiry.
     - **Windows Build**: Displays OS build information.
     - **Script Update**: Shows update status for the application itself.

3. **Content Fetching**:
   - Retrieves dynamic content (Announcements, Support) from a configurable URL (e.g., GitHub raw JSON).
   - Caches content with state tracking to detect changes and alert users via UI indicators.
   - Falls back to default content if fetching fails, ensuring continuity.
   - Refresh interval configurable (`RefreshInterval`, default 900 seconds/15 minutes). Security updates are checked each cycle.

4. **Certificate Monitoring**:
   - Checks YubiKey certificate expiry using `ykman.exe` across multiple PIV slots.
   - Monitors Microsoft Virtual Smart Card certificates in the user store.
   - Displays combined expiry status in the Compliance section, with caching for performance (interval: 86400 seconds).

5. **Logging and Diagnostics**:
   - Logs all operations to `EndPoint Advisor.log` with rotation at 2MB (configurable).
   - Supports detailed error handling, retry logic, and .NET version logging.
   - No GUI log viewer; logs are file-based for troubleshooting.

6. **Auto-Update**:
   - Checks for updates via a version file on GitHub.
   - Automatically downloads and replaces the script if a newer version is available, with a restart via a temporary batch file.
   - Displays update status in the UI.
7. **BESClient Local API Setup**:
   - Configures the registry keys required for the BigFix BESClient local API.
   - Restarts the `besclient` service so the API settings take effect.
8. **Security Update Detection**:
   - Queries the BESClient local API for relevant security fixlets.
   - Displays the list of security updates in the Patching section.

## Functionality

- **System Monitoring**: Collects and displays real-time system metrics using PowerShell cmdlets (e.g., registry checks for pending restarts, OS build detection).
- **Content Fetching**: Fetches external JSON content periodically using asynchronous jobs, with source indicators to show whether data is from remote or default.
- **User Interaction**: Runs in the user context to display a system tray icon and GUI, with silent execution (`-WindowStyle Hidden`) for minimal disruption.
- **Configuration**: Uses `EndPoint Advisor.config.json` for customizable settings (e.g., `ContentDataUrl`, `RefreshInterval`), with defaults in the script for reliability.
- **Deployment**: Designed for mass deployment via BigFix or similar, with files in a directory like `C:\ProgramData\EndPoint Advisor` and run via a scheduled task at user logon.
- **Updates**: Supports auto-updates from GitHub, with version checks and seamless replacement (introduced in version 4.3.5).

## Use Case

EndPoint Advisor is ideal for enterprise environments with thousands of Windows endpoints, providing IT teams with a lightweight tool to:
- Monitor system health and certificate compliance.
- Communicate announcements and support details to users.
- Integrate with BigFix for deployment and patch reporting.
- Ensure reliable content delivery with caching and fallback mechanisms.
- Self-update to keep the tool current without manual intervention.

## Technical Requirements

- **OS**: Windows 10/11.
- **PowerShell**: Version 5.1 (Windows PowerShell) for `System.Windows.Forms` and WPF.
- **Dependencies**: Icon files (`LL_LOGO.ico`, `LL_LOGO_MSG.ico`), optional `EndPoint Advisor.config.json`, and `ykman.exe` for YubiKey checks.
- **Deployment**: BigFix or similar for mass deployment to a program data directory.

---

**End of Overview**
