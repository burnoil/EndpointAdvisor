Lincoln Laboratory Endpoint Advisor - User Guide
Article ID: KB-LLEA-001
Category: Software & Applications
Audience: All Lincoln Laboratory Users
Last Updated: October 2025

Overview
The Lincoln Laboratory Endpoint Advisor is a system tray application that provides important information about your computer's health, security, and maintenance status. The application runs automatically in the background and displays notifications when action is needed.

What Does It Do?
The Endpoint Advisor helps you stay informed about:

Laboratory Announcements - Important IT updates and notifications
Software Updates - Available application and Windows patches
Driver Updates - Windows driver maintenance (required monthly)
System Status - Pending restarts and system health
Support Information - Quick access to IT support resources
Certificate Expiration - YubiKey certificate alerts (14 days before expiration)


How to Access
System Tray Icon
The Endpoint Advisor runs in your system tray (bottom-right corner of your screen, near the clock).
Icon States:

🔵 Blue LL Logo - Everything is up to date, no alerts
🔴 Red LL Logo - New announcements or updates available

Note: If you don't see the icon, click the ^ arrow in the system tray to expand hidden icons, then drag the Endpoint Advisor icon to the main taskbar to keep it visible.
Opening the Dashboard
Click the icon to open the Endpoint Advisor dashboard window.
Right-click the icon for quick options:

Show Dashboard
Refresh Now
Set Update Interval (10, 15, or 20 minutes)
Exit


Using the Application
Dashboard Sections
The dashboard is organized into expandable sections. Click any section header to expand or collapse it.
1. Announcements

Displays important Laboratory-wide IT announcements
Red dot (🔴) indicates new or unread announcements
May include links to additional information
Content updates automatically from ISD

2. Patching and Updates
This section shows three types of updates:
Application Updates (BigFix)

Shows available application patches
Click "App Updates" button to open BigFix Self-Service
Install updates when convenient (usually doesn't require restart)

Windows OS Patches

Shows pending Windows and Office updates
Click "Install Patches" button to open Microsoft Software Center
These updates typically require a restart

Windows Driver Updates

Required once per month to keep your system secure and stable
Status shows when drivers were last updated
Button appears when updates are 30+ days old
Important: Your computer will automatically restart 5 minutes after driver installation completes

3. Support

Quick access to IT support contact information
Red dot (🔴) indicates new or updated support information
May include links to self-service resources

4. Certificate Status

Monitors YubiKey certificate expiration
You'll receive a popup alert 14 days before expiration
Contact IT Support if you receive a certificate expiration warning


Common Tasks
Installing Driver Updates
Driver updates are required monthly and ensure your computer has the latest hardware compatibility and security fixes.
Prerequisites:

Save all your work - Your computer will restart automatically
Plug into AC power - Driver updates require being plugged in (not on battery)

Steps:

Open the Endpoint Advisor dashboard
Expand the "Patching and Updates" section
Click the "Install Drivers" button
Confirm the installation when prompted
The progress panel will show the installation status
Your computer will automatically restart in 5 minutes after completion

Progress Indicators:

"Scanning for available driver updates..."
"Downloading driver updates..."
"Installing driver updates..."
"No driver updates are currently available" (if system is up to date)
"Installation complete. Your computer will restart in 5 minutes."

Clearing Alerts
The red dots (🔴) on sections indicate new content you haven't viewed yet.
To clear alerts:

Option 1: Open each section with a red dot to view the new content
Option 2: Click the "Clear Alerts" button at the bottom of the dashboard

Changing Update Frequency
The dashboard automatically checks for new content every 15 minutes by default.
To change the interval:

Right-click the system tray icon
Select "Set Update Interval"
Choose 10, 15, or 20 minutes


Troubleshooting
The icon is missing from my system tray
Solution: The icon may be in the hidden icons area.

Click the ^ (up arrow) in your system tray
Look for the Lincoln Laboratory logo
Drag the icon to your main taskbar to keep it visible


I can't install driver updates - "AC Power Required" message
Solution: Driver updates require your laptop to be plugged into AC power.

Connect your laptop to its power adapter
Wait a few seconds for Windows to recognize AC power
Try the driver update again


The dashboard shows "No announcements at this time"
Possible Causes:

You may be on VPN or off-network when the app starts
Network connectivity issues

Solutions:

Right-click the icon and select "Refresh Now"
Ensure you're connected to the network (on-site or via GlobalProtect VPN)
If the issue persists, restart the application or contact IT Support


Driver update button doesn't appear
This is normal! The driver update button only appears when:

It has been 30+ days since the last driver update
Driver updates have never been run on your computer

If drivers were recently updated, you'll see the status "Last run X days ago" and no button will appear.

"Pending Restart Status" shows but I just restarted
Solution: Windows sometimes requires multiple restarts after updates.

Restart your computer again
The status should clear after the restart completes
If the message persists after 2-3 restarts, contact IT Support


Frequently Asked Questions
Q: Can I close the dashboard window?
A: Yes! Closing the window hides it. The application continues running in the system tray. Click the icon to show it again.
Q: Will this slow down my computer?
A: No. The application uses minimal resources and only checks for updates every 15 minutes.
Q: Can I uninstall or disable this application?
A: The Endpoint Advisor is a required Laboratory application and should not be disabled. If you have concerns, please contact IT Support.
Q: How often should I install driver updates?
A: Driver updates are required once per month. The application will remind you when it's time.
Q: What if I'm in the middle of important work when drivers finish installing?
A: You'll receive a 5-minute warning before the automatic restart. Save your work immediately when you see this notification. Consider running driver updates during lunch or at the end of your workday.
Q: Can I schedule driver updates for a specific time?
A: Currently, driver updates run immediately when initiated. Plan to run them when you can afford a 5-minute interruption and automatic restart.
Q: What if I receive a certificate expiration warning?
A: Contact IT Support immediately at 781-981-4357 (HELP) to renew your certificate before it expires. Certificate expiration can prevent you from accessing Laboratory systems.
