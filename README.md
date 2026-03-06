# Rexie Tools

```
            __
           / _) 
    .-^^^-/ / 
 __/       /  
<__.|_|-|_|   
```

PowerShell automation toolkit for endpoint support workflows.

Rexie Tools is a macOS-hosted PowerShell remote administration launcher designed for endpoint support engineers. It provides a fast menu-driven interface for running diagnostics and remediation tasks on remote Windows devices using WinRM.

⸻

Project Status

Version: 1.2.0

PowerShell | macOS Host | Windows Endpoint Automation

⸻

Features

Remote Administration Toolkit

Menu-driven automation for common endpoint support tasks.

1. Group Policy (GPO)
• gpupdate /force
• certificate auto-enrollment
• Skynet fix workflow
• certificate request tools

2. Event Log Scan
• scans System and Application logs
• groups errors and warnings for quick troubleshooting

3. Computer Information

Displays detailed hardware and system data:

• model and serial number
• OS version and build
• RAM usage
• CPU speed
• system uptime
• dock detection
• monitor detection
• network connection type
• logged-in user

4. Dell Command Update

• runs Dell Command Update remotely
• streams update progress to console

5. Scheduled Reboot

• reboot immediately
• schedule a one-time reboot using scheduled tasks

6. Hostname Reservation Assistant

• integrates with Rex hostname reservation API
• determines next available device number automatically
• includes self-heal logic if the API service is offline

7. Login Shark

Session intelligence tool used to determine reboot safety.

Displays:

• active user session
• last logged-on user
• recent authentication events
• lock/unlock status
• reboot safety recommendation

8. SCCM / Software Center Actions

Remotely trigger Configuration Manager client cycles:

• Machine Policy Retrieval
• User Policy Retrieval
• Application Deployment Evaluation
• Software Update Evaluation
• Hardware Inventory
• Full Client Check-In

9. Battery Report

• generates battery health report on remote laptops
• opens report via SMB from the admin share

⸻

Architecture

Operator System
macOS with PowerShell 7+

Target Systems
Windows endpoints with WinRM enabled

All tasks are executed using PowerShell remoting (WinRM).

⸻

Installation

Windows
	1.	Go to the Releases section of this repository
	2.	Download RexieTools.exe
	3.	Run the executable

No installation required.

⸻

macOS

Install PowerShell:

brew install –cask powershell

Clone the repository:

git clone https://github.com/c0ry-s/rexie-tools.git

Run Rexie Tools:

./run_rexietools.command

⸻

Credential Storage

Rexie Tools supports persistent credential storage.

Credentials are stored locally using encrypted PowerShell credential objects.

Location:

~/Documents/UniversalCredential.xml

These credentials are used for WinRM connections to remote Windows devices.

No credentials are stored in this repository.

⸻

Versioning

Rexie Tools follows semantic versioning.

Format:

Major.Minor.Patch

Example:

1.2.0

The script checks GitHub automatically and will notify users when a newer version is available.

⸻

License

Internal use and distribution permitted.
