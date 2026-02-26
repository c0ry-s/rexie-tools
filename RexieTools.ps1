<#!
.SYNOPSIS
    Rexie Tools – Mac-hosted PowerShell remote admin launcher for UNC Health.
.DESCRIPTION
    Presents a menu of common remote administration tasks (WinRM-based). Stores a universal
    credential (optional) in the user's Documents folder and re-uses it across tasks.
.VERSION
    1.0.0
.AUTHOR
    c0ryS (Cory Smith)
.LAST UPDATED
    2025-09-10
.REQUIREMENTS
    • PowerShell 7+
    • Network access to \\vscifs1 for version check
    • WinRM enabled on target Windows devices
.NOTES
    Universal credential path: ~/Documents/UniversalCredential.xml
#>
# Enable common parameters (-WhatIf/-Confirm) and named-only params
[CmdletBinding(
    SupportsShouldProcess = $true,
    ConfirmImpact         = 'Medium',
    PositionalBinding     = $false
)]
param()
# Helper to centralize ShouldProcess checks

function Invoke-IfShouldProcess {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory)] [string]$Target,
        [Parameter(Mandatory)] [string]$Operation,
        [Parameter(Mandatory)] [scriptblock]$Action
    )
    if ($PSCmdlet.ShouldProcess($Target, $Operation)) {
        try { & $Action }
        catch { Write-Error -ErrorRecord $_ }
    }
    else {
        Write-Verbose "Skipped: $Operation on $Target"
    }
}

# --- Standardized Console Output ---------------------------------------------
# Levels: INFO, OK, WARN, ERROR, DEBUG
# ---------------------------------------------------------------------------
function Write-Status {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [ValidateSet('INFO','OK','WARN','ERROR','DEBUG')] [string]$Level,
        [Parameter(Mandatory)] [string]$Message,
        [switch]$NoNewLine
    )

    $prefix = switch ($Level) {
        'INFO'  { '[INFO ]' }
        'OK'    { '[OK   ]' }
        'WARN'  { '[WARN ]' }
        'ERROR' { '[ERROR]' }
        'DEBUG' { '[DEBUG]' }
    }

    $color = switch ($Level) {
        'INFO'  { 'Cyan' }
        'OK'    { 'Green' }
        'WARN'  { 'Yellow' }
        'ERROR' { 'Red' }
        'DEBUG' { 'DarkGray' }
    }

    $out = "$prefix $Message"
    if ($NoNewLine) {
        Write-Host $out -ForegroundColor $color -NoNewline
    } else {
        Write-Host $out -ForegroundColor $color
    }
}

# --- Rexie Robo (Audiology Deploy) --------------------------------------------
# Copies \\vscifs1\eusfiles\Installs\Audiology
# To C:\HCSTools\Audiology on a target device using Scheduled Task
# -------------------------------------------------------------------------------
function Rexie-Robo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Hostname,
        [Parameter(Mandatory)] [pscredential]$Credential
    )

    $source = '\\vscifs1\eusfiles\Installs\Audiology'
    $dest   = 'C:\HCSTools\Audiology'

    Write-Status -Level INFO -Message "Starting Rexie Robo (Scheduled Task mode) on $Hostname"
    Write-Status -Level INFO -Message "Source: $source"
    Write-Status -Level INFO -Message "Destination: $dest"

    # Use same credential for WinRM and scheduled task run-as
    $runUser = $Credential.UserName
    $runPass = $Credential.GetNetworkCredential().Password

    try {
        $result = Invoke-Command -ComputerName $Hostname -Credential $Credential -ScriptBlock {
            param($runUser, $runPass, $src, $dst)

            New-Item -ItemType Directory -Path $dst -Force | Out-Null

            $ts  = Get-Date -Format 'yyyyMMdd-HHmmss'
            $logDir = 'C:\HCSTools\Logs\RexieRobo'
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null
            $log = Join-Path $logDir "audiology-$ts.log"

            $task = "Rexie-Robo-$ts"
            $cmdFile = Join-Path $logDir "RexieRobo-$ts.cmd"

            $cmdContent = @"
@echo off
 echo ==== Rexie Robo Start %date% %time% ====>>"$log"
 echo Host: %COMPUTERNAME%   User: %USERNAME%>>"$log"
 echo Source: $src>>"$log"
 echo Dest:   $dst>>"$log"
 echo.>>"$log"
 robocopy "$src" "$dst" *.* /E /ZB /R:5 /W:5 /COPY:DAT /DCOPY:DAT /V /FP /TS /ETA /BYTES /NP /UNILOG+:"$log" /TEE
 echo.>>"$log"
 echo ==== Rexie Robo End %date% %time%  RC=%errorlevel% ====>>"$log"
"@

            Set-Content -Path $cmdFile -Value $cmdContent -Encoding ASCII

            schtasks /Delete /TN $task /F *> $null 2>&1

            schtasks /Create /TN $task /SC ONCE /ST 23:59 /RL HIGHEST `
                /RU $runUser /RP $runPass `
                /TR "cmd.exe /c `"$cmdFile`"" /F | Out-Null

            schtasks /Run /TN $task | Out-Null

            [pscustomobject]@{
                TaskName = $task
                Log      = $log
            }
        } -ArgumentList $runUser, $runPass, $source, $dest

        Write-Status -Level OK -Message "Scheduled Task Created: $($result.TaskName)"
        Write-Status -Level INFO -Message "Log Path: $($result.Log)"
    }
    catch {
        Write-Status -Level ERROR -Message "Rexie Robo failed: $($_.Exception.Message)"
    }
}

# --- Hostname Reservation API Self-Heal Helpers --------------------------------
# Option 6 depends on a small API running on RXCRY01TECHLT01:8080.
# If the port refuses connections, attempt to start the Scheduled Task remotely
# (requires WinRM to the API host) and retry once.
# ------------------------------------------------------------------------------
$HostnameApiHost     = 'RXCRY01TECHLT01'
$HostnameApiPort     = 8080
$HostnameApiTaskName = 'Rexie Hostname API'

function Test-TcpPort {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$HostName,
        [Parameter(Mandatory)] [int]$Port,
        [int]$TimeoutMs = 1500
    )
    try {
        $client = [System.Net.Sockets.TcpClient]::new()
        $iar = $client.BeginConnect($HostName, $Port, $null, $null)
        if (-not $iar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)) {
            $client.Close()
            return $false
        }
        $client.EndConnect($iar) | Out-Null
        $client.Close()
        return $true
    } catch {
        try { $client.Close() } catch { }
        return $false
    }
}

function Start-HostnameReservationApiRemote {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$ApiHost,
        [Parameter(Mandatory)] [string]$TaskName,
        [Parameter()] [pscredential]$Credential
    )

    $sb = {
        param($TaskName)
        try {
            Start-ScheduledTask -TaskName $TaskName -ErrorAction Stop
            "Started scheduled task: $TaskName"
        } catch {
            "Failed to start scheduled task: $TaskName. $($_.Exception.Message)"
        }
    }

    try {
        if ($Credential) {
            Invoke-Command -ComputerName $ApiHost -Credential $Credential -ScriptBlock $sb -ArgumentList $TaskName -ErrorAction Stop
        } else {
            Invoke-Command -ComputerName $ApiHost -ScriptBlock $sb -ArgumentList $TaskName -ErrorAction Stop
        }
    } catch {
        "Remote start attempt failed (WinRM). $($_.Exception.Message)"
    }
}

# Define the current version of this script
$currentVersion = [version]"1.0.0"

# TODO: Break repeated code blocks into reusable functions for maintainability.

#region Version Check & Banner
# Define the remote version file path
$remoteVersionFile = "\\vscifs1\eusfiles\Corys Home\Version Checker\Rexie Tools\version.txt"
Write-Status -Level INFO -Message "-=*Rexie Tools by c0ryS*=-"
Write-Host @"
            __
           / _) 
    .-^^^-/ / 
 __/       /  
<__.|_|-|_|   
"@ -ForegroundColor Blue
# Check if the remote version file exists and compare versions
try {
    if (Test-Path $remoteVersionFile) {
        $latestVersionString = Get-Content $remoteVersionFile -ErrorAction Stop | Select-Object -First 1
        if ($null -eq $latestVersionString -or $latestVersionString.Trim() -eq "") {
            Write-Status -Level WARN -Message "Remote version file is empty. Skipping version check."
        } else {
            $latestVersion = [version]($latestVersionString.Trim())
            Write-Status -Level INFO -Message "Current version: $currentVersion. Latest: $latestVersion."
            if ($latestVersion -gt $currentVersion) {
                Write-Status -Level WARN -Message "Update available ($latestVersion). Pull the latest file from \\vscifs1\\eusfiles\\Corys Home."
                return
            } else {
                Write-Status -Level OK -Message "Script is up to date."
            }
        }
    } else {
        Write-Status -Level WARN -Message "Remote version file not found at $remoteVersionFile. Skipping version check."
    }
} catch {
    Write-Status -Level WARN -Message "Version check error: $($_.Exception.Message)"
}
#endregion Version Check & Banner
 #region Session Loop & Credential Handling
$repeatSession = $true
do {

# --- Credential Handling ------------------------------------------------------
# Loads a stored credential from ~/Documents/UniversalCredential.xml when present.
# If not present, prompts once and optionally persists for future runs.
# -----------------------------------------------------------------------------
  # Define the universal credential file path in the user's Documents folder
   $credPath = Join-Path ([Environment]::GetFolderPath('MyDocuments')) "UniversalCredential.xml"

   # Load stored credentials if available; otherwise, prompt and optionally store them
$Cred = $null
if (Test-Path $credPath) {
    $Cred = Import-Clixml -Path $credPath
} else {
    Write-Status -Level WARN -Message "No stored credentials found."
    $Cred = Get-Credential -Message "Enter credentials for remote access"

    $storeAnswer = Read-Host "Do you want to store these credentials for future use? (Y/N)"
    if ($storeAnswer.ToUpper().StartsWith("Y")) {
        $Cred | Export-Clixml -Path $credPath
        Write-Status -Level OK -Message "Credentials stored at $credPath"
    }
}
$cred = $Cred

# --- Main Menu ---------------------------------------------------------------
    Write-Status -Level INFO -Message "Select an option:"
    Write-Host "1. Group Policy (GPO)"
    Write-Host "2. View Installed and Failed Windows Updates"
    Write-Host "3. View Computer Info"
    Write-Host "4. Run Dell Command Update"
    Write-Host "5. Schedule One-Time Reboot"
    Write-Host "6. Hostname Reservation Assistant"
    Write-Host "7. Scan Event Logs (System & Application)"
    Write-Host "8. Update Tidepool Uploader"
    Write-Host "9. Battery Report"
    Write-Host "10. Rexie Robo (Audiology Deploy)"
    $selection = Read-Host "Enter your choice (1-10, Q to exit)"

    if ($selection -match '^[Qq]$') {
        break
    } elseif ($selection -in @('6','7')) {
        $hostname = $null
    }

    switch ($selection) {
        # --- Option 8: Tidepool Uploader Updater ------------------------------------
        # Checks installed version on target, compares with latest GitHub release, then
        # downloads to C:\HCSTools\Software and runs a high-privilege scheduled task to
        # install silently. Includes retryable hostname prompt and ping check.
        # -----------------------------------------------------------------------------
        '8' {
            if ($null -eq $hostname) {
                do {
                    $hostname = Read-Host "Enter the hostname or IP address of the Windows PC"
                    if ([string]::IsNullOrWhiteSpace($hostname)) {
                        Write-Host "Hostname cannot be empty." -ForegroundColor Red
                        continue
                    }
                    $pingSuccess = Test-Connection -ComputerName $hostname -Count 1 -Quiet
                    if (-not $pingSuccess) {
                        Write-Host "Device $hostname is not reachable via ping." -ForegroundColor Yellow
                        Write-Host "`nWhat would you like to do?"
                        Write-Host "1. Try the same hostname again"
                        Write-Host "2. Enter a different hostname"
                        Write-Host "3. Return to main menu"
                        $choice = Read-Host "Select an option (1-3)"
                        switch ($choice) {
                            '1' { continue }
                            '2' { $hostname = $null; continue }
                            '3' { $hostname = $null; continue 1 }
                            default { Write-Host "Invalid input. Returning to main menu." -ForegroundColor Yellow; $hostname = $null; continue 1 }
                        }
                    }
                } while (-not $pingSuccess)
            }
            # Create remote session
            $session = New-PSSession -ComputerName $hostname -Credential $Cred

            # Test session
            if ($null -eq $session) {
                Write-Host "Failed to create a session to $hostname" -ForegroundColor Red
                exit
            }

            # Get installed Tidepool version by checking executable file version
            $installedVersion = Invoke-Command -Session $session -ScriptBlock {
                $exePath = "C:\Program Files\Tidepool Uploader\Tidepool Uploader.exe"
                if (Test-Path $exePath) {
                    (Get-Item $exePath).VersionInfo.ProductVersion
                } else {
                    ""
                }
            }
            Write-Host "DEBUG: Retrieved InstalledVersion from EXE = '$installedVersion'"

            $installedVersionTrimmed = $installedVersion -replace '^((\d+\.\d+\.\d+)).*','$1'
            Write-Host "DEBUG: Trimmed InstalledVersion for comparison = '$installedVersionTrimmed'"

            # Get latest version from GitHub
            $headers = @{ 'User-Agent' = 'RexieTools/1.0' }
            $latestRelease = Invoke-RestMethod -Uri "https://api.github.com/repos/tidepool-org/uploader/releases/latest" -Headers $headers
            $latestVersion = $latestRelease.tag_name.TrimStart("v")

            Write-Host "Installed Tidepool version on ${hostname}: $installedVersion"
            Write-Host "Latest available Tidepool version: $latestVersion"

            # Build direct GitHub download URL
            $downloadUrl = "https://github.com/tidepool-org/uploader/releases/download/v$latestVersion/Tidepool-Uploader-Setup-$latestVersion.exe"
            $remoteInstallerPath = "C:\HCSTools\Software\Tidepool-Uploader-Setup.exe"

            if ($installedVersionTrimmed -ne $latestVersion) {
                if (-not $installedVersion) {
                    Write-Host "Tidepool is not installed on ${hostname}. Latest available version is $latestVersion." -ForegroundColor Yellow
                } else {
                    Write-Host "Update recommended. Remote machine has $installedVersion, latest is $latestVersion." -ForegroundColor Yellow
                }

                # Check if installer already exists on remote machine and matches expected version
                $installerExists = Invoke-Command -Session $session -ScriptBlock {
                    $installerPath = "C:\HCSTools\Software\Tidepool-Uploader-Setup.exe"
                    if (Test-Path $installerPath) {
                        return $true
                    } else {
                        return $false
                    }
                }

                if ($installerExists) {
                    Write-Host "Installer already exists on ${hostname} at C:\HCSTools\Software. Skipping download." -ForegroundColor Cyan
                } else {
                    Invoke-Command -Session $session -ScriptBlock {
                        $downloadUrl = $using:downloadUrl
                        $installerPath = $using:remoteInstallerPath
                        Invoke-WebRequest -Uri $downloadUrl -OutFile $installerPath -Headers @{ 'User-Agent' = 'RexieTools/1.0' }
                    }
                    Write-Host "Installer downloaded to remote machine." -ForegroundColor Cyan
                }
                # Always attempt to run installer silently, headless with extra logging
                Invoke-Command -Session $session -ScriptBlock {
                    $installerPath = $using:remoteInstallerPath

                    # Stop any running Tidepool processes before install
                    Get-Process | Where-Object { $_.ProcessName -like "*tidepool*" } | Stop-Process -Force -ErrorAction SilentlyContinue

                    # Create and run the scheduled task as SYSTEM with highest privileges
                    schtasks /Create /TN "RunTidepoolInstaller" /TR "`"$installerPath /S`"" /SC ONCE /ST 00:00 /RU SYSTEM /RL HIGHEST /F
                    schtasks /Run /TN "RunTidepoolInstaller"
                    Write-Host "Scheduled task created and started via schtasks.exe as SYSTEM"
                }
                Write-Host "Installer triggered on remote machine via headless scheduled task." -ForegroundColor Cyan
            } else {
                Write-Host "Tidepool is up to date." -ForegroundColor Green
            }

            # Clean up session
            Remove-PSSession $session
        }
        # --- Option 1: Group Policy Update -------------------------------------------
        '1' {
            # Prompt for hostname if it is explicitly null
            if ($null -eq $hostname) {
                do {
                    $hostname = Read-Host "Enter the hostname or IP address of the Windows PC"
                    if ([string]::IsNullOrWhiteSpace($hostname)) {
                        Write-Host "Hostname cannot be empty." -ForegroundColor Red
                        continue
                    }
                    $pingSuccess = Test-Connection -ComputerName $hostname -Count 1 -Quiet
                    if (-not $pingSuccess) {
                        Write-Host "Device $hostname is not reachable via ping." -ForegroundColor Yellow
                        Write-Host "`nWhat would you like to do?"
                        Write-Host "1. Try the same hostname again"
                        Write-Host "2. Enter a different hostname"
                        Write-Host "3. Return to main menu"
                        $choice = Read-Host "Select an option (1-3)"
                        switch ($choice) {
                            '1' { continue }
                            '2' { $hostname = $null; continue }
                            '3' { $hostname = $null; continue 1 }
                            default { Write-Host "Invalid input. Returning to main menu." -ForegroundColor Yellow; $hostname = $null; continue 1 }
                        }
                    }
                } while (-not $pingSuccess)
            }

            # Sub-options for GPO-related actions
            Write-Host "`nGPO Actions:" -ForegroundColor Cyan
            Write-Host "  1. gpupdate /force"
            Write-Host "  2. Renew machine certificate (certutil -pulse) + verify"
            Write-Host "  3. Both (gpupdate then cert renew)"
            Write-Host "  4. Full Skynet Fix (gpupdate /force, gpupdate /force /sync, cert renew + verify)"
            Write-Host "  5. Request New Cert (certreq -enroll <TemplateName> /machine)"
            $gpChoice = Read-Host "Select (1-5) [Default: 1]"
            if ([string]::IsNullOrWhiteSpace($gpChoice)) { $gpChoice = '1' }

            switch ($gpChoice) {
                '2' {
                    # Cert renew only
                    $scriptBlock = {
                        Write-Host "Forcing certificate auto-enrollment (certutil -pulse)..." -ForegroundColor Cyan
                        certutil -pulse | Out-String | Write-Host
                        Write-Host "`nValid Client Authentication certificates (LocalMachine\\My):" -ForegroundColor Yellow
                        try {
                            Get-ChildItem Cert:\LocalMachine\My |
                              Where-Object { $_.EnhancedKeyUsageList.FriendlyName -match "Client Authentication" } |
                              Sort-Object NotAfter -Descending |
                              Select-Object Subject, NotBefore, NotAfter |
                              Format-Table -AutoSize
                        } catch {
                            Write-Host "Failed to enumerate machine cert store: $($_.Exception.Message)" -ForegroundColor Red
                        }
                    }
                }
                '3' {
                    # Both: gpupdate then cert renew
                    $scriptBlock = {
                        Write-Host "Running Group Policy Update..." -ForegroundColor Cyan
                        gpupdate /force | Out-String | Write-Host
                        Write-Host "Group Policy Update completed." -ForegroundColor Green

                        Write-Host "`nForcing certificate auto-enrollment (certutil -pulse)..." -ForegroundColor Cyan
                        certutil -pulse | Out-String | Write-Host
                        Write-Host "`nValid Client Authentication certificates (LocalMachine\\My):" -ForegroundColor Yellow
                        try {
                            Get-ChildItem Cert:\LocalMachine\My |
                              Where-Object { $_.EnhancedKeyUsageList.FriendlyName -match "Client Authentication" } |
                              Sort-Object NotAfter -Descending |
                              Select-Object Subject, NotBefore, NotAfter |
                              Format-Table -AutoSize
                        } catch {
                            Write-Host "Failed to enumerate machine cert store: $($_.Exception.Message)" -ForegroundColor Red
                        }
                    }
                }
                '4' {
                    $scriptBlock = {
                        Write-Host "Running Group Policy Update (/force)..." -ForegroundColor Cyan
                        gpupdate /force | Out-String | Write-Host
                        Write-Host "Group Policy Update (/force) completed." -ForegroundColor Green

                        # Ask whether to run synchronous Computer GPO and reboot now
                        $rebootAns = Read-Host "Run synchronous Computer GPO and reboot now? (Y/N) [Default: Y]"
                        if ([string]::IsNullOrWhiteSpace($rebootAns)) { $rebootAns = 'Y' }
                        if ($rebootAns.ToUpper().StartsWith('Y')) {
                            Write-Host "`nRunning Group Policy Update in SYNC mode with reboot..." -ForegroundColor Cyan
                            Write-Host "This will restart the machine to apply foreground Computer policy. The remote session will disconnect." -ForegroundColor Yellow
                            gpupdate /target:computer /force /sync /boot | Out-String | Write-Host
                            return
                        }

                        # No immediate reboot: proceed with SYNC without /boot, then renew certs
                        Write-Host "`nRunning Group Policy Update in SYNC mode (no reboot)..." -ForegroundColor Cyan
                        gpupdate /target:computer /force /sync | Out-String | Write-Host
                        Write-Host "Synchronous Group Policy application completed." -ForegroundColor Green

                        Write-Host "`nForcing certificate auto-enrollment (certutil -pulse)..." -ForegroundColor Cyan
                        certutil -pulse | Out-String | Write-Host

                        Write-Host "`nVerifying Client Authentication certificates (LocalMachine\\My):" -ForegroundColor Yellow
                        try {
                            $clientAuth = Get-ChildItem Cert:\LocalMachine\My | Where-Object {
                                $_.EnhancedKeyUsageList.FriendlyName -match "Client Authentication"
                            } | Sort-Object NotAfter -Descending

                            if ($clientAuth) {
                                $clientAuth | Select-Object Subject, NotBefore, NotAfter | Format-Table -AutoSize
                            } else {
                                Write-Host "No Client Authentication certs found." -ForegroundColor Red
                            }
                        } catch {
                            Write-Host "Failed to enumerate machine cert store: $($_.Exception.Message)" -ForegroundColor Red
                        }

                        Write-Host "`nFull Skynet Fix sequence completed (no reboot)." -ForegroundColor Green
                    }
                }
                '5' {
                    $tmpl = Read-Host "Enter certificate template name (e.g., 'Workstation Authentication'). Leave blank for 'Workstation Authentication'"
                    if ([string]::IsNullOrWhiteSpace($tmpl)) { $tmpl = 'Workstation Authentication' }
                    $templateToUse = $tmpl
                    $scriptBlock = {
                        param($Template)
                        Write-Host "Requesting new machine certificate using template: '$Template'..." -ForegroundColor Cyan
                        try {
                            certreq -enroll -q $Template /machine | Out-String | Write-Host
                        } catch {
                            Write-Host "certreq failed: $($_.Exception.Message)" -ForegroundColor Red
                        }
                        Write-Host "`nVerifying Authentication certs (LocalMachine\\My):" -ForegroundColor Yellow
                        try {
                            Get-ChildItem Cert:\LocalMachine\My |
                              Where-Object { $_.EnhancedKeyUsageList.FriendlyName -match "(Server|Client) Authentication" } |
                              Sort-Object NotAfter -Descending |
                              Select-Object Subject, Issuer, NotBefore, NotAfter, Thumbprint |
                              Format-Table -AutoSize
                        } catch {
                            Write-Host "Failed to enumerate cert store: $($_.Exception.Message)" -ForegroundColor Red
                        }
                    }
                    # Rebind with the template value for Invoke-Command -ArgumentList
                    $scriptArgs = @($templateToUse)
                }
                Default {
                    # gpupdate only
                    $scriptBlock = {
                        Write-Host "Running Group Policy Update..." -ForegroundColor Cyan
                        gpupdate /force
                        Write-Host "Group Policy Update completed." -ForegroundColor Green
                    }
                }
            }
        }
        # --- Option 7: Event Log Scan -------------------------------------------------
        # Prompts for hours back; queries System & Application logs and prints grouped
        # summaries by Level with a few sample entries.
        # -----------------------------------------------------------------------------
        '7' {
            if ($null -eq $hostname) {
                do {
                    $hostname = Read-Host "Enter the hostname or IP address of the Windows PC"
                    if ([string]::IsNullOrWhiteSpace($hostname)) {
                        Write-Host "Hostname cannot be empty." -ForegroundColor Red
                        continue
                    }
                    $pingSuccess = Test-Connection -ComputerName $hostname -Count 1 -Quiet
                    if (-not $pingSuccess) {
                        Write-Host "Device $hostname is not reachable via ping." -ForegroundColor Yellow
                        Write-Host "`nWhat would you like to do?"
                        Write-Host "1. Try the same hostname again"
                        Write-Host "2. Enter a different hostname"
                        Write-Host "3. Return to main menu"
                        $choice = Read-Host "Select an option (1-3)"
                        switch ($choice) {
                            '1' { continue }
                            '2' { $hostname = $null; continue }
                            '3' { $hostname = $null; continue 1 }
                            default { Write-Host "Invalid input. Returning to main menu." -ForegroundColor Yellow; $hostname = $null; continue 1 }
                        }
                    }
                } while (-not $pingSuccess)
            }
            $scriptBlock = {
                $hoursBack = Read-Host "How many hours back do you want to scan? (Default: 1)"
                if ([string]::IsNullOrWhiteSpace($hoursBack)) { $hoursBack = 1 }
                $startTime = (Get-Date).AddHours(-[int]$hoursBack)
                $logs = @("System", "Application")
                foreach ($log in $logs) {
                    Write-Host "`n===== $log Log =====" -ForegroundColor Yellow
                    try {
                        $events = Get-WinEvent -FilterHashtable @{LogName=$log; StartTime=$startTime} -MaxEvents 100 |
                                  Select-Object TimeCreated, Id, LevelDisplayName, ProviderName, Message

                        if ($events.Count -eq 0) {
                            Write-Host "No events found." -ForegroundColor DarkGray
                            continue
                        }

                        $grouped = $events | Group-Object LevelDisplayName
                        foreach ($group in $grouped) {
                            Write-Host "`n--- $($group.Name) Events ---" -ForegroundColor Cyan
                            foreach ($entry in $group.Group | Select-Object -First 5) {
                                Write-Host "[$($entry.TimeCreated)] [$($entry.Id)] $($entry.ProviderName)"
                                Write-Host "  $($entry.Message.Split("`n")[0])`n"
                            }
                        }
                    } catch {
                        Write-Host "Could not retrieve $log log: $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
            }
        }
        # --- Option 2: Windows Updates View ------------------------------------------
        # Shows installed hotfixes and generates WindowsUpdate.log for review.
        # -----------------------------------------------------------------------------
         '2' {
             # Prompt for hostname if it is explicitly null
             if ($null -eq $hostname) {
                 do {
                     $hostname = Read-Host "Enter the hostname or IP address of the Windows PC"
                     if ([string]::IsNullOrWhiteSpace($hostname)) {
                         Write-Host "Hostname cannot be empty." -ForegroundColor Red
                         continue
                     }
                     $pingSuccess = Test-Connection -ComputerName $hostname -Count 1 -Quiet
                     if (-not $pingSuccess) {
                         Write-Host "Device $hostname is not reachable via ping." -ForegroundColor Yellow
                         Write-Host "`nWhat would you like to do?"
                         Write-Host "1. Try the same hostname again"
                         Write-Host "2. Enter a different hostname"
                         Write-Host "3. Return to main menu"
                         $choice = Read-Host "Select an option (1-3)"
                         switch ($choice) {
                             '1' { continue }
                             '2' { $hostname = $null; continue }
                             '3' { $hostname = $null; continue 1 }
                             default { Write-Host "Invalid input. Returning to main menu." -ForegroundColor Yellow; $hostname = $null; continue 1 }
                         }
                     }
                 } while (-not $pingSuccess)
             }
             $scriptBlock = {
                 Write-Host "Viewing Installed and Failed Windows Updates..." -ForegroundColor Cyan
                 Write-Host "Installed Updates:" -ForegroundColor Yellow
                 Get-HotFix | Format-Table -AutoSize
                 Write-Host "`nWindows Update Log:" -ForegroundColor Yellow
                 Get-WindowsUpdateLog
                 Write-Host "Completed viewing updates." -ForegroundColor Green
             }
         }
        # --- Option 3: Computer Info --------------------------------------------------
        # Collects model, serial, OS, RAM, CPU, uptime, logged-in user, and monitor data
        # using CIM queries; includes a fallback for monitor info.
        # -----------------------------------------------------------------------------
         '3' {
             # Prompt for hostname if it is explicitly null
             if ($null -eq $hostname) {
                 do {
                     $hostname = Read-Host "Enter the hostname or IP address of the Windows PC"
                     if ([string]::IsNullOrWhiteSpace($hostname)) {
                         Write-Host "Hostname cannot be empty." -ForegroundColor Red
                         continue
                     }
                     $pingSuccess = Test-Connection -ComputerName $hostname -Count 1 -Quiet
                     if (-not $pingSuccess) {
                         Write-Host "Device $hostname is not reachable via ping." -ForegroundColor Yellow
                         Write-Host "`nWhat would you like to do?"
                         Write-Host "1. Try the same hostname again"
                         Write-Host "2. Enter a different hostname"
                         Write-Host "3. Return to main menu"
                         $choice = Read-Host "Select an option (1-3)"
                         switch ($choice) {
                             '1' { continue }
                             '2' { $hostname = $null; continue }
                             '3' { $hostname = $null; continue 1 }
                             default { Write-Host "Invalid input. Returning to main menu." -ForegroundColor Yellow; $hostname = $null; continue 1 }
                         }
                     }
                 } while (-not $pingSuccess)
             }
             $scriptBlock = {
                 function Convert-EdidChars {
                     param([uint16[]]$Chars)
                     if (-not $Chars) { return "" }
                     $bytes = @()
                     foreach ($c in $Chars) {
                         if ($c -eq 0) { break }
                         $bytes += [byte]$c
                     }
                     return -join ($bytes | ForEach-Object {[char]$_})
                 }

                 # --- Core system ---
                 $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
                 $bios           = Get-CimInstance -ClassName Win32_BIOS
                 $os             = Get-CimInstance -ClassName Win32_OperatingSystem
                 $ram            = Get-CimInstance -ClassName Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum
                 $totalRAM       = [math]::Round($ram.Sum / 1GB, 2)

                 $totalMem = [float]$os.TotalVisibleMemorySize
                 $freeMem  = [float]$os.FreePhysicalMemory
                 $usedMem  = $totalMem - $freeMem
                 $ramUtil  = if ($totalMem -gt 0) { [math]::Round(($usedMem / $totalMem) * 100, 2) } else { 0 }

                 $cpu           = Get-CimInstance -ClassName Win32_Processor | Select-Object -First 1
                 $cpuSpeedGHz   = if ($cpu) { [math]::Round($cpu.MaxClockSpeed / 1000, 2) } else { "N/A" }

                 if ($os.LastBootUpTime) {
                     $uptime = (Get-Date) - $os.LastBootUpTime
                     $uptimeFormatted = "{0} Days, {1} Hours, {2} Minutes" -f $uptime.Days, $uptime.Hours, $uptime.Minutes
                 } else {
                     $uptimeFormatted = "Unavailable"
                 }

                 # --- Active monitors in use (root\wmi) ---
                 try {
                     $monBasic = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorBasicDisplayParams -ErrorAction Stop | Where-Object { $_.Active -eq $true }
                     $monId    = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorID -ErrorAction Stop
                     $activeMons = @()

                     foreach ($m in $monBasic) {
                         $idMatch = $monId | Where-Object { $_.InstanceName -eq $m.InstanceName } | Select-Object -First 1
                         $friendly = if ($idMatch) { Convert-EdidChars $idMatch.UserFriendlyName } else { $null }
                         $mfg      = if ($idMatch) { Convert-EdidChars $idMatch.ManufacturerName } else { $null }
                         $serial   = if ($idMatch) { Convert-EdidChars $idMatch.SerialNumberID } else { $null }

                         $obj = [pscustomobject]@{
                             InstanceName = $m.InstanceName
                             FriendlyName = if ([string]::IsNullOrWhiteSpace($friendly)) { $null } else { $friendly }
                             Manufacturer = if ([string]::IsNullOrWhiteSpace($mfg)) { $null } else { $mfg }
                             Serial       = if ([string]::IsNullOrWhiteSpace($serial)) { $null } else { $serial }
                             Manufacture  = if ($m.WeekOfManufacture -gt 0 -and $m.YearOfManufacture -gt 0) { "W{0} {1}" -f $m.WeekOfManufacture, $m.YearOfManufacture } else { $null }
                             SizeCM       = if ($m.MaxHorizontalImageSize -and $m.MaxVerticalImageSize) { "{0}x{1}" -f $m.MaxHorizontalImageSize, $m.MaxVerticalImageSize } else { $null }
                         }
                         $activeMons += $obj
                     }
                 } catch {
                     $activeMons = @()
                 }

                 # --- Dock detection (heuristic over PnP entities) ---
                 $dockPatterns = '(?i)(dock|port replicator|usb[-\s]?c dock|thunderbolt\s*dock|wd1[59]|wd2[02]|tb16|k16a|ultra\s*dock|kensington\s*sd|thinkpad\s*dock|plugable|wavlink)'
                 try {
                     $dockDevices = Get-CimInstance Win32_PnPEntity -ErrorAction Stop |
                                    Where-Object { ($_.PNPClass -eq 'Dock') -or ($_.Name -match $dockPatterns) -or ($_.Description -match $dockPatterns) }
                 } catch {
                     $dockDevices = @()
                 }
                 $isDocked = if ($dockDevices -and $dockDevices.Count -gt 0) { "Yes" } else { "No or Unknown" }

                 # --- Primary network (Wi‑Fi vs Ethernet) ---
                 try {
                     $adaptersUp = Get-NetAdapter -Physical | Where-Object { $_.Status -eq 'Up' }
                     $ipcfg = Get-NetIPConfiguration | Where-Object { $_.IPv4DefaultGateway -ne $null } | Select-Object -First 1
                     $primaryIf = $null
                     if ($ipcfg) {
                         $primaryIf = $adaptersUp | Where-Object { $_.InterfaceIndex -eq $ipcfg.InterfaceIndex } | Select-Object -First 1
                     }
                     if (-not $primaryIf) { $primaryIf = $adaptersUp | Select-Object -First 1 }

                     if ($primaryIf) {
                         $isWifi = ($primaryIf.NdisPhysicalMedium -eq 9) -or ($primaryIf.InterfaceDescription -match '(?i)wi-?fi|wlan|802\.11')
                         $netType = if ($isWifi) { 'Wi-Fi' } else { 'Ethernet' }
                         $netOut  = "{0} ({1})" -f $netType, $primaryIf.InterfaceAlias
                         $ipAddr  = (Get-NetIPAddress -AddressFamily IPv4 -InterfaceIndex $primaryIf.InterfaceIndex -ErrorAction SilentlyContinue |
                                     Where-Object { $_.PrefixOrigin -ne 'WellKnown' } |
                                     Select-Object -ExpandProperty IPAddress -First 1)
                     } else {
                         $netOut = "Unknown"
                         $ipAddr = $null
                     }
                 } catch {
                     $netOut = "Unknown"
                     $ipAddr = $null
                 }

                 # --- Output ---
                 Write-Host "`n===== System Info ====="
                 Write-Host "Model:           $($computerSystem.Model)"
                 Write-Host "Serial Number:   $($bios.SerialNumber)"
                 Write-Host "OS:              $($os.Caption)"
                 Write-Host "OS Version:      $($os.Version)"
                 Write-Host "OS Build:        $($os.BuildNumber)"
                 Write-Host "Total RAM (GB):  $totalRAM"
                 Write-Host "RAM Used:        $ramUtil%"
                 Write-Host "CPU Speed (GHz): $cpuSpeedGHz"
                 Write-Host "Uptime:          $uptimeFormatted"
                 Write-Host "Logged-in User:  $($computerSystem.UserName)"

                 # Network
                 Write-Host "`n===== Network ====="
                 Write-Host "Primary Link:    $netOut"
                 if ($ipAddr) { Write-Host "IPv4 Address:    $ipAddr" }

                 # Dock
                 Write-Host "`n===== Dock ====="
                 Write-Host "Dock Detected:   $isDocked"
                 if ($dockDevices) {
                     $dockDevices | Select-Object -First 2 | ForEach-Object {
                         $dockName = if ($_.Name) { $_.Name } elseif ($_.Description) { $_.Description } else { '' }
                         Write-Host (" - {0}" -f $dockName)
                     }
                     if ($dockDevices.Count -gt 2) { Write-Host (" (+{0} more devices matched 'dock')" -f ($dockDevices.Count - 2)) }
                 }

                 # Monitors in use
                 Write-Host "`n===== Monitors In Use (Active) ====="
                 if (-not $activeMons -or $activeMons.Count -eq 0) {
                     Write-Host "No active external displays detected via WMI (root\wmi)."
                 } else {
                     $idx = 1
                     foreach ($m in $activeMons) {
                         $name = if ($m.FriendlyName) { $m.FriendlyName } elseif ($m.Manufacturer) { $m.Manufacturer } else { $m.InstanceName }
                         $serialOut = if ($m.Serial) { $m.Serial } else { 'n/a' }
                         $sizeOut   = if ($m.SizeCM) { $m.SizeCM } else { 'n/a' }
                         $mfgOut    = if ($m.Manufacture) { $m.Manufacture } else { '' }
                         Write-Host ("{0}. {1}  Serial: {2}  Size(cm): {3}  {4}" -f $idx, $name, $serialOut, $sizeOut, $mfgOut)
                         $idx++
                     }
                 }
             }
         }
        # --- Option 4: Dell Command Update -------------------------------------------
        # Executes dcu-cli.exe /applyUpdates and streams progress to the console.
        # -----------------------------------------------------------------------------
         '4' {
             # Prompt for hostname if it is explicitly null
             if ($null -eq $hostname) {
                 do {
                     $hostname = Read-Host "Enter the hostname or IP address of the Windows PC"
                     if ([string]::IsNullOrWhiteSpace($hostname)) {
                         Write-Host "Hostname cannot be empty." -ForegroundColor Red
                         continue
                     }
                     $pingSuccess = Test-Connection -ComputerName $hostname -Count 1 -Quiet
                     if (-not $pingSuccess) {
                         Write-Host "Device $hostname is not reachable via ping." -ForegroundColor Yellow
                         Write-Host "`nWhat would you like to do?"
                         Write-Host "1. Try the same hostname again"
                         Write-Host "2. Enter a different hostname"
                         Write-Host "3. Return to main menu"
                         $choice = Read-Host "Select an option (1-3)"
                         switch ($choice) {
                             '1' { continue }
                             '2' { $hostname = $null; continue }
                             '3' { $hostname = $null; continue 1 }
                             default { Write-Host "Invalid input. Returning to main menu." -ForegroundColor Yellow; $hostname = $null; continue 1 }
                         }
                     }
                 } while (-not $pingSuccess)
             }
             $scriptBlock = {
                 $dcuPath = "C:\Program Files (x86)\Dell\CommandUpdate\dcu-cli.exe"
                 if (-Not (Test-Path $dcuPath)) {
                     Write-Host "Dell Command Update CLI not found at $dcuPath" -ForegroundColor Red
                     return
                 }
 
                 Write-Host "Running Dell Command Update..." -ForegroundColor Cyan
 
                 $processInfo = New-Object System.Diagnostics.ProcessStartInfo
                 $processInfo.FileName = $dcuPath
                 $processInfo.Arguments = "/applyUpdates"
                 $processInfo.RedirectStandardOutput = $true
                 $processInfo.UseShellExecute = $false
                 $processInfo.CreateNoWindow = $true
 
                 $process = [System.Diagnostics.Process]::Start($processInfo)
                 $reader = $process.StandardOutput
 
                 while (-not $reader.EndOfStream) {
                     $line = $reader.ReadLine()
                     if ($line -match "Progress: (\d+)%") {
                         $percent = [int]$matches[1]
                         Write-Progress -Activity "Dell Updates" -Status "$percent% Complete" -PercentComplete $percent
                     } else {
                         Write-Host $line
                     }
                 }
 
                 $process.WaitForExit()
                 Write-Host "Dell updates completed with exit code: $($process.ExitCode)" -ForegroundColor Green
             }
         }
        # --- Option 5: One-Time Reboot -----------------------------------------------
        # Immediate reboot or schedules a one-shot SYSTEM reboot via schtasks.
        # -----------------------------------------------------------------------------
         '5' {
             # Prompt for hostname if it is explicitly null
             if ($null -eq $hostname) {
                 do {
                     $hostname = Read-Host "Enter the hostname or IP address of the Windows PC"
                     if ([string]::IsNullOrWhiteSpace($hostname)) {
                         Write-Host "Hostname cannot be empty." -ForegroundColor Red
                         continue
                     }
                     $pingSuccess = Test-Connection -ComputerName $hostname -Count 1 -Quiet
                     if (-not $pingSuccess) {
                         Write-Host "Device $hostname is not reachable via ping." -ForegroundColor Yellow
                         Write-Host "`nWhat would you like to do?"
                         Write-Host "1. Try the same hostname again"
                         Write-Host "2. Enter a different hostname"
                         Write-Host "3. Return to main menu"
                         $choice = Read-Host "Select an option (1-3)"
                         switch ($choice) {
                             '1' { continue }
                             '2' { $hostname = $null; continue }
                             '3' { $hostname = $null; continue 1 }
                             default { Write-Host "Invalid input. Returning to main menu." -ForegroundColor Yellow; $hostname = $null; continue 1 }
                         }
                     }
                 } while (-not $pingSuccess)
             }
             $scriptBlock = {
                 $rebootChoice = Read-Host "Do you want to reboot now or schedule it for later? (N/L)"
                 switch ($rebootChoice.ToUpper()) {
                     'N' {
                         Write-Host "Rebooting now. Goodbye!" -ForegroundColor Cyan
                         shutdown.exe /r /t 0
                         exit
                     }
                     'L' {
                         $validInput = $false
                         while (-not $validInput) {
                             try {
                                 $defaultDate = (Get-Date).ToString("MM/dd/yyyy")
                                 $defaultTime = "18:00"
                                 $rebootDate = Read-Host -Prompt "Enter the reboot date (MM/DD/YYYY) [Default: $defaultDate]"
                                 if ([string]::IsNullOrWhiteSpace($rebootDate)) { $rebootDate = $defaultDate }
                                 $rebootTime = Read-Host -Prompt "Enter the reboot time (HH:MM in 24-hour format) [Default: $defaultTime]"
                                 if ([string]::IsNullOrWhiteSpace($rebootTime)) { $rebootTime = $defaultTime }
 
                                 if ($rebootDate -notmatch "/") {
                                     $rebootDate = $rebootDate.Insert(2, "/").Insert(5, "/")
                                 }
                                 if ($rebootTime -notmatch ":") {
                                     $rebootTime = $rebootTime.Insert(2, ":")
                                 }
 
                                 $scheduledDateTime = [datetime]::ParseExact("$rebootDate $rebootTime", "MM/dd/yyyy HH:mm", $null)
 
                                 if ($scheduledDateTime -lt (Get-Date)) {
                                     Write-Host "Time is in the past. Enter a future date/time." -ForegroundColor Yellow
                                 } else {
                                     $validInput = $true
                                 }
                             } catch {
                                 Write-Host "Invalid format. Use MM/DD/YYYY and HH:MM (24-hour format)." -ForegroundColor Red
                             }
                         }
 
                         $formattedTime = $scheduledDateTime.ToString("HH:mm")
                         $formattedDate = $scheduledDateTime.ToString("MM/dd/yyyy")
                         $schtasksCommand = "schtasks /create /tn 'OneTimeReboot' /tr 'shutdown.exe /r /t 0' /sc once /sd $formattedDate /st $formattedTime /RU SYSTEM /F"
 
                         Invoke-Expression $schtasksCommand
                         Write-Host "Scheduled reboot at $scheduledDateTime." -ForegroundColor Green
                     }
                     default {
                         Write-Host "Invalid selection. Cancelling reboot." -ForegroundColor Yellow
                     }
                 }
             }
         }
        # --- Option 6: Hostname Reservation Assistant --------------------------------
        # Validates a 13-character base then POSTs to RXCRY01TECHLT01:8080 for the next
        # available device number and shows the response.
        # -----------------------------------------------------------------------------
        '6' {
            do {
                do {
                    $base = Read-Host "Enter the 13-character hostname base (e.g., RXCARY01OBGYN)"
                    if ($base.Length -ne 13 -or $base -notmatch '^[A-Z0-9]{13}$') {
                        Write-Host "`nInvalid input. Must be exactly 13 uppercase alphanumeric characters." -ForegroundColor Red
                        $base = $null
                    }
                } while (-not $base)

                $uri = "http://$HostnameApiHost`:$HostnameApiPort/"

                try {
                    $response = Invoke-RestMethod -Method POST -Uri $uri -Body $base
                    Write-Host "`n$response" -ForegroundColor Green
                } catch {
                    $errMsg = $_.Exception.Message

                    # Connection refused usually means: host reachable, but nothing is listening on the port.
                    $isRefused = $errMsg -match '(?i)refused'

                    if ($isRefused) {
                        Write-Status -Level WARN -Message "Hostname Reservation API refused connection at $HostnameApiHost`:$HostnameApiPort."
                        Write-Status -Level INFO -Message "Self-heal: starting '$HostnameApiTaskName' on $HostnameApiHost..."

                        # Try to start the API task remotely (requires WinRM to API host)
                        $startResult = Start-HostnameReservationApiRemote -ApiHost $HostnameApiHost -TaskName $HostnameApiTaskName -Credential $Cred
                        if ($startResult) { Write-Host $startResult }

                        # Give it a moment to bind the port
                        Start-Sleep -Seconds 3

                        if (Test-TcpPort -HostName $HostnameApiHost -Port $HostnameApiPort) {
                            Write-Status -Level OK -Message "API port is responding. Retrying request..."
                            try {
                                $response2 = Invoke-RestMethod -Method POST -Uri $uri -Body $base
                                Write-Host "`n$response2" -ForegroundColor Green
                            } catch {
                                Write-Status -Level ERROR -Message "Retry failed: $($_.Exception.Message)"
                            }
                        } else {
                            Write-Status -Level ERROR -Message "Self-heal failed: still not listening on $HostnameApiHost`:$HostnameApiPort."
                            Write-Status -Level WARN -Message "Check scheduled task '$HostnameApiTaskName' on $HostnameApiHost and its logs/output."
                        }
                    } else {
                        Write-Status -Level ERROR -Message "Hostname Reservation API error: $errMsg"
                    }
                }

                Write-Host "`nWhat would you like to do next?"
                Write-Host "1. Try another hostname"
                Write-Host "2. Return to main menu"
                $next = Read-Host "Select an option (1-2)"
                if ($next -eq '2') { break }
            } while ($true)
        }
        # --- Option 9: Battery Report ------------------------------------------------
        # Generates an HTML battery report on the remote device and saves it to C:\HCSTools.
        # If the device has no battery (desktop), prints an error and exits the option.
        # -----------------------------------------------------------------------------
        '9' {
            if ($null -eq $hostname) {
                do {
                    $hostname = Read-Host "Enter the hostname or IP address of the Windows PC"
                    if ([string]::IsNullOrWhiteSpace($hostname)) {
                        Write-Host "Hostname cannot be empty." -ForegroundColor Red
                        continue
                    }
                    $pingSuccess = Test-Connection -ComputerName $hostname -Count 1 -Quiet
                    if (-not $pingSuccess) {
                        Write-Host "Device $hostname is not reachable via ping." -ForegroundColor Yellow
                        Write-Host "`nWhat would you like to do?"
                        Write-Host "1. Try the same hostname again"
                        Write-Host "2. Enter a different hostname"
                        Write-Host "3. Return to main menu"
                        $choice = Read-Host "Select an option (1-3)"
                        switch ($choice) {
                            '1' { continue }
                            '2' { $hostname = $null; continue }
                            '3' { $hostname = $null; continue 1 }
                            default { Write-Host "Invalid input. Returning to main menu." -ForegroundColor Yellow; $hostname = $null; continue 1 }
                        }
                    }
                } while (-not $pingSuccess)
            }

            $scriptBlock = {
                # Detect battery presence
                try {
                    $bats = Get-CimInstance -ClassName Win32_Battery -ErrorAction Stop
                } catch {
                    $bats = $null
                }

                if (-not $bats) {
                    Write-Host "This device is not a laptop (no battery detected)." -ForegroundColor Red
                    "__NOT_LAPTOP__"
                    return
                }

                $outDir = 'C:\HCSTools'
                try {
                    if (-not (Test-Path $outDir)) {
                        New-Item -Path $outDir -ItemType Directory -Force | Out-Null
                    }
                } catch {
                    Write-Host "Failed to create/access $outDir : $($_.Exception.Message)" -ForegroundColor Red
                    return
                }

                $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
                $outFile = Join-Path $outDir ("BatteryReport_{0}.html" -f $stamp)

                try {
                    powercfg /batteryreport /output $outFile | Out-Null
                } catch {
                    Write-Host "Failed to generate battery report: $($_.Exception.Message)" -ForegroundColor Red
                    return
                }

                if (Test-Path $outFile) {
                    Write-Host "Battery report generated:" -ForegroundColor Green
                    Write-Host "   $outFile"
                    $outFile
                } else {
                    Write-Host "Battery report did not generate as expected." -ForegroundColor Red
                    "__REPORT_FAILED__"
                }
            }
        }
        '10' {
            if ($null -eq $hostname) {
                do {
                    $hostname = Read-Host "Enter the hostname or IP address of the Windows PC"
                    if ([string]::IsNullOrWhiteSpace($hostname)) {
                        Write-Host "Hostname cannot be empty." -ForegroundColor Red
                        continue
                    }
                    $pingSuccess = Test-Connection -ComputerName $hostname -Count 1 -Quiet
                    if (-not $pingSuccess) {
                        Write-Host "Device $hostname is not reachable via ping." -ForegroundColor Yellow
                        $hostname = $null
                        continue
                    }
                } while (-not $hostname)
            }

            Rexie-Robo -Hostname $hostname -Credential $Cred
        }
        default {
            Write-Host "Invalid selection. Please choose a valid menu option." -ForegroundColor Red
            continue
        }
     }
 
    # (no action here, moved logic above)

    # --- Remote Execution Orchestrator -------------------------------------------
    # For options that run on a remote host, this section:
    #   1) Verifies the host is reachable
    #   2) Validates WinRM connectivity with up to 3 credential retries
    #   3) Invokes the previously prepared $scriptBlock on the remote host
    # -----------------------------------------------------------------------------
    if ($selection -in @('1','2','3','4','5','7','9')) {
        Write-Status -Level INFO -Message "Selected option $selection"

        # Check if the remote machine is online
        if (-not (Test-Connection -ComputerName $hostname -Count 2 -Quiet)) {
            Write-Status -Level ERROR -Message "Host $hostname is offline or unreachable."
            do {
                Write-Host "`nWhat would you like to do?"
                Write-Host "1. Try a different hostname"
                Write-Host "2. Retry the same hostname"
                Write-Host "3. Exit to main menu"
                $offlineChoice = Read-Host "Select an option (1-3)"
                switch ($offlineChoice) {
                    '1' { $hostname = $null; continue 2 }
                    '2' { continue 2 }
                    '3' { $hostname = $null; continue 1 }
                    default { Write-Host "Invalid input. Please enter 1, 2, or 3." -ForegroundColor Yellow }
                }
            } while ($true)
        }

        ## Test remote WinRM connectivity with retry if credentials are bad
        $connectionTestSuccess = $false
        $maxAttempts = 3
        $attempt = 0
        while (-not $connectionTestSuccess -and $attempt -lt $maxAttempts) {
            try {
                Invoke-Command -ComputerName $hostname -Credential $Cred -ScriptBlock { Test-WSMan } -ErrorAction Stop
                Write-Status -Level OK -Message "WinRM connected to $hostname. Executing remote command."
                $connectionTestSuccess = $true
            } catch {
                Write-Status -Level ERROR -Message "WinRM connection failed to ${hostname}: $($_.Exception.Message)"
                $Cred = Get-Credential -Message "Enter credentials for remote access"
                $attempt++
                if ($attempt -ge $maxAttempts) {
                    Write-Status -Level ERROR -Message "Maximum WinRM connection attempts reached. Exiting."
                    exit 1
                }
            }
        }

        try {
            $invokeResult = $null
            if ($null -ne $scriptArgs) {
                $invokeResult = Invoke-Command -ComputerName $hostname -Credential $Cred -ScriptBlock $scriptBlock -ArgumentList $scriptArgs -ErrorAction Stop
            } else {
                $invokeResult = Invoke-Command -ComputerName $hostname -Credential $Cred -ScriptBlock $scriptBlock -ErrorAction Stop
            }
            $scriptArgs = $null

            # If Battery Report (Option 9), open admin share in Finder to C$\HCSTools when possible
            if ($selection -eq '9') {
                if ($invokeResult -contains '__NOT_LAPTOP__') {
                    Write-Host "Battery report not available: device appears to be a desktop (no battery)." -ForegroundColor Yellow
                } elseif ($invokeResult -contains '__REPORT_FAILED__') {
                    Write-Host "Battery report generation failed on the remote device." -ForegroundColor Red
                } else {
                    # Prefer opening directly to the HCSTools folder on the C$ admin share
                    $smbPath = "smb://$hostname/C$/HCSTools"
                    try {
                        Write-Status -Level INFO -Message "Opening admin share: $smbPath"
                        Start-Process $smbPath
                    } catch {
                        # Fallback: open the root of C$
                        $smbRoot = "smb://$hostname/C$"
                        Write-Status -Level WARN -Message "Could not open HCSTools directly. Opening: $smbRoot"
                        Start-Process $smbRoot
                    }
                }
            }
        } catch {
            Write-Status -Level ERROR -Message "Remote command failed: $($_.Exception.Message)"
        }
    }

    # --- Post-Task Prompt ---------------------------------------------------------
    # Lets the operator reuse the same hostname, switch hosts, or exit the session.
    # -----------------------------------------------------------------------------
    Write-Status -Level INFO -Message "What would you like to do next?"
    Write-Host "1. Run another task on this computer"
    Write-Host "2. Enter a new hostname"
    Write-Host "3. Exit"
    $nextAction = Read-Host "Select an option (1-3)"
    switch ($nextAction) {
        '1' { }  # continue with same hostname
        '2' { $hostname = $null }
        '3' { $repeatSession = $false }
        default { $repeatSession = $false }
    }
    if ($repeatSession -eq $false) {
        $hostname = $null
    }
} while ($repeatSession)
#endregion Session Loop & Credential Handling
        