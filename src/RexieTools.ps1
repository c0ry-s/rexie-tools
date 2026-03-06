<#
.SYNOPSIS
    Rexie Tools – Mac-hosted PowerShell remote admin launcher for UNC Health.
.DESCRIPTION
    Presents a menu of common remote administration tasks (WinRM-based). Stores a universal
    credential (optional) in the user's Documents folder and re-uses it across tasks.
.VERSION
    1.0.2
.AUTHOR
    c0ryS (Cory Smith)
.LAST UPDATED
    2026-02-26
.REQUIREMENTS
    • PowerShell 7+
    • Internet access for GitHub version check (if unavailable, script still runs)
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

# --- Core Helpers -------------------------------------------------------------
# Shared helpers for credentials, hostname prompting, WinRM validation, and
# remote execution. Several menu options depend on these.
# -----------------------------------------------------------------------------
function Pause-Rexie {
    [CmdletBinding()]
    param([string]$Message = "Press Enter to continue")
    Read-Host $Message | Out-Null
}

function Get-RexieCredential {
    [CmdletBinding()]
    param([string]$CredPath)

    if (Test-Path $CredPath) {
        try { return Import-Clixml -Path $CredPath }
        catch {
            Write-Status -Level WARN -Message "Stored credential exists but could not be read. Prompting for credentials."
        }
    }

    Write-Status -Level WARN -Message "No stored credentials found."
    $cred = Get-Credential -Message "Enter credentials for remote access"

    $storeAnswer = Read-Host "Do you want to store these credentials for future use? (Y/N)"
    if ($storeAnswer.ToUpper().StartsWith("Y")) {
        try {
            $cred | Export-Clixml -Path $CredPath
            Write-Status -Level OK -Message "Credentials stored at $CredPath"
        } catch {
            Write-Status -Level WARN -Message "Failed to store credentials: $($_.Exception.Message)"
        }
    }
    return $cred
}

function Read-RexieHostname {
    [CmdletBinding()]
    param(
        [string]$CurrentHostname,
        [string]$Prompt = "Enter the hostname or IP address of the Windows PC",
        [switch]$ForceNew
    )

    if (-not $ForceNew -and -not [string]::IsNullOrWhiteSpace($CurrentHostname)) {
        $reuse = Read-Host "Reuse current hostname '$CurrentHostname'? (Y/N) [Default: Y]"
        if ([string]::IsNullOrWhiteSpace($reuse) -or $reuse.ToUpper().StartsWith("Y")) {
            return $CurrentHostname
        }
    }

    while ($true) {
        $h = Read-Host $Prompt
        if ([string]::IsNullOrWhiteSpace($h)) {
            Write-Host "Hostname cannot be empty." -ForegroundColor Red
            continue
        }

        $pingSuccess = Test-Connection -ComputerName $h -Count 1 -Quiet
        if ($pingSuccess) { return $h }

        Write-Host "Device $h is not reachable via ping." -ForegroundColor Yellow
        Write-Host "`nWhat would you like to do?"
        Write-Host "1. Try the same hostname again"
        Write-Host "2. Enter a different hostname"
        Write-Host "3. Return to main menu"
        $choice = Read-Host "Select an option (1-3)"
        switch ($choice) {
            '1' { continue }
            '2' { continue }
            '3' { return $null }
            default { Write-Host "Invalid input. Returning to main menu." -ForegroundColor Yellow; return $null }
        }
    }
}

function Test-RexieWinRM {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Hostname,
        [Parameter(Mandatory)] [pscredential]$Credential,
        [int]$MaxAttempts = 3
    )

    $attempt = 0
    while ($attempt -lt $MaxAttempts) {
        try {
            Invoke-Command -ComputerName $Hostname -Credential $Credential -ScriptBlock { Test-WSMan } -ErrorAction Stop | Out-Null
            return $true
        } catch {
            Write-Status -Level ERROR -Message "WinRM connection failed to ${Hostname}: $($_.Exception.Message)"
            $Credential = Get-Credential -Message "Enter credentials for remote access"
            $script:Cred = $Credential
            $attempt++
        }
    }
    return $false
}

function Invoke-RexieRemote {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Hostname,
        [Parameter(Mandatory)] [pscredential]$Credential,
        [Parameter(Mandatory)] [scriptblock]$ScriptBlock,
        [object[]]$ArgumentList
    )

    Write-Status -Level INFO -Message "Selected remote target: $Hostname"

    if (-not (Test-Connection -ComputerName $Hostname -Count 2 -Quiet)) {
        Write-Status -Level ERROR -Message "Host $Hostname is offline or unreachable."
        return $null
    }

    if (-not (Test-RexieWinRM -Hostname $Hostname -Credential $Credential)) {
        Write-Status -Level ERROR -Message "Maximum WinRM connection attempts reached. Returning to menu."
        return $null
    }

    try {
        if ($ArgumentList) {
            return Invoke-Command -ComputerName $Hostname -Credential $Credential -ScriptBlock $ScriptBlock -ArgumentList $ArgumentList -ErrorAction Stop
        } else {
            return Invoke-Command -ComputerName $Hostname -Credential $Credential -ScriptBlock $ScriptBlock -ErrorAction Stop
        }
    } catch {
        Write-Status -Level ERROR -Message "Remote command failed: $($_.Exception.Message)"
        return $null
    }
}

# --- ASCII Splash: Login Shark -------------------------------------------------
function Show-LoginSharkSplash {
    [CmdletBinding()]
    param()

    Clear-Host
    Write-Host @"

                            ,_.
                           ./  |                                          _-
                         ./    |                                       _-'/
      ______.,         ./      /                                     .'  (
 _---'___._.  '----___/       (                                    ./  /`'
(,----,_  O \                  \_.                               ./   :
 \___   "--_                      "--._,                       ./    /
 /^^^^^-__          ,,,,,               "-._       /|         /     /
 `,       -        /////                    "`--__/ (_,    ,_/    ./
   "-_,           ''''' __,                            `--'      /
       "-_,             \\ `-_                                  (
           "-_.          \\   \.                                 \_
          /    "--__,      \\   \.                       ____.     "-._,
         /        ./ `---____\\   \.______________,---\ (     \,        "-.,
        |       ./             \\   \        /\  |     \|       `--______---`
        |     ./                 \\  \      /_/\_!
        |   ./                     \\ \
        |  /     *:Login SHARK:*     \_\
        |_/
"@ -ForegroundColor Cyan
}


# --- Login Shark (Session Intelligence) ---------------------------------------
# Shows active session info, last logged-on user, recent interactive auth events,
# lock state, and a reboot / remote login recommendation.
# -------------------------------------------------------------------------------
function Rexie-LoginShark {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Hostname,
        [Parameter(Mandatory)] [pscredential]$Credential
    )

    Write-Status -Level INFO -Message "Login Shark scanning $Hostname ..."

    # 1) Active sessions (LOGON TIME / IDLE TIME)
    $quserOut = $null
    try {
        $quserOut = Invoke-Command -ComputerName $Hostname -Credential $Credential -ScriptBlock { quser } -ErrorAction Stop
    } catch {
        $quserOut = $null
    }

    # 2) Last logged-on user (registry)
    $lastLoggedOn = $null
    try {
        $lastLoggedOn = Invoke-Command -ComputerName $Hostname -Credential $Credential -ScriptBlock {
            (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI" -ErrorAction SilentlyContinue).LastLoggedOnUser
        } -ErrorAction Stop
    } catch {
        $lastLoggedOn = $null
    }

    # 3) Current interactive user (best-effort): owner of explorer.exe
    $interactiveUser = $null
    try {
        $interactiveUser = Invoke-Command -ComputerName $Hostname -Credential $Credential -ScriptBlock {
            try {
                $p = Get-Process explorer -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($p) {
                    $o = $p.GetOwner()
                    if ($o -and $o.User) { return $o.User }
                }
            } catch { }
            return $null
        } -ErrorAction Stop
    } catch {
        $interactiveUser = $null
    }

    # 4) Recent auth events (Security log): last interactive logon + last lock/unlock
    $lastInteractiveLogon = $null
    $lastLockEvent        = $null
    $lastUnlockEvent      = $null

    try {
        $auth = Invoke-Command -ComputerName $Hostname -Credential $Credential -ScriptBlock {
            $out = [ordered]@{
                LastInteractiveLogon = $null
                LastLock             = $null
                LastUnlock           = $null
            }

            try {
                $ev = Get-WinEvent -FilterHashtable @{ LogName='Security'; Id=4624; StartTime=(Get-Date).AddDays(-2) } -MaxEvents 200 -ErrorAction Stop
                foreach ($e in $ev) {
                    $user = $e.Properties[5].Value
                    $dom  = $e.Properties[6].Value
                    $lt   = [int]$e.Properties[8].Value

                    if ([string]::IsNullOrWhiteSpace($user)) { continue }
                    if ($user -match '\$$') { continue }
                    if ($user -in @('SYSTEM','LOCAL SERVICE','NETWORK SERVICE','ANONYMOUS LOGON')) { continue }
                    if ($lt -notin 2,7,11) { continue }

                    $out.LastInteractiveLogon = [pscustomobject]@{
                        Time      = $e.TimeCreated
                        User      = if ($dom) { "$dom\$user" } else { $user }
                        LogonType = $lt
                    }
                    break
                }
            } catch { }

            try {
                $lu = Get-WinEvent -FilterHashtable @{ LogName='Security'; Id=4800,4801; StartTime=(Get-Date).AddDays(-2) } -MaxEvents 50 -ErrorAction Stop
                foreach ($e in $lu) {
                    if ($e.Id -eq 4800 -and -not $out.LastLock) {
                        $out.LastLock = [pscustomobject]@{ Time = $e.TimeCreated }
                    }
                    if ($e.Id -eq 4801 -and -not $out.LastUnlock) {
                        $out.LastUnlock = [pscustomobject]@{ Time = $e.TimeCreated }
                    }
                    if ($out.LastLock -and $out.LastUnlock) { break }
                }
            } catch { }

            return [pscustomobject]$out
        } -ErrorAction Stop

        $lastInteractiveLogon = $auth.LastInteractiveLogon
        $lastLockEvent        = $auth.LastLock
        $lastUnlockEvent      = $auth.LastUnlock
    } catch {
        $lastInteractiveLogon = $null
        $lastLockEvent        = $null
        $lastUnlockEvent      = $null
    }

    # Parse quser output to detect active interactive session
    $hasInteractiveUser = $false
    $activeUser = $null
    if ($quserOut) {
        $activeSessionLines = $quserOut | Select-Object -Skip 1 | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        foreach ($line in $activeSessionLines) {
            if ($line -match '\s+Active\s+') {
                $hasInteractiveUser = $true
                $tokens = ($line -split '\s+') | Where-Object { $_ -ne '' }
                if ($tokens -and $tokens.Count -gt 0) { $activeUser = $tokens[0] }
                break
            }
        }
    }

    # Best-effort lock state
    $lockState = 'Unknown'
    if ($lastLockEvent -and $lastUnlockEvent) {
        $lockState = if ($lastLockEvent.Time -gt $lastUnlockEvent.Time) { 'Locked' } else { 'Unlocked' }
    } elseif ($lastLockEvent -and -not $lastUnlockEvent) {
        $lockState = 'Locked (no unlock seen)'
    } elseif ($lastUnlockEvent -and -not $lastLockEvent) {
        $lockState = 'Unlocked (no lock seen)'
    }

    Clear-Host
    Write-Host "=================================================" -ForegroundColor Cyan
    Write-Host "              Rexie Tools - Login Shark" -ForegroundColor Cyan
    Write-Host "=================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Hostname: $Hostname"
    Write-Host "Status:   Online (WinRM reachable)"
    if ($interactiveUser) { Write-Host ("User:     {0}" -f $interactiveUser) }
    if ($lockState)       { Write-Host ("Lock:     {0}" -f $lockState) }
    Write-Host ""

    Write-Host "-------------------------------------------------"
    Write-Host "Active Session (quser)"
    Write-Host "-------------------------------------------------"
    if ($quserOut) {
        $quserOut | ForEach-Object { Write-Host $_ }
    } else {
        Write-Host "Unavailable (quser failed)."
    }

    Write-Host ""
    Write-Host "-------------------------------------------------"
    Write-Host "Last Logged-On User"
    Write-Host "-------------------------------------------------"
    if ($lastLoggedOn) {
        Write-Host $lastLoggedOn
    } else {
        Write-Host "Unavailable"
    }

    Write-Host ""
    Write-Host "-------------------------------------------------"
    Write-Host "Recent Authentication (best-effort)"
    Write-Host "-------------------------------------------------"
    if ($lastInteractiveLogon) {
        Write-Host ("Last interactive logon: {0}  ({1})" -f $lastInteractiveLogon.Time, $lastInteractiveLogon.User)
        Write-Host ("LogonType: {0}" -f $lastInteractiveLogon.LogonType)
    } else {
        Write-Host "Last interactive logon: Unavailable (no access or not found in last 48h)"
    }

    if ($lastLockEvent)   { Write-Host ("Last lock:   {0}" -f $lastLockEvent.Time) }
    else                  { Write-Host "Last lock:   Unavailable" }

    if ($lastUnlockEvent) { Write-Host ("Last unlock: {0}" -f $lastUnlockEvent.Time) }
    else                  { Write-Host "Last unlock: Unavailable" }

    Write-Host ""
    Write-Host "-------------------------------------------------"
    Write-Host "System Recommendation"
    Write-Host "-------------------------------------------------"
    if ($hasInteractiveUser) {
        Write-Host "⚠ User session active" -ForegroundColor Yellow
        if ($activeUser) { Write-Host ("User: {0}" -f $activeUser) -ForegroundColor Yellow }

        if ($lockState -like 'Locked*') {
            Write-Host "Session appears LOCKED. Reboot is still risky, but less disruptive than an active unlocked session." -ForegroundColor Yellow
        }

        Write-Host "Not safe for reboot" -ForegroundColor Yellow
        Write-Host "Remote login possible (may interrupt user)" -ForegroundColor Yellow
    } else {
        if ($lockState -like 'Locked*' -or $lockState -like 'Unlocked*') {
            Write-Host "⚠ No ACTIVE quser session detected, but lock/unlock evidence suggests a user was recently present." -ForegroundColor Yellow
            Write-Host "Reboot is probably safe, but proceed with caution." -ForegroundColor Yellow
        } else {
            Write-Host "✓ Safe for reboot" -ForegroundColor Green
            Write-Host "✓ Safe for remote login" -ForegroundColor Green
        }
    }

    Write-Host ""
    Read-Host "Press Enter to return to Rexie Tools" | Out-Null
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
$currentVersion = [version]"1.0.2"

# TODO: Break repeated code blocks into reusable functions for maintainability.

#region Version Check & Banner (GitHub)
# GitHub-based version check (personal repo)
$GitHubOwner  = "c0ry-s"
$GitHubRepo   = "rexie-tools"
$GitHubBranch = "main"

$versionUrl = "https://raw.githubusercontent.com/$GitHubOwner/$GitHubRepo/$GitHubBranch/version.txt"

Write-Status -Level INFO -Message "-=*Rexie Tools by c0ryS*=-"
Write-Host @"
            __
           / _) 
    .-^^^-/ / 
 __/       /  
<__.|_|-|_|   
"@ -ForegroundColor Blue

try {
    $latestVersionString = (Invoke-RestMethod -Uri $versionUrl -Method Get -TimeoutSec 5 -ErrorAction Stop).ToString().Trim()

    if ([string]::IsNullOrWhiteSpace($latestVersionString)) {
        Write-Status -Level WARN -Message "GitHub version file is empty. Skipping version check."
    } else {
        $latestVersion = [version]$latestVersionString
        Write-Status -Level INFO -Message "Current version: $currentVersion. Latest: $latestVersion."

        if ($latestVersion -gt $currentVersion) {
            $releaseUrl = "https://github.com/$GitHubOwner/$GitHubRepo/releases/latest"
            Write-Status -Level WARN -Message "Update available ($latestVersion). Download latest from: $releaseUrl"
        } else {
            Write-Status -Level OK -Message "Script is up to date."
        }
    }
}
catch {
    Write-Status -Level WARN -Message "GitHub version check failed: $($_.Exception.Message)"
}
#endregion Version Check & Banner (GitHub)
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
    Write-Host "2. Event Log Scan"
    Write-Host "3. View Computer Info"
    Write-Host "4. Run Dell Command Update"
    Write-Host "5. Schedule One-Time Reboot"
    Write-Host "6. Hostname Reservation Assistant"
    Write-Host "7. Login Shark"
    Write-Host "8. SCCM / Software Center Actions"
    Write-Host "9. Battery Report"
    $selection = Read-Host "Enter your choice (1-9, Q to exit)"

    if ($selection -match '^[Qq]$') {
        break
    } elseif ($selection -in @('6','7')) {
        $hostname = $null
    }

    switch ($selection) {
        # --- Option 1: Group Policy Update -------------------------------------------
        '1' {
            $hostname = Read-RexieHostname -CurrentHostname $hostname
            if (-not $hostname) { break }

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
        # --- Option 2: Event Log Scan ------------------------------------------------
        # Prompts for hours back; queries System & Application logs and prints grouped
        # summaries by Level with a few sample entries.
        # -----------------------------------------------------------------------------
         '2' {
            $hostname = Read-RexieHostname -CurrentHostname $hostname
            if (-not $hostname) { break }
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
        # --- Option 3: Computer Info --------------------------------------------------
        # Collects model, serial, OS, RAM, CPU, uptime, logged-in user, and monitor data
        # using CIM queries; includes a fallback for monitor info.
        # -----------------------------------------------------------------------------
         '3' {
            $hostname = Read-RexieHostname -CurrentHostname $hostname
            if (-not $hostname) { break }
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
            $hostname = Read-RexieHostname -CurrentHostname $hostname
            if (-not $hostname) { break }
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
            $hostname = Read-RexieHostname -CurrentHostname $hostname
            if (-not $hostname) { break }
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
        # --- Option 7: Login Shark ---------------------------------------------------
        # Session intelligence: active user, logon evidence, lock state, and reboot risk.
        # -----------------------------------------------------------------------------
        '7' {
            Show-LoginSharkSplash
            $hostname = Read-RexieHostname -CurrentHostname $hostname
            if (-not $hostname) { break }

            if (-not (Test-RexieWinRM -Hostname $hostname -Credential $Cred)) {
                Write-Status -Level ERROR -Message "Maximum WinRM connection attempts reached. Returning to menu."
                break
            }

            Rexie-LoginShark -Hostname $hostname -Credential $Cred
            break
        }
        # --- Option 8: SCCM / Software Center Actions -------------------------------
        # Triggers common Configuration Manager client cycles remotely.
        # -----------------------------------------------------------------------------
        '8' {
            Write-Status -Level INFO -Message "Selected option 8 - SCCM / Software Center Actions"
            $hostname = Read-RexieHostname -CurrentHostname $hostname
            if (-not $hostname) { break }

            Write-Host "`nSCCM Client Actions:" -ForegroundColor Cyan
            Write-Host "1. Machine Policy Retrieval"
            Write-Host "2. User Policy Retrieval"
            Write-Host "3. Application Deployment Evaluation"
            Write-Host "4. Software Update Evaluation"
            Write-Host "5. Hardware Inventory"
            Write-Host "6. Full Client Check-In (all cycles)"

            $sccmChoice = Read-Host "Select an option (1-6)"
            $sccmActionName = switch ($sccmChoice) {
                '1' { 'Machine Policy Retrieval' }
                '2' { 'User Policy Retrieval' }
                '3' { 'Application Deployment Evaluation' }
                '4' { 'Software Update Evaluation' }
                '5' { 'Hardware Inventory' }
                '6' { 'Full Client Check-In (all cycles)' }
                default { 'Invalid / Unknown SCCM action' }
            }
            Write-Status -Level INFO -Message "Selected SCCM action: $sccmActionName"

            $scriptBlock = {
                param($choice)

                $ErrorActionPreference = 'Stop'

                function Invoke-Cycle($guid,$name){
                    try {
                        $result = Invoke-WmiMethod -Namespace root\ccm -Class SMS_Client -Name TriggerSchedule -ArgumentList $guid -ErrorAction Stop
                        $returnCode = $null
                        if ($null -ne $result -and $result.PSObject.Properties.Name -contains 'ReturnValue') {
                            $returnCode = [int]$result.ReturnValue
                        }

                        if ($null -eq $returnCode -or $returnCode -eq 0) {
                            Write-Host "$name triggered successfully." -ForegroundColor Green
                        }
                        else {
                            Write-Host "$name returned non-zero code: $returnCode" -ForegroundColor Yellow
                        }
                    }
                    catch {
                        $msg = $_.Exception.Message
                        Write-Host "Failed to trigger ${name}: $msg" -ForegroundColor Red
                    }
                }

                switch($choice){
                    '1' { Invoke-Cycle '{00000000-0000-0000-0000-000000000021}' 'Machine Policy Retrieval' }
                    '2' { Invoke-Cycle '{00000000-0000-0000-0000-000000000027}' 'User Policy Retrieval' }
                    '3' { Invoke-Cycle '{00000000-0000-0000-0000-000000000121}' 'Application Deployment Evaluation' }
                    '4' { Invoke-Cycle '{00000000-0000-0000-0000-000000000108}' 'Software Update Evaluation' }
                    '5' { Invoke-Cycle '{00000000-0000-0000-0000-000000000001}' 'Hardware Inventory' }
                    '6' {
                        Invoke-Cycle '{00000000-0000-0000-0000-000000000021}' 'Machine Policy Retrieval'
                        Invoke-Cycle '{00000000-0000-0000-0000-000000000022}' 'Machine Policy Evaluation'
                        Invoke-Cycle '{00000000-0000-0000-0000-000000000027}' 'User Policy Retrieval'
                        Invoke-Cycle '{00000000-0000-0000-0000-000000000028}' 'User Policy Evaluation'
                        Invoke-Cycle '{00000000-0000-0000-0000-000000000121}' 'Application Deployment Evaluation'
                        Invoke-Cycle '{00000000-0000-0000-0000-000000000108}' 'Software Update Evaluation'
                        Invoke-Cycle '{00000000-0000-0000-0000-000000000001}' 'Hardware Inventory'
                        Invoke-Cycle '{00000000-0000-0000-0000-000000000002}' 'Software Inventory'
                        Invoke-Cycle '{00000000-0000-0000-0000-000000000003}' 'Discovery Data Collection'
                    }
                    default { Write-Host "Invalid selection." -ForegroundColor Yellow }
                }
            }

            $scriptArgs = @($sccmChoice)
        }
        # --- Option 9: Battery Report ------------------------------------------------
        # Generates an HTML battery report on the remote device and saves it to C:\HCSTools.
        # If the device has no battery (desktop), prints an error and exits the option.
        # -----------------------------------------------------------------------------
        '9' {
            $hostname = Read-RexieHostname -CurrentHostname $hostname
            if (-not $hostname) { break }

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
    if ($selection -in @('1','2','3','4','5','8','9') -and $scriptBlock) {
        Write-Status -Level INFO -Message "Executing option $selection on $hostname ..."

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
        