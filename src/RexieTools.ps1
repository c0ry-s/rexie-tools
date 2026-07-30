<#
.SYNOPSIS
    Rexie Tools – Mac-hosted PowerShell remote admin launcher for UNC Health.
.DESCRIPTION
    Presents a menu of common remote administration tasks (WinRM-based). Stores a universal
    credential (optional) in the user's Documents folder and re-uses it across tasks.
.VERSION
    1.3.0
.AUTHOR
    c0ryS (Cory Smith)
.LAST UPDATED
    2026-07-30
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
# Shared helpers for hostname prompting, validated menu input, and WinRM
# connectivity checks. Several menu options depend on these.
# -----------------------------------------------------------------------------
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
        $choice = Read-RexieChoice -Prompt "Select an option (1-3)" -ValidChoices '1','2','3'
        if (-not $choice -or $choice -eq '3') { return $null }
        # '1' and '2' both just loop back to re-prompt for a hostname
    }
}

function Read-RexieChoice {
    # Returns $null if the user cancels with 'Q' - callers should treat that as
    # "back out to the main menu" rather than a valid selection.
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Prompt,
        [Parameter(Mandatory)] [string[]]$ValidChoices,
        [string]$Default
    )

    while ($true) {
        $promptText = if ($Default) { "$Prompt [Default: $Default] (Q to cancel)" } else { "$Prompt (Q to cancel)" }
        $choice = (Read-Host $promptText).ToUpper()

        if ($choice -eq 'Q') { return $null }
        if ([string]::IsNullOrWhiteSpace($choice) -and $Default) { return $Default }
        if ($choice -in $ValidChoices) { return $choice }

        Write-Host "Invalid selection. Please enter one of: $($ValidChoices -join ', '), or Q to cancel." -ForegroundColor Yellow
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
function Show-LoginShark {
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
                $proc = Get-CimInstance -ClassName Win32_Process -Filter "Name='explorer.exe'" -ErrorAction Stop | Select-Object -First 1
                if ($proc) {
                    $ownerInfo = Invoke-CimMethod -InputObject $proc -MethodName GetOwner -ErrorAction Stop
                    if ($ownerInfo -and $ownerInfo.ReturnValue -eq 0 -and $ownerInfo.User) {
                        return $ownerInfo.User
                    }
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
        Write-Host "[WARN] User session active" -ForegroundColor Yellow
        if ($activeUser) { Write-Host ("User: {0}" -f $activeUser) -ForegroundColor Yellow }

        if ($lockState -like 'Locked*') {
            Write-Host "Session appears LOCKED. Reboot is still risky, but less disruptive than an active unlocked session." -ForegroundColor Yellow
        }

        Write-Host "Not safe for reboot" -ForegroundColor Yellow
        Write-Host "Remote login possible (may interrupt user)" -ForegroundColor Yellow
    } else {
        if ($lockState -like 'Locked*' -or $lockState -like 'Unlocked*') {
            Write-Host "[WARN] No ACTIVE quser session detected, but lock/unlock evidence suggests a user was recently present." -ForegroundColor Yellow
            Write-Host "Reboot is probably safe, but proceed with caution." -ForegroundColor Yellow
        } else {
            Write-Host "[OK] Safe for reboot" -ForegroundColor Green
            Write-Host "[OK] Safe for remote login" -ForegroundColor Green
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
$currentVersion = [version]"1.3.0"

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
:sessionLoop do {

# Reset per-iteration state so an aborted option can never reuse a stale scriptblock/args
# from a previous selection.
$scriptBlock  = $null
$scriptArgs   = $null
$justShutDown = $false

# --- Credential Handling ------------------------------------------------------
# Loads a stored credential from ~/Documents/UniversalCredential.xml when present.
# If not present, prompts once and optionally persists for future runs.
# -----------------------------------------------------------------------------
$credPath = Join-Path ([Environment]::GetFolderPath('MyDocuments')) "UniversalCredential.xml"

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

# --- Main Menu ---------------------------------------------------------------
    Write-Status -Level INFO -Message "Select an option:"
    Write-Host "1. Group Policy (GPO)"
    Write-Host "2. Event Log Scan"
    Write-Host "3. View Computer Info"
    Write-Host "4. Run Dell Command Update"
    Write-Host "5. Reboot / Shutdown (Now or Scheduled)"
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

            $scriptBlock = {
                Write-Host "Running Group Policy Update..." -ForegroundColor Cyan
                gpupdate /force | Out-String | Write-Host
                Write-Host "Group Policy Update completed." -ForegroundColor Green
            }
        }
        # --- Option 2: Event Log Scan ------------------------------------------------
        # Prompts for hours back; queries System & Application logs for Warning/Error/
        # Critical events and prints grouped summaries by Level with a few sample entries.
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
                        $events = Get-WinEvent -FilterHashtable @{LogName=$log; StartTime=$startTime; Level=1,2,3} -MaxEvents 100 |
                                  Select-Object TimeCreated, Id, LevelDisplayName, ProviderName, Message

                        if ($events.Count -eq 0) {
                            Write-Host "No events found." -ForegroundColor DarkGray
                            continue
                        }

                        $grouped = $events | Group-Object LevelDisplayName
                        foreach ($group in $grouped) {
                            Write-Host "`n--- $($group.Name) Events ---" -ForegroundColor Cyan
                            $distinctEvents = $group.Group | Group-Object Id, ProviderName, { $_.Message.Split("`n")[0] } | Sort-Object { $_.Group[0].TimeCreated } -Descending
                            foreach ($dup in $distinctEvents | Select-Object -First 5) {
                                $entry = $dup.Group[0]
                                $countSuffix = if ($dup.Count -gt 1) { " (x$($dup.Count))" } else { "" }
                                Write-Host "[$($entry.TimeCreated)] [$($entry.Id)] $($entry.ProviderName)$countSuffix"
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
        # Collects model, serial, OS, RAM, CPU, uptime, logged-in user, BitLocker
        # encryption status, primary network link, dock detection, and active
        # monitor info via CIM/WMI queries.
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
                    $dockDevices = @(Get-CimInstance Win32_PnPEntity -ErrorAction Stop |
                                   Where-Object { ($_.PNPClass -eq 'Dock') -or ($_.Name -match $dockPatterns) -or ($_.Description -match $dockPatterns) })
                } catch {
                    $dockDevices = @()
                }
                $isDocked = if ($dockDevices.Count -gt 0) { "Yes" } else { "No or Unknown" }

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

                # --- Encryption / BitLocker Status ---
                $bitlockerStatus = "Unknown"
                $bitlockerColor  = "Yellow"

                try {
                    $blv = Get-BitLockerVolume -MountPoint $env:SystemDrive -ErrorAction Stop

                    if ($blv.ProtectionStatus -eq 'On') {
                        $bitlockerStatus = "Encrypted"
                        $bitlockerColor  = "Green"
                    }
                    else {
                        $bitlockerStatus = "NOT Encrypted"
                        $bitlockerColor  = "Red"
                    }
                }
                catch {
                    try {
                        # Fallback for systems without BitLocker module
                        $manageBde = manage-bde -status $env:SystemDrive 2>$null

                        if ($manageBde -match "Protection Status:\s+Protection On") {
                            $bitlockerStatus = "Encrypted"
                            $bitlockerColor  = "Green"
                        }
                        elseif ($manageBde -match "Protection Status:\s+Protection Off") {
                            $bitlockerStatus = "NOT Encrypted"
                            $bitlockerColor  = "Red"
                        }
                    }
                    catch {
                        $bitlockerStatus = "Unable to determine"
                        $bitlockerColor  = "Yellow"
                    }
                }
                # --- Output ---
                Write-Host "`n===== System Info ====="
                Write-Host "Model:           $($computerSystem.Model)"
                Write-Host "Serial Number:   $($bios.SerialNumber)"
                Write-Host "OS:              $($os.Caption)"
                Write-Host "OS Version:      $($os.Version)"
                Write-Host "OS Build:        $($os.BuildNumber)"
                Write-Host "Encryption:      $bitlockerStatus" -ForegroundColor $bitlockerColor
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
        # Executes dcu-cli.exe /applyUpdates, streams live progress to the console,
        # and reports a human-readable meaning for the exit code (see Dell's CLI
        # Reference Guide for the full code list).
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

                # Source: Dell Command | Update CLI Reference Guide - Exit/Error Codes
                $dcuExitCodes = @{
                    0    = 'Success'
                    1    = 'Success - reboot required'
                    2    = 'Unknown application error'
                    3    = 'System manufacturer is not Dell'
                    4    = 'CLI was not launched with administrative privilege'
                    5    = 'A reboot was pending from a previous operation'
                    6    = 'Another instance of Dell Command Update is already running'
                    7    = 'This system model is not supported'
                    8    = 'No update filters have been applied or configured'
                    500  = 'No updates were found for the system'
                    501  = 'An error occurred while determining the available updates'
                    502  = 'The operation was cancelled'
                    503  = 'An error occurred while downloading a file during the scan'
                    1000 = 'An error occurred retrieving the result of the apply updates operation'
                    1001 = 'The operation was cancelled'
                    1002 = 'An error occurred while downloading a file during the apply updates operation'
                }

                $exitCode = $process.ExitCode
                $successCodes = @(0)
                $benignCodes  = @(1, 5, 500, 502, 1001)

                $description = $dcuExitCodes[$exitCode]
                if (-not $description) {
                    if ($exitCode -ge 100 -and $exitCode -le 113) {
                        $description = 'Invalid command-line parameters'
                    } else {
                        $description = 'Unrecognized exit code'
                    }
                }

                $exitColor = if ($exitCode -in $successCodes) { 'Green' } elseif ($exitCode -in $benignCodes) { 'Yellow' } else { 'Red' }
                Write-Host "Dell updates completed with exit code $exitCode - $description" -ForegroundColor $exitColor
            }
        }
        # --- Option 5: One-Time Reboot / Shutdown ------------------------------------
        # Immediate or scheduled one-shot SYSTEM reboot/shutdown via schtasks.
        # -----------------------------------------------------------------------------
        '5' {
            $hostname = Read-RexieHostname -CurrentHostname $hostname
            if (-not $hostname) { break }
            $scriptBlock = {
                $isShutdown = $false
                :actionPrompt while ($true) {
                    $actionChoice = $null
                    while ($actionChoice -notin @('R','S')) {
                        $actionChoice = (Read-Host "Reboot or shut down the machine? (R/S, or Q to cancel)").ToUpper()
                        if ($actionChoice -eq 'Q') { Write-Host "Cancelled." -ForegroundColor Yellow; return }
                        if ($actionChoice -notin @('R','S')) {
                            Write-Host "Invalid selection. Please enter R or S, or Q to cancel." -ForegroundColor Yellow
                        }
                    }

                    if ($actionChoice -eq 'R') {
                        $isShutdown = $false
                        break actionPrompt
                    }

                    Write-Host "WARNING: This device will be unavailable until it is manually powered back on." -ForegroundColor Red
                    $confirmChoice = $null
                    while ($confirmChoice -notin @('Y','N','B')) {
                        $confirmChoice = (Read-Host "Are you sure you want to shut down? (Y/N/B to go back)").ToUpper()
                        if ($confirmChoice -notin @('Y','N','B')) {
                            Write-Host "Invalid selection. Please enter Y, N, or B." -ForegroundColor Yellow
                        }
                    }
                    switch ($confirmChoice) {
                        'Y' { $isShutdown = $true; break actionPrompt }
                        'N' { Write-Host "Cancelled." -ForegroundColor Yellow; return }
                        'B' { continue actionPrompt }
                    }
                }
                $powerFlag  = if ($isShutdown) { '/s' } else { '/r' }
                $actionName = if ($isShutdown) { 'Shutdown' } else { 'Reboot' }
                $taskName   = "OneTime$actionName"

                $timingChoice = $null
                while ($timingChoice -notin @('N','L')) {
                    $timingChoice = (Read-Host "Do you want to $($actionName.ToLower()) now or schedule it for later? (N/L, or Q to cancel)").ToUpper()
                    if ($timingChoice -eq 'Q') { Write-Host "Cancelled." -ForegroundColor Yellow; return }
                    if ($timingChoice -notin @('N','L')) {
                        Write-Host "Invalid selection. Please enter N or L, or Q to cancel." -ForegroundColor Yellow
                    }
                }
                switch ($timingChoice) {
                    'N' {
                        Write-Host "$actionName now. Goodbye!" -ForegroundColor Cyan
                        if ($isShutdown) { "__SHUTDOWN_NOW__" } else { "__REBOOT_NOW__" }
                        shutdown.exe $powerFlag /t 0
                        exit
                    }
                    'L' {
                        $validInput = $false
                        while (-not $validInput) {
                            try {
                                $defaultDate = (Get-Date).ToString("MM/dd/yyyy")
                                $defaultTime = "18:00"
                                $inputDate = Read-Host -Prompt "Enter the $($actionName.ToLower()) date (MM/DD/YYYY) [Default: $defaultDate]"
                                if ([string]::IsNullOrWhiteSpace($inputDate)) { $inputDate = $defaultDate }
                                $inputTime = Read-Host -Prompt "Enter the $($actionName.ToLower()) time (HH:MM in 24-hour format) [Default: $defaultTime]"
                                if ([string]::IsNullOrWhiteSpace($inputTime)) { $inputTime = $defaultTime }

                                if ($inputDate -notmatch "/") {
                                    $inputDate = $inputDate.Insert(2, "/").Insert(5, "/")
                                }
                                if ($inputTime -notmatch ":") {
                                    $inputTime = $inputTime.Insert(2, ":")
                                }

                                $scheduledDateTime = [datetime]::ParseExact("$inputDate $inputTime", "MM/dd/yyyy HH:mm", $null)

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
                        # Task deletes itself right before the power action so it doesn't linger in Task Scheduler.
                        $trCommand = "cmd.exe /c `"schtasks /delete /tn $taskName /f & shutdown.exe $powerFlag /t 0`""

                        schtasks /create /tn $taskName /tr $trCommand /sc once /sd $formattedDate /st $formattedTime /RU SYSTEM /F
                        Write-Host "Scheduled $($actionName.ToLower()) at $scheduledDateTime." -ForegroundColor Green
                    }
                }
            }
        }
        # --- Option 6: Hostname Reservation Assistant --------------------------------
        # Validates a 13-character base and how many hostnames are needed, then POSTs
        # "base|count" to RXCRY01TECHLT01:8080 and shows the available hostname(s).
        # The API computes gaps against live AD state on every call - it does not
        # reserve anything, so the count must be requested server-side, not guessed
        # by incrementing a single result locally.
        # -----------------------------------------------------------------------------
        '6' {
            :hostnameLoop do {
                do {
                    $base = (Read-Host "Enter the 13-character hostname base (e.g., RXCRY01TECHLT), or Q to cancel").ToUpper()
                    if ($base -eq 'Q') { Write-Host "Cancelled." -ForegroundColor Yellow; break hostnameLoop }
                    if ($base.Length -ne 13 -or $base -notmatch '^[A-Z0-9]{13}$') {
                        Write-Host "`nInvalid input. Must be exactly 13 alphanumeric characters, or Q to cancel." -ForegroundColor Red
                        $base = $null
                    }
                } while (-not $base)

                $count = $null
                while (-not $count) {
                    $countInput = Read-Host "How many hostnames do you need? (Default: 1, or Q to cancel)"
                    if ($countInput.ToUpper() -eq 'Q') { Write-Host "Cancelled." -ForegroundColor Yellow; break hostnameLoop }
                    if ([string]::IsNullOrWhiteSpace($countInput)) {
                        $count = 1
                    } elseif ($countInput -match '^\d+$' -and [int]$countInput -ge 1 -and [int]$countInput -le 99) {
                        $count = [int]$countInput
                    } else {
                        Write-Host "Invalid input. Enter a number from 1-99, leave blank for 1, or Q to cancel." -ForegroundColor Red
                    }
                }

                $uri = "http://$HostnameApiHost`:$HostnameApiPort/"
                $requestBody = "$base|$count"

                try {
                    $response = Invoke-RestMethod -Method POST -Uri $uri -Body $requestBody -TimeoutSec 10
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
                                $response2 = Invoke-RestMethod -Method POST -Uri $uri -Body $requestBody -TimeoutSec 10
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
                $next = Read-RexieChoice -Prompt "Select an option (1-2)" -ValidChoices '1','2'
                if (-not $next -or $next -eq '2') { break }
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

            Show-LoginShark -Hostname $hostname -Credential $Cred
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

            $sccmChoice = Read-RexieChoice -Prompt "Select an option (1-6)" -ValidChoices '1','2','3','4','5','6'
            if (-not $sccmChoice) { break }
            $sccmActionName = switch ($sccmChoice) {
                '1' { 'Machine Policy Retrieval' }
                '2' { 'User Policy Retrieval' }
                '3' { 'Application Deployment Evaluation' }
                '4' { 'Software Update Evaluation' }
                '5' { 'Hardware Inventory' }
                '6' { 'Full Client Check-In (all cycles)' }
            }
            Write-Status -Level INFO -Message "Selected SCCM action: $sccmActionName"

            $scriptBlock = {
                param($choice)

                $ErrorActionPreference = 'Stop'

                function Invoke-Cycle($guid,$name){
                    try {
                        $result = Invoke-CimMethod -Namespace 'root\ccm' -ClassName 'SMS_Client' -MethodName 'TriggerSchedule' -Arguments @{ sScheduleID = $guid } -ErrorAction Stop
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
                    '2' { Invoke-Cycle '{00000000-0000-0000-0000-000000000026}' 'User Policy Retrieval' }
                    '3' { Invoke-Cycle '{00000000-0000-0000-0000-000000000121}' 'Application Deployment Evaluation' }
                    '4' { Invoke-Cycle '{00000000-0000-0000-0000-000000000108}' 'Software Update Evaluation' }
                    '5' { Invoke-Cycle '{00000000-0000-0000-0000-000000000001}' 'Hardware Inventory' }
                    '6' {
                        Invoke-Cycle '{00000000-0000-0000-0000-000000000021}' 'Machine Policy Retrieval'
                        Invoke-Cycle '{00000000-0000-0000-0000-000000000022}' 'Machine Policy Evaluation'
                        Invoke-Cycle '{00000000-0000-0000-0000-000000000026}' 'User Policy Retrieval'
                        Invoke-Cycle '{00000000-0000-0000-0000-000000000027}' 'User Policy Evaluation'
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
            continue sessionLoop
        }
    }

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
            Write-Host "`nWhat would you like to do?"
            Write-Host "1. Try a different hostname"
            Write-Host "2. Retry the same hostname"
            Write-Host "3. Exit to main menu"
            $offlineChoice = Read-RexieChoice -Prompt "Select an option (1-3)" -ValidChoices '1','2','3'
            if ($offlineChoice -ne '2') { $hostname = $null }
            continue sessionLoop
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

            # If Battery Report (Option 9), open admin share in Finder to C$\HCSTools when possible
            if ($selection -eq '9') {
                if ($invokeResult -contains '__NOT_LAPTOP__') {
                    Write-Host "Battery report not available: device appears to be a desktop (no battery)." -ForegroundColor Yellow
                } elseif ($invokeResult -contains '__REPORT_FAILED__') {
                    Write-Host "Battery report generation failed on the remote device." -ForegroundColor Red
                } else {
                    # Start-Process on macOS just hands the URL to Finder and returns
                    # immediately - it can't detect an async SMB connection failure, so
                    # check reachability first instead of relying on a try/catch fallback.
                    if (Test-TcpPort -HostName $hostname -Port 445) {
                        $smbPath = "smb://$hostname/C$/HCSTools"
                        Write-Status -Level INFO -Message "Opening admin share: $smbPath"
                        try {
                            Start-Process $smbPath
                        } catch {
                            Write-Status -Level WARN -Message "Could not launch Finder for $smbPath. Report saved at C:\HCSTools on $hostname."
                        }
                    } else {
                        Write-Status -Level WARN -Message "SMB (port 445) is not reachable on $hostname. Report saved at C:\HCSTools on $hostname - open it manually once reachable."
                    }
                }
            }

            # If an immediate reboot/shutdown (Option 5) was triggered, the host is about
            # to go offline - clear it on shutdown so "Run another task" can't silently
            # target a machine that's now powered off.
            if ($selection -eq '5') {
                if ($invokeResult -contains '__SHUTDOWN_NOW__') {
                    Write-Status -Level WARN -Message "$hostname was shut down and will not be reachable again until it is powered back on."
                    $hostname = $null
                    $justShutDown = $true
                } elseif ($invokeResult -contains '__REBOOT_NOW__') {
                    Write-Status -Level WARN -Message "$hostname is rebooting now - it may take a minute or two to come back online."
                }
            }
        } catch {
            Write-Status -Level ERROR -Message "Remote command failed: $($_.Exception.Message)"
        }
    }

    # --- Post-Task Prompt ---------------------------------------------------------
    # Lets the operator reuse the same hostname, switch hosts, or exit the session.
    # Option 6 doesn't target a specific hostname and already has its own "try
    # another / return to main menu" loop, so it skips straight back here.
    # -----------------------------------------------------------------------------
    if ($selection -eq '6') {
        continue sessionLoop
    }

    Write-Status -Level INFO -Message "What would you like to do next?"
    if ($justShutDown) {
        Write-Host "1. Return to main menu"
    } else {
        Write-Host "1. Return to main menu (keep hostname: $hostname)"
    }
    Write-Host "2. Return to main menu (enter a new hostname)"
    Write-Host "3. Exit"
    $nextAction = Read-RexieChoice -Prompt "Select an option (1-3)" -ValidChoices '1','2','3'
    switch ($nextAction) {
        '1' { }  # continue with same hostname
        '2' { $hostname = $null }
        '3' { $repeatSession = $false }
        $null { }  # cancelled with Q - back out safely, same as "return to main menu", not Exit
    }
    if ($repeatSession -eq $false) {
        $hostname = $null
    }
} while ($repeatSession)
#endregion Session Loop & Credential Handling
