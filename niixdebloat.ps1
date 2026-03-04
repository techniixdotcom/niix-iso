# Windows ISO Debloater - Customized
# Description: Modifies a Windows ISO by debloating + Tweaking the registry           
# Thanks to: itsNileshHere, ChrisTitusTech

param(
    # -- noPrompt is ON by default -- all choices are pre-set below -------------
    # The only interactive prompts remaining are:
    #   1. File picker dialog  -> select your source ISO
    #   2. Windows edition     -> pick which edition from the ISO (e.g. Home / Pro)
    #   3. Output ISO name     -> type the filename for the finished ISO
    [switch]$noPrompt = $true,
    [string]$isoPath = "",
    [string]$winEdition = "",
    [string]$outputISO = "",
    [ValidateSet("yes", "no")]$useDISM = "",

    # -- Core Operations -------------------------------------------------------
    [ValidateSet("yes", "no")]$AppxRemove          = "yes",   # Remove unnecessary packages     (default: yes)
    [ValidateSet("yes", "no")]$CapabilitiesRemove  = "yes",   # Remove unnecessary features     (default: yes)
    [ValidateSet("yes", "no")]$OnedriveRemove      = "yes",   # Remove OneDrive                 (default: yes)
    [ValidateSet("yes", "no")]$EDGERemove          = "yes",   # Remove Microsoft Edge           (default: yes)
    [ValidateSet("yes", "no")]$AIRemove            = "yes",   # Remove AI Components            (default: yes)
    [ValidateSet("yes", "no")]$TPMBypass           = "yes",   # Bypass TPM check                (default: no)
    [ValidateSet("yes", "no")]$UserFoldersEnable   = "yes",   # Enable user folders             (default: yes)
    [ValidateSet("yes", "no")]$DriverIntegrate     = "no",    # Integrate Intel RST/VMD drivers (default: no)
    [ValidateSet("yes", "no")]$ESDConvert          = "no",    # Compress the ISO                (default: no)
    [ValidateSet("yes", "no")]$useOscdimg          = "yes",   # Use Oscdimg for ISO creation    (default: yes)

    # -- Privacy & Tweaks ------------------------------------------------------
    [ValidateSet("yes", "no")]$DisableActivityHistory         = "yes",  # 11  (default: yes)
    [ValidateSet("yes", "no")]$DisableLocationTracking        = "yes",  # 12  (default: yes)
    [ValidateSet("yes", "no")]$DisablePS7Telemetry            = "yes",  # 13  (default: yes)
    [ValidateSet("yes", "no")]$DisableWPBT                    = "yes",  # 14  (default: yes)
    [ValidateSet("yes", "no")]$DisableCrossDeviceResume       = "yes",  # 15  (default: yes)
    [ValidateSet("yes", "no")]$DisableHibernation             = "no",   # 16  (default: no)
    [ValidateSet("yes", "no")]$SetServicesManual              = "yes",  # 17  (default: yes)
    [ValidateSet("yes", "no")]$DisableBackgroundApps          = "yes",  # 18  (default: yes)
    [ValidateSet("yes", "no")]$DisableFullscreenOptimizations = "yes",  # 19  (default: yes)
    [ValidateSet("yes", "no")]$SetClassicRightClickMenu       = "yes",  # 20  (default: yes)
    [ValidateSet("yes", "no")]$EnableEndTaskRightClick        = "yes",  # 21  (default: yes)
    [ValidateSet("yes", "no")]$DisableExplorerAutoDiscovery   = "yes",  # 22  (default: yes)
    [ValidateSet("yes", "no")]$SetDarkTheme                   = "yes",  # 23  (default: yes)
    [ValidateSet("yes", "no")]$DisableTeredo                  = "yes",  # 24  (default: yes)
    [ValidateSet("yes", "no")]$PreferIPv4overIPv6             = "no",   # 25  (default: no)
    [ValidateSet("yes", "no")]$DisableIPv6                    = "yes",  # 26  (default: no)
    [ValidateSet("yes", "no")]$RemoveXboxComponents           = "yes",  # 27  (default: yes)
    [ValidateSet("yes", "no")]$DisableGameBarProtocols        = "yes",  # 28  (default: yes)
    [ValidateSet("yes", "no")]$DisableCopilotExtra            = "yes",  # 29  (default: yes)
    [ValidateSet("yes", "no")]$BlockAdobeNetwork              = "yes",  # 30  (default: no)
    [ValidateSet("yes", "no")]$BlockRazerInstalls             = "yes",  # 31  (default: no)
    [ValidateSet("yes", "no")]$SetDisplayForPerformance       = "yes",  # 32  (default: yes)
    [ValidateSet("yes", "no")]$EnableDetailedBSoD             = "yes",  # 33  (default: yes)
    [ValidateSet("yes", "no")]$DisableBingSearch              = "yes",  # 34  (default: yes)

    # -- Security & Privacy Deep-Clean -----------------------------------------
    [ValidateSet("yes", "no")]$DisableWER                  = "yes",  # A1  (default: yes)
    [ValidateSet("yes", "no")]$DisableDeliveryOptimization = "yes",  # A3  (default: yes)
    [ValidateSet("yes", "no")]$DisableAutoLogger           = "yes",  # A4  (default: yes)
    [ValidateSet("yes", "no")]$DisableCEIP                 = "yes",  # A5  (default: yes)
    [ValidateSet("yes", "no")]$RedirectNTP                 = "yes",  # B5  (default: yes)
    [ValidateSet("yes", "no")]$DisableAppAccountInfo       = "yes",  # C1  (default: yes)
    [ValidateSet("yes", "no")]$DisableAppContactsCalendar  = "yes",  # C2  (default: yes)
    [ValidateSet("yes", "no")]$DisableAppCameraMic         = "yes",  # C3  (default: yes)
    [ValidateSet("yes", "no")]$DisableAppMessaging         = "yes",  # C4  (default: yes)
    [ValidateSet("yes", "no")]$DisableClipboardHistory     = "yes",  # C5  (default: yes)
    [ValidateSet("yes", "no")]$DisableSmartScreenExplorer  = "yes",  # D1  (default: yes)
    [ValidateSet("yes", "no")]$DisableSmartScreenStore     = "yes",  # D2  (default: yes)
    [ValidateSet("yes", "no")]$DisableDefenderMAPS         = "yes",  # D3  (default: yes)
    [ValidateSet("yes", "no")]$DisableCloudSearch          = "yes",  # D4  (default: yes)
    [ValidateSet("yes", "no")]$RemoveDeviceCensusTask      = "yes",  # E2  (default: yes)
    [ValidateSet("yes", "no")]$RemoveStartupAppTask        = "yes",  # E4  (default: yes)
    [ValidateSet("yes", "no")]$DisableRemoteAssistance     = "yes",  # F1  (default: yes)
    [ValidateSet("yes", "no")]$DisableAutoRun              = "yes",  # F2  (default: yes)
    [ValidateSet("yes", "no")]$DisableDCOM                 = "no"    # F3  (default: no)
)

# noPrompt is on by default -- isoPath, winEdition, and outputISO are still
# prompted interactively (file picker, edition list, and ISO name input).
# All other yes/no choices are pre-set in the param block above.

# Disable Pause if -noprompt is used
if ($noPrompt) { function Pause { } }

# Administrator Privileges
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "This script must be run as Administrator. Re-launching with elevated privileges..." -ForegroundColor Yellow
    
    # Resolve relative paths before relaunching
    if ($isoPath -and -not [System.IO.Path]::IsPathRooted($isoPath)) {
        $isoPath = Join-Path -Path $PSScriptRoot -ChildPath $isoPath | Resolve-Path -ErrorAction SilentlyContinue
        if (-not $isoPath) {
            $isoPath = Join-Path -Path (Get-Location).Path -ChildPath $PSBoundParameters['isoPath']
            if (Test-Path $isoPath) {
                $isoPath = (Get-Item $isoPath).FullName
            }
        }
    }
    if ($outputISO -and -not [System.IO.Path]::IsPathRooted($outputISO)) {
        $outputISO = Join-Path -Path (Get-Location).Path -ChildPath $outputISO
        $outputISO = [System.IO.Path]::GetFullPath($outputISO)
    }
    
    $params = @()
    $PSBoundParameters.GetEnumerator() | ForEach-Object {
        if ($_.Value -is [switch] -and $_.Value) { $params += "-$($_.Key)" }
        elseif ($_.Value -is [string] -and $_.Value) { 
            # Use resolved paths for isoPath and outputISO
            if ($_.Key -eq 'isoPath' -and $isoPath) { $params += "-$($_.Key)", "`"$isoPath`"" }
            elseif ($_.Key -eq 'outputISO' -and $outputISO) { $params += "-$($_.Key)", "`"$outputISO`"" }
            else { $params += "-$($_.Key)", "`"$($_.Value)`"" }
        }
    }    
    $argss = "-NoProfile -ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Path)`" $($params -join ' ')"
    if (Get-Command wt -ErrorAction SilentlyContinue) { Start-Process wt "PowerShell $argss" -Verb RunAs }
    else { Start-Process PowerShell $argss -Verb RunAs }
    Exit
}
Clear-Host
# -- Colour palette ------------------------------------------------------------
$C = @{
    Pink    = [ConsoleColor]::Magenta
    Purple  = [ConsoleColor]::DarkMagenta
    Green   = [ConsoleColor]::Green
    Lime    = [ConsoleColor]::DarkGreen
    Cyan    = [ConsoleColor]::Cyan
    White   = [ConsoleColor]::White
    Gray    = [ConsoleColor]::DarkGray
    Yellow  = [ConsoleColor]::Yellow
    Red     = [ConsoleColor]::Red
}

# -- Progress bar helper -------------------------------------------------------
function Show-Bar {
    param([int]$Pct, [int]$Width = 40, [string]$Label = "")
    $filled = [int]([Math]::Round($Pct / 100 * $Width))
    $empty  = $Width - $filled
    $bar    = ("#" * $filled) + ("-" * $empty)
    Write-Host -NoNewline "  "
    Write-Host -NoNewline $bar           -ForegroundColor Magenta
    Write-Host -NoNewline "  "
    Write-Host -NoNewline ("{0,3}%" -f $Pct) -ForegroundColor Green
    if ($Label) { Write-Host -NoNewline "  $Label" -ForegroundColor DarkGray }
    Write-Host ""
}

# -- Section header helper -----------------------------------------------------
function Show-Section {
    param([string]$Title, [string]$Icon = "*")
    $line = "-" * 58
    Write-Host ""
    Write-Host "  $Icon  $Title" -ForegroundColor Magenta
    Write-Host "  $line"         -ForegroundColor DarkMagenta
}

# -- Step result helper --------------------------------------------------------
function Show-Done  { param([string]$Msg) Write-Host "  v  $Msg" -ForegroundColor Green }
function Show-Skip  { param([string]$Msg) Write-Host "  -  $Msg" -ForegroundColor DarkGray }
function Show-Info  { param([string]$Msg) Write-Host "  .  $Msg" -ForegroundColor Cyan }
function Show-Warn  { param([string]$Msg) Write-Host "  !  $Msg" -ForegroundColor Yellow }
function Show-Fail  { param([string]$Msg) Write-Host "  x  $Msg" -ForegroundColor Red }

# -- Banner --------------------------------------------------------------------
$banner = @"

    )             )   (                                      
 ( /(          ( /(   )\ )            )  (                )  
 )\()) (   (   )\()) (()/(     (   ( /(  )\         )  ( /(  
((_)\  )\  )\ ((_)\   /(_))   ))\  )\())((_) (   ( /(  )\()) 
 _((_)((_)((_)__((_) (_))_   /((_)((_)\  _   )\  )(_))(_))/  
| \| | (_) (_)\ \/ /  |   \ (_))  | |(_)| | ((_)((_)_ | |_   
| .` | | | | | >  <   | |) |/ -_) | '_ \| |/ _ \/ _` ||  _|  
|_|\_| |_| |_|/_/\_\  |___/ \___| |_.__/|_|\___/\__,_| \__|  
                                                             
"@
Write-Host $banner -ForegroundColor Magenta
Write-Host "                    NiiX DEBLOATER   //  Privacy + Performance" -ForegroundColor DarkMagenta
Write-Host ("  " + ("-" * 73)) -ForegroundColor DarkMagenta
Write-Host ""

$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'
$PSDefaultParameterValues['*:Encoding'] = 'utf8'
$scriptDirectory = "$PSScriptRoot"
$logFilePath = Join-Path -Path $scriptDirectory -ChildPath 'script_log.txt'         # Log File Path
$transcript = "$env:TEMP\transcript_$(Get-Random).txt"                              # Start Transcript
Start-Transcript $transcript -Append -ErrorAction SilentlyContinue 2>&1 | Out-Null

# Get system information (WMI-free: uses registry and env vars only, no timeout risk)
$osCaption  = $(if($p=(Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -ErrorAction SilentlyContinue)){$p.ProductName}else{"Unknown"})
$osBuild    = $(if($p=(Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -ErrorAction SilentlyContinue)){$p.CurrentBuildNumber}else{"Unknown"})
$osVersion  = [System.Environment]::OSVersion.Version.ToString()
$osArch     = $env:PROCESSOR_ARCHITECTURE
$launchedAs = "$($MyInvocation.MyCommand.Path)"
$logEntry = @"
$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Script started
- Launched As: $launchedAs
- Windows Version: $osCaption (Build $osBuild)
- OS Version: $osVersion
- System Architecture: $osArch
- System Language: $((Get-Culture).DisplayName)
- Default Language: $((Get-UICulture).DisplayName)
- Windows Directory: $($env:windir)`n
"@

# Initialize log file
$logEntry | Out-File -FilePath $logFilePath -Append

# Function to write logs
function Write-Log {
    [CmdletBinding()]
    param ([Parameter(ValueFromPipeline=$true)][object]$InputObj, [string]$msg, [switch]$Raw, [string]$Sep = " || ")
    process {
        $content = if ($msg) { $msg } elseif ($null -ne $InputObj) { if ($InputObj -is [string]) { $InputObj } else { $InputObj | Out-String } } else { return }
        if (-not $Raw -and ($content = $content.Trim())) {
            $lines = @($content -split '\n' | Where-Object { $_.Trim() })
            $cut = $lines | Where-Object { $_ -match '^\s*\+\s*(CategoryInfo|FullyQualifiedErrorId)\s*:' } | Select-Object -First 1
            if ($cut) { $lines = $lines[0..($lines.IndexOf($cut) - 1)] }
            if ($lines.Count -gt 1) {
                $processedLines = foreach ($line in $lines) {
                    $trimmed = $line.Trim()
                    if ($trimmed -match '^At\s+(.+)') { "At $($matches[1])" }
                    elseif ($trimmed -match '^\s*\+\s*~+') { continue }  # Skip underline line
                    elseif ($trimmed -match '^\s*\+\s*(.+)') { "+ " + ($matches[1] -replace '\s{2,}', ' ') }
                    elseif ($trimmed -match '^\s*\+?\s*(\w+\w+)\s*:\s*(.+)') { "$($matches[1]): $($matches[2])" }
                    elseif ($trimmed -notmatch '^-{4,}' -and $trimmed) { $trimmed -replace '\s{2,}', ' ' }
                }
                $content = $processedLines -join $Sep
            } else { $content = $content -replace '\s{2,}', ' ' }
        }
        if ($content) { Add-Content -Path "$logFilePath" -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $content" }
    }
}

# Function to invoke DISM commands if powershell fails
function Invoke-DismFailsafe {
    param([scriptblock]$PS, [scriptblock]$Dism)
    if ($useDISM -ieq "yes") {
        & $Dism 2>&1 | Write-Log
    } else {
        try { & $PS 2>&1 | Write-Log } catch { & $Dism 2>&1 | Write-Log }
    }
}

# Confirmation Function
function Get-Confirmation { 
    param([string]$Question, [bool]$DefaultValue = $true, [string]$Description = "") 
    $defaultText = if ($DefaultValue) { "Y" } else { "N" }
    $optionsText = if ($DefaultValue) { "Y/n" } else { "y/N" }
    do { 
        Write-Host "$Question" -ForegroundColor Cyan -NoNewline
        if ($Description) { Write-Host " - $Description" -ForegroundColor DarkGray -NoNewline }
        Write-Host " ($optionsText): " -ForegroundColor White -NoNewline
        $answer = Read-Host 
        if ([string]::IsNullOrWhiteSpace($answer)) {
            Write-Host "Using default: $defaultText" -ForegroundColor Yellow
            return $DefaultValue
        }
        $answer = $answer.ToUpper()
        if ($answer -eq 'Y') { return $true }
        if ($answer -eq 'N') { return $false }
        Write-Warning "Invalid input. Enter 'Y' for Yes, 'N' for No, or Enter for default ($defaultText)."
    } while ($true) 
}

# Parameter Value Validation Function
function Get-ParameterValue {
    param( [string]$ParameterValue, [bool]$DefaultValue, [string]$Question, [string]$Description )
    if ($ParameterValue -ne "") { return $ParameterValue -eq "yes" }
    if ($noPrompt) { return $DefaultValue }
    # If neither noPrompt nor param was provided, prompt the user
    return Get-Confirmation -Question $Question -DefaultValue $DefaultValue -Description $Description
}

# Cleanup Function
function Remove-TempFiles {
    Remove-Item -Path $destinationPath -Recurse -Force 2>&1 | Write-Log
    Remove-Item -Path $installMountDir -Recurse -Force 2>&1 | Write-Log
    Remove-Item -Path (Join-Path $scriptDirectory "WIDTemp") -Recurse -Force 2>&1 | Write-Log
    Stop-Transcript 2>&1 | Write-Log
    $content = Get-Content $transcript | Where-Object { $_ -notmatch "^(Windows PowerShell transcript|Start time:|Username:|RunAs User:|Configuration|Host Application:|Process ID:|PS[A-Z]|BuildVersion:|CLRVersion:|WSManStackVersion:|SerializationVersion:|Transcript started|PS C:\\|^\*{10,}|End time:)" -and $_.Trim() }
    Add-Content $logFilePath -Value ("`n" + "="*50 + "`nTerminal Snapshot - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" + "`n" + "="*50 + "`n" + ($content -join "`n"))
    Remove-Item $transcript  -Force 2>&1 | Write-Log
}

# Set Ownership Permissions
function Set-Ownership {
    param([string]$Path, [string[]]$Registry) 
    if ($Path) {
        try {
            $FullPath = [System.IO.Path]::GetFullPath($Path)
            if (-not (Test-Path -Path $FullPath)) { return $true }
            $IsFolder = (Get-Item $FullPath).PSIsContainer
            
            # Try ACL method
            try {
                $Acl = Get-Acl $FullPath
                $Acl.SetOwner([System.Security.Principal.NTAccount]"Administrators")
                $CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
                $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($CurrentUser, "FullControl", $(if ($IsFolder) {"ContainerInherit,ObjectInherit"} else {"None"}), "None", "Allow")
                $Acl.SetAccessRule($AccessRule)
                Set-Acl -Path $FullPath -AclObject $Acl
                
                if ($IsFolder) { Get-ChildItem -Path $FullPath -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object { 
                        try { $ChildAcl = Get-Acl $_.FullName
                            $ChildAcl.SetOwner([System.Security.Principal.NTAccount]"Administrators")
                            $ChildAcl.SetAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($CurrentUser, "FullControl", "Allow")))
                            Set-Acl -Path $_.FullName -AclObject $ChildAcl 
                        }
                        catch {}
                    }
                }
                Write-Log -msg "[ACL] Set ownership for: $FullPath"
                return $true
            }

            # Fallback to icals
            catch { Write-Log -msg "ACL method failed for: $FullPath"
                try {
                    & icacls.exe "$FullPath" /setowner "Administrators" /T /C 2>&1 | Out-Null
                    & icacls.exe "$FullPath" /grant "${CurrentUser}:(F)" /T /C 2>&1 | Out-Null
                    & icacls.exe "$FullPath" /grant "Administrators:(F)" /T /C 2>&1 | Out-Null
                    Write-Log -msg "[icacls] Set ownership for: $FullPath"
                    return $true
                }
                catch { Write-Log -msg "icacls fallback failed for: $FullPath - $($_.Exception.Message)"; return $false }
            }
        } 
        catch { Write-Log -msg "Failed to own path: $Path - $($_.Exception.Message)"; return $false }
    }
    if ($Registry) {
        try {
            $sid = (New-Object System.Security.Principal.NTAccount("BUILTIN\Administrators")).Translate([System.Security.Principal.SecurityIdentifier])
            $rule = New-Object System.Security.AccessControl.RegistryAccessRule("Administrators", "FullControl", "ContainerInherit", "None", "Allow")
            foreach ($keyPath in $Registry) {
                try {
                    $key = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($keyPath, [Microsoft.Win32.RegistryKeyPermissionCheck]::ReadWriteSubTree, [System.Security.AccessControl.RegistryRights]::TakeOwnership)
                    if ($key) { $acl = $key.GetAccessControl()
                        $acl.SetOwner($sid)
                        $key.SetAccessControl($acl)
                        $key.Close()
                        $key = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($keyPath, [Microsoft.Win32.RegistryKeyPermissionCheck]::ReadWriteSubTree, [System.Security.AccessControl.RegistryRights]::ChangePermissions)
                        if ($key) { $acl = $key.GetAccessControl()
                            $acl.SetAccessRule($rule)
                            $key.SetAccessControl($acl)
                            $key.Close()
                            Write-Log -msg "Set ownership for registry: $keyPath"
                        }
                    } else { Write-Log -msg "Unable to open reg-key: $keyPath" }
                } catch {}
            }
            return $true
        } catch { Write-Log -msg "Failed to own reg-key: $($_.Exception.Message)"; return $false }
    }
    return $false
}

# Force Remove Function
function Set-OwnAndRemove {
    param([Parameter(Mandatory=$true)][string]$Path)
    try {
        $FullPath = [System.IO.Path]::GetFullPath($Path)
        if (-not (Test-Path -Path $FullPath)) { return $true }
        try {
            $ownershipResult = Set-Ownership -Path $Path
            if (-not $ownershipResult) { throw "ACL method failed" }
            Remove-Item -Path $FullPath -Force -Recurse -ErrorAction Stop
            Write-Log -msg "Removed with ACL: $FullPath"
            return $true
        } catch {
            Write-Log -msg "ACL method failed for: $FullPath"
            try {
                $IsFolder = (Get-Item $FullPath -ErrorAction Stop).PSIsContainer
                $CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
                if($IsFolder) { takeown /F "$FullPath" /R /D Y 2>&1 | Write-Log }
                else { takeown /F "$FullPath" /A 2>&1 | Write-Log }
                foreach ($Perm in @("*S-1-5-32-544:F", "System:F", "Administrators:F", "$CurrentUser`:F")) {
                    try {
                        if($IsFolder) { icacls "$FullPath" /grant:R "$Perm" /T /C 2>&1 | Write-Log }
                        else { icacls "$FullPath" /grant:R "$Perm" 2>&1 | Write-Log }
                        if ($LASTEXITCODE -eq 0) { break }
                    } catch { continue }
                }
                Remove-Item -Path $FullPath -Force -Recurse -ErrorAction Stop
                Write-Log -msg "Removed with icacls: $FullPath"
                return $true
            } catch { Write-Log -msg "Failed to remove: $FullPath - $($_.Exception.Message)"; return $false }
        }
    } catch { Write-Log -msg "Error processing path: $Path - $($_.Exception.Message)"; return $false }
}

# Function to check internet connection
function Test-InternetConnection {
    param (
        [int]$maxAttempts = 3,
        [int]$retryDelay = 5,
        [string]$hostname = "1.1.1.1", # Cloudflare DNS
        [int]$port = 53,
        [int]$timeout = 5000
    )
    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        try {
            $client = [Net.Sockets.TcpClient]::new()
            if ($client.ConnectAsync($hostname, $port).Wait($timeout)) {
                $client.Close(); return $true
            }
            $client.Close()
        } catch {}
        Write-Host "Internet connection not available, Trying in $retryDelay seconds..."
        Start-Sleep -Seconds $retryDelay
    }  
    Write-Host "`nInternet connection not available after $maxAttempts attempts." -ForegroundColor Red
    Write-Host "A working internet connection is required to download oscdimg.exe."
    Write-Host "Check your connection and try again."

    while ($true) {
        $internetChoice = Read-Host -Prompt "`nPress 't' to try again or 'q' to quit"
        switch ($internetChoice.ToLower()) {
            't' { return Test-InternetConnection @PSBoundParameters }
            'q' {
                Remove-TempFiles
                Exit
            }
            default { Write-Host "Invalid input. Enter 't' or 'q'." }
        }
    }
}

# Image Info Function
function Get-WimDetails {
    param ( [Parameter(Mandatory = $true)][string]$MountPath )
    try {
        $out = dism /Image:$MountPath /Get-Intl /English | Out-String
        Write-Log -msg "DISM Output for Get-WimDetails:`n$out"
        $buildMatch = [regex]::Match($out, "Image Version: \d+\.\d+\.(\d+)\.\d+")
        $langMatch = [regex]::Match($out, "(?i)Default\s+system\s+UI\s+language\s*:\s*([a-z]{2}-[A-Z]{2})")
        [PSCustomObject]@{
            BuildNumber = if ($buildMatch.Success) { $buildMatch.Groups[1].Value } else { $null }
            Language = if ($langMatch.Success) { $langMatch.Groups[1].Value } else { $null }
        }
    }
    catch {
        Write-Host "Failed to get WIM info: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

# Get Image Index Function
function Get-ImageIndex {
    param ( [Parameter(Mandatory = $true)][string]$ImagePath )
    try {
        $out = & dism.exe /get-wiminfo /wimfile:$ImagePath /english 2>$null
        Write-Log -msg "DISM Output for Get-ImageIndex:`n$out"
        if ($LASTEXITCODE -ne 0) { throw "DISM failed to read image file: $ImagePath" }
        $images = @()
        $indexPattern = "Index\s*:\s*(\d+)"
        $namePattern = "Name\s*:\s*(.+)"
        for ($i = 0; $i -lt $out.Count; $i++) {
            if ($out[$i] -match $indexPattern) {
                $index = $matches[1]
                for ($j = $i + 1; $j -lt [Math]::Min($i + 5, $out.Count); $j++) {
                    if ($out[$j] -match $namePattern) {
                        $name = $matches[1].Trim()
                        $images += [PSCustomObject]@{
                            Index = [int]$index
                            ImageName = $name
                        }
                        break
                    }
                }
            }
        }
        return $images
    }
    catch {
        Write-Log -msg "Failed to get image information: $($_.Exception.Message)"
        return $null
    }
}

# Oscdimg Path
$OscdimgPath = "$env:SystemDrive\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg"
$Oscdimg = Join-Path -Path $OscdimgPath -ChildPath 'oscdimg.exe'

# Autounattend.xml Path
$autounattendXmlPath = Join-Path -Path $scriptDirectory -ChildPath "Autounattend.xml"

# Download Autounattend.xml if not exists
if (-not (Test-Path $autounattendXmlPath)) {
    $ProgressPreference = 'SilentlyContinue'
    try { Invoke-WebRequest "https://itsnileshhere.github.io/Windows-ISO-Debloater/autounattend.xml" -OutFile $autounattendXmlPath -UseBasicParsing }
    catch { Write-Log -msg "Warning: Unable to download Autounattend.xml" }
    finally { $ProgressPreference = 'Continue' }
}

# Mount ISO Dialog
# -- Auto-detect ISO in script directory --------------------------------------
if ($isoPath) {
    $isoFilePath = $isoPath
} else {
    $isoFiles = Get-ChildItem -Path $scriptDirectory -Filter "*.iso" -File | Sort-Object Name
    if ($isoFiles.Count -eq 0) {
        Write-Host "`n[ERROR] No ISO files found in script directory: $scriptDirectory" -ForegroundColor Red
        Write-Host "        Place a Windows ISO in the same folder as this script and run again." -ForegroundColor Yellow
        Write-Log -msg "No ISO files found in script directory"
        Pause
        Exit
    } elseif ($isoFiles.Count -eq 1) {
        $isoFilePath = $isoFiles[0].FullName
        Show-Done "Auto-detected: $($isoFiles[0].Name)"
        Write-Log -msg "Auto-detected ISO: $isoFilePath"
    } else {
        Show-Warn "Multiple ISO files found -- please choose one:"
        for ($i = 0; $i -lt $isoFiles.Count; $i++) {
            Write-Host "  $($i + 1). $($isoFiles[$i].Name)" -ForegroundColor White
        }
        do {
            $isoChoice = Read-Host -Prompt "`nEnter the number of the ISO to use"
            $isoIndex  = [int]$isoChoice - 1
        } while ($isoIndex -lt 0 -or $isoIndex -ge $isoFiles.Count)
        $isoFilePath = $isoFiles[$isoIndex].FullName
        Show-Done "Selected: $($isoFiles[$isoIndex].Name)"
        Write-Log -msg "User selected ISO: $isoFilePath"
    }
}
if ($null -eq $isoFilePath -or -not (Test-Path $isoFilePath)) {
    Write-Host "ISO not found. Exiting Script" -ForegroundColor Red
    Write-Log -msg "ISO not found: $isoFilePath"
    Pause
    Exit
}

Show-Done "ISO selected: $(Split-Path $isoFilePath -Leaf)"
Write-Log -msg "ISO Path: $isoFilePath"

# -- Locate or install 7-Zip ---------------------------------------------------
function Find-SevenZip {
    # Check standard install locations
    $candidates = @(
        "$env:ProgramFiles\7-Zip\7z.exe",
        "${env:ProgramFiles(x86)}\7-Zip\7z.exe"
    )
    foreach ($path in $candidates) {
        if (Test-Path $path) { return $path }
    }
    # Also check PATH
    $inPath = Get-Command "7z.exe" -ErrorAction SilentlyContinue
    if ($inPath) { return $inPath.Source }
    return $null
}

function Install-SevenZip {
    Show-Section "Attempting to install 7-Zip..."
    Write-Log -msg "Installing 7-Zip"

    # Try winget first
    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if ($winget) {
        Show-Info "Installing via winget..."
        & winget install --id 7zip.7zip --silent --accept-package-agreements --accept-source-agreements 2>&1 | Write-Log
        $found = Find-SevenZip
        if ($found) {
            Show-Done "7-Zip installed via winget"
            Write-Log -msg "7-Zip installed via winget at $found"
            return $found
        }
    }

    # Fall back to direct download from 7-zip.org
    Show-Info "Downloading from 7-zip.org..."
    Write-Log -msg "Downloading 7-Zip installer directly"
    $installerUrl = "https://www.7-zip.org/a/7z2408-x64.exe"
    $installerPath = Join-Path $env:TEMP "7zip_installer.exe"
    try {
        Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath -UseBasicParsing -ErrorAction Stop
        Show-Info "Running silent installer..."
        $proc = Start-Process -FilePath $installerPath -ArgumentList "/S" -Wait -PassThru
        Remove-Item $installerPath -Force -ErrorAction SilentlyContinue
        if ($proc.ExitCode -eq 0) {
            $found = Find-SevenZip
            if ($found) {
                Show-Done "7-Zip installed"
                Write-Log -msg "7-Zip installed via direct download at $found"
                return $found
            }
        }
        throw "Installer exited with code $($proc.ExitCode)"
    }
    catch {
        Write-Log -msg "7-Zip direct download/install failed: $_"
        return $null
    }
}

Show-Section "Checking for 7-Zip..."
$sevenZip = Find-SevenZip
if ($sevenZip) {
    Show-Done "7-Zip found at $sevenZip"
    Write-Log -msg "7-Zip found at $sevenZip"
} else {
    Write-Host "`n  7-Zip was not found on this system." -ForegroundColor Yellow
    Write-Host "  It is required to extract the ISO." -ForegroundColor Yellow
    $install7z = Get-Confirmation -Question "Install 7-Zip now?" -DefaultValue $true -Description "Will try winget first, then direct download from 7-zip.org"
    if ($install7z) {
        $sevenZip = Install-SevenZip
    }
    if (-not $sevenZip) {
        Write-Host "`n[ERROR] 7-Zip could not be found or installed." -ForegroundColor Red
        Write-Host "        Please install it manually from https://7-zip.org and re-run the script." -ForegroundColor Yellow
        Write-Log -msg "7-Zip not available - aborting"
        Pause
        Exit
    }
}

$destinationPath = Join-Path $scriptDirectory "WIDTemp\winlite"          # Destination Path (script dir)
$installMountDir = Join-Path $scriptDirectory "WIDTemp\mountdir\installWIM"  # Mount Directory (script dir)

Show-Section "Extracting ISO" "*"
Show-Info "Source: $isoFilePath"
Write-Host "  From: " -NoNewline; Write-Host "`"$isoFilePath`"" -ForegroundColor Yellow
Write-Host "  To:   " -NoNewline; Write-Host "`"$destinationPath`"" -ForegroundColor Yellow
Write-Log -msg "Extracting ISO with 7-Zip from $isoFilePath to $destinationPath"

try {
    if (-not (Test-Path $destinationPath)) { New-Item -ItemType Directory -Path $destinationPath -Force -EA Stop | Out-Null }
    $7zOutput = & $sevenZip x "$isoFilePath" -o"$destinationPath" -y 2>&1
    $7zOutput | Write-Log
    if ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq 1) {
        Show-Bar 100 40 "Extraction complete"
Show-Done "ISO extracted successfully"
        Write-Log -msg "7-Zip extraction completed (Exit: $LASTEXITCODE)"
        Write-Log -msg "Removing read-only attributes..."
        Get-ChildItem -Path $destinationPath -Recurse | ForEach-Object { $_.Attributes = $_.Attributes -band (-bnot [System.IO.FileAttributes]::ReadOnly) } | Out-Null
    }
    else { throw "7-Zip extraction failed with exit code: $LASTEXITCODE" }
} catch {
    Write-Host "Extraction failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Log -msg "Extraction failed: $($_.Exception.Message)"
    Pause
    Exit
}

# Check files availability
$installWimPath = Join-Path $destinationPath "sources\install.wim"
$installEsdPath = Join-Path $destinationPath "sources\install.esd"
New-Item -ItemType Directory -Path $installMountDir 2>&1 | Write-Log

# Handling install.wim and install.esd
if (-not (Test-Path $installWimPath)) {
    Write-Host "`ninstall.wim not found. Searching for install.esd..."
    if (Test-Path $installEsdPath) {
        Write-Host "`ninstall.esd found at " -NoNewline -ForegroundColor Cyan; Write-Host "$installEsdPath"
        Write-Log -msg "install.esd found. Converting..."
        Write-Host "Details for image: " -NoNewline -ForegroundColor Cyan; Write-Host "$installEsdPath"
        try {
            # Get image info from install.esd
            $esdInfo = Get-ImageIndex -ImagePath $installEsdPath
            if (-not $esdInfo) { 
                Write-Host "Error: Could not retrieve image info from WIM file" -ForegroundColor Red
                Remove-TempFiles
                Pause
                Exit
            }
            # Print image details from install.esd
            foreach ($image in $esdInfo) {
                Write-Host "$($image.Index). $($image.ImageName)"
            }
            # If winEdition is specified, find the index; else prompt user
            if ($winEdition) {
                $matchedImage = $esdInfo | Where-Object { $_.ImageName -ieq $winEdition }
                if ($matchedImage) { $sourceIndex = $matchedImage.Index }
                else { $sourceIndex = 1 }
            }
            else { $sourceIndex = Read-Host -Prompt "`nEnter the index to convert and mount" }
            # Check if the index is valid, print selected "ImageIndex - ImageName"
            $selectedImage = $esdInfo | Where-Object { $_.Index -eq [int]$sourceIndex }
            if ($selectedImage) {
                Write-Host "`nMounting image: " -NoNewline -ForegroundColor Cyan; Write-Host "$sourceIndex. $($selectedImage.ImageName)"
                Write-Log -msg "Converting and Mounting image: $sourceIndex. $($selectedImage.ImageName)"
            }

            # Convert ESD to WIM
            Invoke-DismFailsafe {Export-WindowsImage -SourceImagePath $installEsdPath -SourceIndex $sourceIndex -DestinationImagePath $installWimPath -CompressionType Maximum -CheckIntegrity} {dism /Export-Image /SourceImageFile:$installEsdPath /SourceIndex:$sourceIndex /DestinationImageFile:$installWimPath /Compress:max /CheckIntegrity}
            # Remove the ESD file after conversion
            Remove-Item $installEsdPath -Force
            # Mount the converted WIM with SourceIndex 1
            Invoke-DismFailsafe {Mount-WindowsImage -ImagePath $installWimPath -Index 1 -Path $installMountDir} {dism /mount-image /imagefile:$installWimPath /index:1 /mountdir:$installMountDir}
            $sourceIndex = 1  # After conversion, the new WIM will have only one image
        }
        catch {
            Write-Host "Failed to convert or mount the ESD image: $_" -ForegroundColor Red
            Write-Log -msg "Failed to mount image: $_"
            Pause
            Exit
        }
    }
    else {
        Write-Host "Neither install.wim nor install.esd found. Make sure to mount the correct ISO" -ForegroundColor Red
        Write-Log -msg "Neither install.wim nor install.esd found"
        Pause
        Exit
    }
}
else {
    Write-Host "`nDetails for image: " -NoNewline -ForegroundColor Cyan; Write-Host "$installWimPath"
    Write-Log -msg "Getting image info"
    try {
        # Get image info from install.wim
        $wimInfo = Get-ImageIndex -ImagePath $installWimPath
        if (-not $wimInfo) { 
            Write-Host "Error: Could not retrieve image info from WIM file" -ForegroundColor Red
            Remove-TempFiles
            Pause
            Exit
        }
        # Print image details from install.wim
        foreach ($image in $wimInfo) {
            Write-Host "$($image.Index). $($image.ImageName)"
        }
        # If winEdition is specified, find the index; else prompt user
        if ($winEdition) {
            $matchedImage = $wimInfo | Where-Object { $_.ImageName -ieq $winEdition }
            if ($matchedImage) { $sourceIndex = $matchedImage.Index }
            else { $sourceIndex = 1 }
        }
        else { $sourceIndex = Read-Host -Prompt "`nEnter the index to mount" }
        # Check if the index is valid, print selected "ImageIndex - ImageName"
        $selectedImage = $wimInfo | Where-Object { $_.Index -eq [int]$sourceIndex }
        if ($selectedImage) {
            Write-Host "`nMounting image: " -NoNewline -ForegroundColor Cyan; Write-Host "$sourceIndex. $($selectedImage.ImageName)"
            Write-Log -msg "Mounting image: $sourceIndex. $($selectedImage.ImageName)"
        }

        Invoke-DismFailsafe {Mount-WindowsImage -ImagePath $installWimPath -Index $sourceIndex -Path $installMountDir} {dism /mount-image /imagefile:$installWimPath /index:$sourceIndex /mountdir:$installMountDir}
    }
    catch {
        Write-Host "Failed to mount the image: $_" -ForegroundColor Red
        Write-Log -msg "Failed to mount image: $_"
        Pause
        Exit
    }
}

# Check if wim-mount was successful
if (-not (Test-Path "$installMountDir\Windows")) {
    Write-Host "Error while mounting image. Try again." -ForegroundColor Red
    Write-Log -msg "Mounted image not found. Exiting"
    Remove-TempFiles
    Pause
    Exit 
}

# Resolve Image Info
$WimDetails = Get-WimDetails -MountPath $installMountDir
if (-not $WimDetails -or -not $WimDetails.BuildNumber -or -not $WimDetails.Language) {
    Write-Host "Error: Could not retrieve WIM information from mounted path" -ForegroundColor Red
    Remove-TempFiles
    Pause
    Exit
}
$langCode = $WimDetails.Language; Write-Log -msg "Detected Language: $langCode"
$buildNumber = $WimDetails.BuildNumber; Write-Log -msg "Detected Build Number: $buildNumber"

Write-Host
$DoAppxRemove = Get-ParameterValue -ParameterValue $AppxRemove -DefaultValue $true -Question "Remove unnecessary packages?" -Description "Recommended: Removes bloatware apps"
$DoCapabilitiesRemove = Get-ParameterValue -ParameterValue $CapabilitiesRemove -DefaultValue $true -Question "Remove unnecessary features?" -Description "Recommended: Removes optional Windows features"
$DoOnedriveRemove = Get-ParameterValue -ParameterValue $OnedriveRemove -DefaultValue $true -Question "Remove OneDrive?" -Description "Optional: Completely removes OneDrive"
$DoEDGERemove = Get-ParameterValue -ParameterValue $EDGERemove -DefaultValue $true -Question "Remove Microsoft Edge?" -Description "Optional: Removes Edge components (Breaks Widgets)"
$DoAIRemove = Get-ParameterValue -ParameterValue $AIRemove -DefaultValue $true -Question "Remove AI Components?" -Description "Optional: Removes everything related to AI"
$DoTPMBypass = Get-ParameterValue -ParameterValue $TPMBypass -DefaultValue $false -Question "Bypass TPM check?" -Description "Only if needed for older hardware"
$DoUserFoldersEnable = Get-ParameterValue -ParameterValue $UserFoldersEnable -DefaultValue $true -Question "Enable user folders?" -Description "Recommended: Enables Desktop, Documents, etc."
$DoDriverIntegrate = Get-ParameterValue -ParameterValue $DriverIntegrate -DefaultValue $false -Question "Integrate Intel RST/VMD drivers?" -Description "Optional: Helps with Intel VMD storage controllers"
$DoESDConvert = Get-ParameterValue -ParameterValue $ESDConvert -DefaultValue $false -Question "Compress the ISO?" -Description "Recommended but slow: Reduces ISO file size"
$DoUseOscdimg = Get-ParameterValue -ParameterValue $useOscdimg -DefaultValue $true -Question "Use Oscdimg for ISO creation?" -Description "Recommended: Oscdimg is more reliable"

# --  Extra Tweaks --------------------------------------------------------------
Show-Section "Privacy & Performance Tweaks" "*"
$DoDisableActivityHistory    = Get-ParameterValue -ParameterValue $DisableActivityHistory    -DefaultValue $true  -Question "Disable Activity History?"           -Description "Stops Windows tracking apps/files you open"
$DoDisableLocationTracking   = Get-ParameterValue -ParameterValue $DisableLocationTracking   -DefaultValue $true  -Question "Disable Location Tracking?"           -Description "Blocks system & app access to your location"
$DoDisablePS7Telemetry       = Get-ParameterValue -ParameterValue $DisablePS7Telemetry       -DefaultValue $true  -Question "Disable PowerShell 7 Telemetry?"      -Description "Sets POWERSHELL_TELEMETRY_OPTOUT env variable"
$DoDisableWPBT               = Get-ParameterValue -ParameterValue $DisableWPBT               -DefaultValue $true  -Question "Disable Windows Platform Binary Table?" -Description "Blocks OEM-injected bloatware that reinstalls itself"
$DoDisableCrossDeviceResume  = Get-ParameterValue -ParameterValue $DisableCrossDeviceResume  -DefaultValue $true  -Question "Disable Cross-Device Resume?"          -Description "Stops Windows syncing activity across devices"
$DoDisableHibernation        = Get-ParameterValue -ParameterValue $DisableHibernation        -DefaultValue $false -Question "Disable Hibernation?"                  -Description "Removes hiberfil.sys - recommended for desktops only"
$DoSetServicesManual         = Get-ParameterValue -ParameterValue $SetServicesManual         -DefaultValue $true  -Question "Set Non-Essential Services to Manual?" -Description "Reduces RAM/CPU usage at boot"
$DoDisableBackgroundApps     = Get-ParameterValue -ParameterValue $DisableBackgroundApps     -DefaultValue $true  -Question "Disable Background Apps?"              -Description "Stops UWP apps from running in the background"
$DoDisableFSO                = Get-ParameterValue -ParameterValue $DisableFullscreenOptimizations -DefaultValue $true -Question "Disable Fullscreen Optimizations?" -Description "Can reduce stuttering in some games"
$DoSetClassicRightClick      = Get-ParameterValue -ParameterValue $SetClassicRightClickMenu  -DefaultValue $true  -Question "Set Classic Right-Click Menu (Win11)?" -Description "Restores the full context menu, skipping 'Show more options'"
$DoEnableEndTask             = Get-ParameterValue -ParameterValue $EnableEndTaskRightClick   -DefaultValue $true  -Question "Enable End Task on Taskbar Right-Click?" -Description "Adds End Task directly to taskbar right-click"
$DoDisableExplorerAutoDisc   = Get-ParameterValue -ParameterValue $DisableExplorerAutoDiscovery -DefaultValue $true -Question "Disable Explorer Auto Folder Discovery?" -Description "Stops Explorer auto-switching folder view templates"
$DoSetDarkTheme              = Get-ParameterValue -ParameterValue $SetDarkTheme              -DefaultValue $true  -Question "Set Dark Theme?"                       -Description "Enables system-wide dark mode"
$DoDisableTeredo             = Get-ParameterValue -ParameterValue $DisableTeredo             -DefaultValue $true  -Question "Disable Teredo?"                       -Description "Disables IPv6 tunnel - can reduce latency"
$DoPreferIPv4                = Get-ParameterValue -ParameterValue $PreferIPv4overIPv6        -DefaultValue $true  -Question "Prefer IPv4 over IPv6?"                -Description "Keeps IPv6 but prefers IPv4 connections"
$DoDisableIPv6               = Get-ParameterValue -ParameterValue $DisableIPv6               -DefaultValue $false -Question "Fully Disable IPv6?"                   -Description "Aggressively disables IPv6 - only if you don't need it"
$DoRemoveXboxComponents      = Get-ParameterValue -ParameterValue $RemoveXboxComponents      -DefaultValue $true  -Question "Remove Xbox & Gaming Components?"      -Description "Registry-level removal of Xbox identity/TCUI components"
$DoDisableGameBarProtocols   = Get-ParameterValue -ParameterValue $DisableGameBarProtocols   -DefaultValue $true  -Question "Disable GameBar Protocol Handlers?"    -Description "Redirects ms-gamebar/gamingoverlay URIs to systray.exe"
$DoDisableCopilotExtra       = Get-ParameterValue -ParameterValue $DisableCopilotExtra       -DefaultValue $true  -Question "Extra Copilot Disable Policies?"       -Description "Additional policy keys on top of AI removal"
$DoBlockAdobeNetwork         = Get-ParameterValue -ParameterValue $BlockAdobeNetwork         -DefaultValue $false -Question "Block Adobe Telemetry Network?"         -Description "Requires internet - downloads hosts blocklist"
$DoBlockRazerInstalls        = Get-ParameterValue -ParameterValue $BlockRazerInstalls        -DefaultValue $false -Question "Block Razer Auto-Installs?"             -Description "Prevents Razer software from silently installing via WU"
$DoSetDisplayPerformance     = Get-ParameterValue -ParameterValue $SetDisplayForPerformance  -DefaultValue $true  -Question "Set Display for Performance?"          -Description "Disables all visual effects for maximum snappiness"
$DoEnableDetailedBSoD        = Get-ParameterValue -ParameterValue $EnableDetailedBSoD        -DefaultValue $true  -Question "Enable Detailed BSoD?"                 -Description "Shows full tech info on Blue Screen instead of QR code"
$DoDisableBingSearch         = Get-ParameterValue -ParameterValue $DisableBingSearch         -DefaultValue $true  -Question "Disable Bing Search in Start Menu?"    -Description "Removes Bing web results from Start Menu search"

# -- Security & Privacy Deep-Clean Prompts -------------------------------------
# Security & Privacy Deep-Clean prompt section
$DoDisableWER                = Get-ParameterValue -ParameterValue $DisableWER                -DefaultValue $true  -Question "Disable Windows Error Reporting?"         -Description "Stops crash/error data being sent to Microsoft. Kills WerSvc & QueueReporting task"
$DoDisableDeliveryOpt        = Get-ParameterValue -ParameterValue $DisableDeliveryOptimization -DefaultValue $true -Question "Disable Delivery Optimization (P2P updates)?" -Description "Stops Windows using your PC as an update relay and sending usage stats"
$DoDisableAutoLogger         = Get-ParameterValue -ParameterValue $DisableAutoLogger         -DefaultValue $true  -Question "Disable ETW AutoLogger Sessions?"           -Description "Kills Diagtrack-Listener and SQMLogger persistent background buffers"
$DoDisableCEIP               = Get-ParameterValue -ParameterValue $DisableCEIP               -DefaultValue $true  -Question "Disable CEIP Registry Keys?"               -Description "Removes CEIP/SQMClient opt-in keys (tasks already removed elsewhere)"
$DoRedirectNTP               = Get-ParameterValue -ParameterValue $RedirectNTP               -DefaultValue $true  -Question "Redirect NTP to pool.ntp.org?"             -Description "Stops time sync going to Microsoft's time.windows.com server"
$DoDisableAppAccountInfo     = Get-ParameterValue -ParameterValue $DisableAppAccountInfo     -DefaultValue $true  -Question "Block App Access to Account Info?"         -Description "Denies all apps from reading your Windows account details"
$DoDisableAppContactsCal     = Get-ParameterValue -ParameterValue $DisableAppContactsCalendar -DefaultValue $true -Question "Block App Access to Contacts & Calendar?"  -Description "Denies apps from reading contacts, appointments, call history"
$DoDisableAppCameraMic       = Get-ParameterValue -ParameterValue $DisableAppCameraMic       -DefaultValue $true  -Question "Block App Access to Camera & Microphone?"  -Description "System-wide deny for webcam and mic consent store"
$DoDisableAppMessaging       = Get-ParameterValue -ParameterValue $DisableAppMessaging       -DefaultValue $true  -Question "Block App Access to Messaging?"            -Description "Denies apps from reading SMS/chat and notification listener access"
$DoDisableClipboardHistory   = Get-ParameterValue -ParameterValue $DisableClipboardHistory   -DefaultValue $true  -Question "Disable Clipboard History & Sync?"         -Description "Disables Win+V history and cross-device clipboard via Microsoft account"
$DoDisableSmartScreenExplorer = Get-ParameterValue -ParameterValue $DisableSmartScreenExplorer -DefaultValue $true -Question "Disable SmartScreen for Explorer/Files?"  -Description "Stops file hash checks being sent to Microsoft cloud reputation service"
$DoDisableSmartScreenStore   = Get-ParameterValue -ParameterValue $DisableSmartScreenStore   -DefaultValue $true  -Question "Disable SmartScreen for Store Apps?"       -Description "Separate SmartScreen key controlling Store app URL evaluation"
$DoDisableDefenderMAPS       = Get-ParameterValue -ParameterValue $DisableDefenderMAPS       -DefaultValue $true  -Question "Disable Defender MAPS/Cloud Protection?"   -Description "Stops Defender sending sample/metadata telemetry to Microsoft. Local AV still works"
$DoDisableCloudSearch        = Get-ParameterValue -ParameterValue $DisableCloudSearch        -DefaultValue $true  -Question "Disable Windows Search Cloud Indexing?"    -Description "Blocks OneDrive/online result indexing and location use in Search"
$DoRemoveDeviceCensusTask    = Get-ParameterValue -ParameterValue $RemoveDeviceCensusTask    -DefaultValue $true  -Question "Remove Device Census Scheduled Task?"      -Description "Deletes the task that runs devicecensus.exe to collect hardware telemetry"
$DoRemoveStartupAppTask      = Get-ParameterValue -ParameterValue $RemoveStartupAppTask      -DefaultValue $true  -Question "Remove StartupAppTask?"                   -Description "Removes task that reports startup app timing data to Microsoft"
$DoDisableRemoteAssistance   = Get-ParameterValue -ParameterValue $DisableRemoteAssistance   -DefaultValue $true  -Question "Disable Remote Assistance?"               -Description "Closes the Remote Assistance attack surface (RemoteAccess svc already disabled)"
$DoDisableAutoRun            = Get-ParameterValue -ParameterValue $DisableAutoRun            -DefaultValue $true  -Question "Disable AutoRun/AutoPlay on all drives?"  -Description "Blocks the classic malware vector of auto-executing from USB/optical drives"
$DoDisableDCOM               = Get-ParameterValue -ParameterValue $DisableDCOM               -DefaultValue $false -Question "Disable Network DCOM?"                    -Description "AGGRESSIVE: Blocks remote COM. Only enable on personal workstations with no remote management"

# Comment out the package don't wanna remove
$appxPatternsToRemove = @(
    "Microsoft.Microsoft3DViewer*",             # 3DViewer
    "Microsoft.WindowsAlarms*",                 # Alarms
    "Microsoft.BingNews*",                      # Bing News
    "Microsoft.BingSearch*",                    # Bing Search
    "Microsoft.BingWeather*",                   # Bing Weather (Removing Breaks Widgets)
    "Windows.CBSPreview*",                      # CBS Preview
    "Clipchamp.Clipchamp*",                     # Clipchamp
    "Microsoft.549981C3F5F10*",                 # Cortana
    "MicrosoftWindows.CrossDevice*",            # CrossDevice
    "Microsoft.Windows.DevHome*",               # DevHome
    "MicrosoftCorporationII.MicrosoftFamily*",  # Family
    "Microsoft.WindowsFeedbackHub*",            # FeedbackHub
    "Microsoft.GetHelp*",                       # GetHelp
    "Microsoft.Getstarted*",                    # GetStarted
    "Microsoft.WindowsCommunicationsapps*",     # Mail
    "Microsoft.WindowsMaps*",                   # Maps
    "Microsoft.MixedReality.Portal*",           # MixedReality
    "Microsoft.ZuneMusic*",                     # Music
    "Microsoft.MicrosoftOfficeHub*",            # OfficeHub
    "Microsoft.Office.OneNote*",                # OneNote
    "Microsoft.OutlookForWindows*",             # Outlook
    "Microsoft.MSPaint*",                       # Paint3D(Windows10)
    "Microsoft.People*",                        # People
    "Microsoft.Windows.PeopleExperienceHost*",  # PeopleExperienceHost
    "Microsoft.YourPhone*",                     # Phone
    "Microsoft.PowerAutomateDesktop*",          # PowerAutomate
    "MicrosoftCorporationII.QuickAssist*",      # QuickAssist
    "Microsoft.SkypeApp*",                      # Skype
    "Microsoft.MicrosoftStickyNotes*",          # Sticky Notes
    "Microsoft.MicrosoftSolitaireCollection*",  # SolitaireCollection
    # "Microsoft.WindowsSoundRecorder*",          # SoundRecorder
    "MicrosoftTeams*",                          # Teams_old
    "MSTeams*",                                 # Teams
    "Microsoft.Windows.Teams*",                 # Teams
    "Microsoft.Todos*",                         # Todos
    "Microsoft.ZuneVideo*",                     # Video
    "Microsoft.Wallet*",                        # Wallet
    "Microsoft.GamingApp*",                     # Xbox
    "Microsoft.XboxApp*",                       # Xbox(Win10)
    "Microsoft.XboxGameOverlay*",               # XboxGameOverlay
    "Microsoft.XboxGamingOverlay*",             # XboxGamingOverlay
    # "Microsoft.XboxIdentityProvider*",          # Xbox Identity Provider (Removing Breaks some Xbox Games)
    "Microsoft.XboxSpeechToTextOverlay*",       # XboxSpeechToTextOverlay
    "Microsoft.Xbox.TCUI*"                      # XboxTitleCallableUI
    # "Microsoft.SecHealthUI*"                    # Windows Security (Caution)
)

$capabilitiesToRemove = @(
    "Browser.InternetExplorer*",
    "Internet-Explorer*",
    "App.StepsRecorder*",
    "Language.Handwriting~~~$langCode*",
    # "Language.OCR~~~$langCode*",           # KEPT: Snipping Tool (Microsoft.ScreenSketch) hard-depends on this
    "Language.Speech~~~$langCode*",
    "Language.TextToSpeech~~~$langCode*",
    "Microsoft.Windows.WordPad*",
    "MathRecognizer*",
    "Microsoft.Windows.PowerShell.ISE*",
    # "Hello.Face*",                                # Removing Breaks Windows-Hello
    "Media.WindowsMediaPlayer*"
)

$windowsPackagesToRemove = @(
    "Microsoft-Windows-InternetExplorer-Optional-Package*",
    "Microsoft-Windows-LanguageFeatures-Handwriting-$langCode-Package*",
    # "Microsoft-Windows-LanguageFeatures-OCR-$langCode-Package*",  # KEPT: Snipping Tool hard-depends on this
    "Microsoft-Windows-LanguageFeatures-Speech-$langCode-Package*",
    "Microsoft-Windows-LanguageFeatures-TextToSpeech-$langCode-Package*",
    "Microsoft-Windows-Wallpaper-Content-Extended-FoD-Package*",
    "Microsoft-Windows-WordPad-FoD-Package*",
    "Microsoft-Windows-MediaPlayer-Package*",
    "Microsoft-Windows-TabletPCMath-Package*",
    # "Microsoft-Windows-Hello-Face-Package",       # Removing Breaks Windows-Hello
    "Microsoft-Windows-StepsRecorder-Package*"
)

function Remove-Packages {
    param( [string[]]$Patterns, [string]$SectionTitle, [string]$PackageType, [string]$MountPath, [int]$StartIndex = 1, [int]$TotalCount, [int]$StatusColumn )

    # Package configurations
    $config = @{
        'AppX' = @{
            GetCommand    = { Get-ProvisionedAppxPackage -Path $MountPath }
            FilterProperty = 'PackageName'
            RemoveCommand = { param($item) Remove-ProvisionedAppxPackage -Path $MountPath -PackageName $item.PackageName }
            LogPrefix     = 'AppX package'
        }
        'Capability' = @{
            GetCommand    = { Get-WindowsCapability -Path $MountPath }
            FilterProperty = 'Name'
            RemoveCommand = { param($item) Remove-WindowsCapability -Path $MountPath -Name $item.Name }
            LogPrefix     = 'capability'
        }
        'WindowsPackage' = @{
            GetCommand    = { Get-WindowsPackage -Path $MountPath }
            FilterProperty = 'PackageName'
            RemoveCommand = { param($item) Remove-WindowsPackage -Path $MountPath -PackageName $item.PackageName }
            LogPrefix     = 'Windows package'
        }
    }
    if ($SectionTitle) { Write-Log -msg $SectionTitle }

    $cfg        = $config[$PackageType]
    $filterProp = $cfg.FilterProperty
    $barWidth   = 40
    $removed    = 0
    $notFound   = 0

    for ($i = 0; $i -lt $Patterns.Count; $i++) {
        $pattern     = $Patterns[$i]
        $displayName = $pattern.TrimEnd('*').Split('.')[-1]   # short name only
        $globalIdx   = $StartIndex + $i - 1
        $pct         = [int](($globalIdx / $TotalCount) * 100)
        $filled      = [int]([Math]::Round($pct / 100 * $barWidth))
        $bar         = ("#" * $filled) + ("-" * ($barWidth - $filled))

        # Overwrite same line with progress bar + current item
        $status = "  {0}  {1,3}%  {2}" -f $bar, $pct, $displayName.PadRight(32).Substring(0, [Math]::Min(32, $displayName.Length + [Math]::Max(0, 32 - $displayName.Length)))
        Write-Host ("`r" + $status) -NoNewline -ForegroundColor Magenta

        try {
            $items        = & $cfg.GetCommand | Where-Object { $_.$filterProp -like $pattern }
            $itemsRemoved = 0
            foreach ($item in $items) {
                try   { & $cfg.RemoveCommand $item 2>&1 | Write-Log; $itemsRemoved++ }
                catch { Write-Log -msg "Removing $($cfg.LogPrefix) $($item.$filterProp) failed: $_" }
            }
            if ($itemsRemoved -gt 0) { $removed++ } else { $notFound++ }
        }
        catch {
            Write-Log -msg "Failed to remove $PackageType matching '$pattern': $_"
        }
    }

    # Final bar at 100%
    $finalBar = "#" * $barWidth
    Write-Host ("`r  $finalBar  100%  Done" + (" " * 20)) -ForegroundColor Green
    Write-Host "  v  Removed: $removed  |  Not present: $notFound" -ForegroundColor Green
}

$allPatterns = $appxPatternsToRemove + $capabilitiesToRemove + $windowsPackagesToRemove
$maxLength = ($allPatterns | ForEach-Object { $_.TrimEnd('*').Length } | Measure-Object -Maximum).Maximum
$statusColumn = $maxLength + 18

# Remove AppX Packages
if ($DoAppxRemove) {
    Show-Section "Removing Bloatware Packages" "+"
Remove-Packages -Patterns $appxPatternsToRemove -SectionTitle "Removing provisioned Packages:" -PackageType "AppX" -MountPath $installMountDir -TotalCount $appxPatternsToRemove.Count -StatusColumn $statusColumn
} else {
    Write-Log -msg "Skipped Package Removal"
}

# Remove Capabilities and Windows Packages
if ($DoCapabilitiesRemove) {
    $capabilitiesAndPackagesTotal = $capabilitiesToRemove.Count + $windowsPackagesToRemove.Count
    Remove-Packages -Patterns $capabilitiesToRemove -SectionTitle "Removing Unnecessary Windows Features:" -PackageType "Capability" -MountPath $installMountDir -TotalCount $capabilitiesAndPackagesTotal -StatusColumn $statusColumn
    Remove-Packages -Patterns $windowsPackagesToRemove -SectionTitle "" -PackageType "WindowsPackage" -MountPath $installMountDir -StartIndex ($capabilitiesToRemove.Count + 1) -TotalCount $capabilitiesAndPackagesTotal -StatusColumn $statusColumn
} else {
    Write-Log -msg "Skipped Features Removal"
}

# # Remove Recall (Have conflict with Explorer)
# Write-Host "`nRemoving Recall..."
# Write-Log -msg "Removing Recall"
# dism /image:$installMountDir /Disable-Feature /FeatureName:'Recall' /Remove 2>&1 | Write-Log
# Write-Host "Done"

# # Remove OutlookPWA
# Write-Host "`nRemoving Outlook..." -ForegroundColor Cyan
# Write-Log -msg "Removing OutlookPWA"
# Get-ChildItem "$installMountDir\Windows\WinSxS\amd64_microsoft-windows-outlookpwa*" -Directory | ForEach-Object { Set-OwnAndRemove -Path $_.FullName } 2>&1 | Write-Log
# Write-Host "Done" -ForegroundColor Green

# Setting Permissions
function Enable-Privilege {
    param([ValidateSet('SeAssignPrimaryTokenPrivilege', 'SeAuditPrivilege', 'SeBackupPrivilege', 'SeChangeNotifyPrivilege', 'SeCreateGlobalPrivilege', 'SeCreatePagefilePrivilege', 'SeCreatePermanentPrivilege', 'SeCreateSymbolicLinkPrivilege', 'SeCreateTokenPrivilege', 'SeDebugPrivilege', 'SeEnableDelegationPrivilege', 'SeImpersonatePrivilege', 'SeIncreaseBasePriorityPrivilege', 'SeIncreaseQuotaPrivilege', 'SeIncreaseWorkingSetPrivilege', 'SeLoadDriverPrivilege', 'SeLockMemoryPrivilege', 'SeMachineAccountPrivilege', 'SeManageVolumePrivilege', 'SeProfileSingleProcessPrivilege', 'SeRelabelPrivilege', 'SeRemoteShutdownPrivilege', 'SeRestorePrivilege', 'SeSecurityPrivilege', 'SeShutdownPrivilege', 'SeSyncAgentPrivilege', 'SeSystemEnvironmentPrivilege', 'SeSystemProfilePrivilege', 'SeSystemtimePrivilege', 'SeTakeOwnershipPrivilege', 'SeTcbPrivilege', 'SeTimeZonePrivilege', 'SeTrustedCredManAccessPrivilege', 'SeUndockPrivilege', 'SeUnsolicitedInputPrivilege')]$Privilege, $ProcessId = $pid, [Switch]$Disable)
    $def = @'
    using System;using System.Runtime.InteropServices;public class AdjPriv{[DllImport("advapi32.dll",ExactSpelling=true,SetLastError=true)]internal static extern bool AdjustTokenPrivileges(IntPtr htok,bool disall,ref TokPriv1Luid newst,int len,IntPtr prev,IntPtr relen);[DllImport("advapi32.dll",ExactSpelling=true,SetLastError=true)]internal static extern bool OpenProcessToken(IntPtr h,int acc,ref IntPtr phtok);[DllImport("advapi32.dll",SetLastError=true)]internal static extern bool LookupPrivilegeValue(string host,string name,ref long pluid);[StructLayout(LayoutKind.Sequential,Pack=1)]internal struct TokPriv1Luid{public int Count;public long Luid;public int Attr;}public static bool EnablePrivilege(long processHandle,string privilege,bool disable){var tp=new TokPriv1Luid();tp.Count=1;tp.Attr=disable?0:2;IntPtr htok=IntPtr.Zero;if(!OpenProcessToken(new IntPtr(processHandle),0x28,ref htok))return false;if(!LookupPrivilegeValue(null,privilege,ref tp.Luid))return false;return AdjustTokenPrivileges(htok,false,ref tp,0,IntPtr.Zero,IntPtr.Zero);}}
'@
    (Add-Type $def -PassThru -EA SilentlyContinue)[0]::EnablePrivilege((Get-Process -id $ProcessId).Handle, $Privilege, $Disable)
}
Enable-Privilege SeTakeOwnershipPrivilege | Out-Null

# Remove OneDrive
if ($DoOnedriveRemove) {
    Show-Section "Removing OneDrive" "+"
    Write-Log -msg "Defining OneDrive Setup file paths"
    $oneDriveSetupPath1 = Join-Path -Path $installMountDir -ChildPath 'Windows\System32\OneDriveSetup.exe'
    $oneDriveSetupPath2 = Join-Path -Path $installMountDir -ChildPath 'Windows\SysWOW64\OneDriveSetup.exe'
    # $oneDriveSetupPath3 = (Join-Path -Path $installMountDir -ChildPath 'Windows\WinSxS\*microsoft-windows-onedrive-setup*\OneDriveSetup.exe' | Get-Item -ErrorAction SilentlyContinue).FullName
    # $oneDriveSetupPath4 = (Get-ChildItem "$installMountDir\Windows\WinSxS\amd64_microsoft-windows-onedrive-setup*" -Directory).FullName
    $oneDriveShortcut = Join-Path -Path $installMountDir -ChildPath 'Users\Default\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\OneDrive.lnk'

    Write-Log -msg "Removing OneDrive"
    Set-OwnAndRemove -Path $oneDriveSetupPath1 | Out-Null
    Set-OwnAndRemove -Path $oneDriveSetupPath2 | Out-Null
    # $oneDriveSetupPath3 | Where-Object { $_ } | ForEach-Object { Set-OwnAndRemove -Path $_ } 2>&1 | Write-Log
    # $oneDriveSetupPath4 | Where-Object { $_ } | ForEach-Object { Set-OwnAndRemove -Path $_ } 2>&1 | Write-Log
    Set-OwnAndRemove -Path $oneDriveShortcut | Out-Null

    Show-Bar 100 40 "Complete"
Show-Done "OneDrive removed"
    Write-Log -msg "OneDrive removed successfully"
} else {
    Write-Log -msg "OneDrive removal skipped"
}

# Remove EDGE
if ($DoEDGERemove) {
    Show-Section "Removing Microsoft Edge" "+"
    Write-Log -msg "Removing EDGE"
    
    # Remove Edge using DISM
    Write-Log -msg "Executing DISM - Remove-Edge"
    dism /image:"$installMountDir" /Remove-Edge 2>&1 | Write-Log
    
    # Edge Patterns
    $EDGEpatterns = @(
        "Microsoft.MicrosoftEdge.Stable*",
        "Microsoft.MicrosoftEdgeDevToolsClient*", 
        "Microsoft.Win32WebViewHost*",
        "MicrosoftWindows.Client.WebExperience*"        # Removing Breaks Widgets
    )

    # Remove Edge Packages
    foreach ($pattern in $EDGEpatterns) {
        $matchedPackages = Get-ProvisionedAppxPackage -Path $installMountDir | 
        Where-Object { $_.PackageName -like $pattern }
        foreach ($package in $matchedPackages) {
            Invoke-DismFailsafe {Remove-ProvisionedAppxPackage -Path $installMountDir -PackageName $package.PackageName} {dism /image:$installMountDir /Remove-ProvisionedAppxPackage /PackageName:$($package.PackageName)}
        }
    }

    # Remove WebView2 if not already removed
    Get-WindowsCapability -Path $installMountDir | Where-Object { $_.Name -like "Edge.Webview2.Platform*" } |
        ForEach-Object { Invoke-DismFailsafe {Remove-WindowsCapability -Path $installMountDir -Name $_.Name} {dism /image:$installMountDir /Remove-Capability /CapabilityName:$($_.Name)} }

    Get-WindowsPackage -Path $installMountDir | Where-Object { $_.PackageName -like "Microsoft-Edge-WebView-FOD-Package*" } |
        ForEach-Object { Invoke-DismFailsafe {Remove-WindowsPackage -Path $installMountDir -PackageName $_.PackageName} {dism /image:$installMountDir /Remove-Package /PackageName:$($_.PackageName)} }

    # Modifying reg keys
    try {
        reg load HKLM\zSOFTWARE "$installMountDir\Windows\System32\config\SOFTWARE" 2>&1 | Write-Log
        reg load HKLM\zSYSTEM "$installMountDir\Windows\System32\config\SYSTEM" 2>&1 | Write-Log
        reg load HKLM\zNTUSER "$installMountDir\Users\Default\ntuser.dat" 2>&1 | Write-Log
        reg load HKLM\zDEFAULT "$installMountDir\Windows\System32\config\default" 2>&1 | Write-Log
          
        # Registry operations
        reg delete "HKLM\zSOFTWARE\Microsoft\EdgeUpdate" /f 2>&1 | Write-Log
        reg delete "HKLM\zSOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft Edge" /f 2>&1 | Write-Log
        reg delete "HKLM\zDEFAULT\Software\Microsoft\EdgeUpdate" /f 2>&1 | Write-Log
        reg delete "HKLM\zNTUSER\Software\Microsoft\EdgeUpdate" /f 2>&1 | Write-Log
        reg delete "HKLM\zSOFTWARE\Microsoft\Active Setup\Installed Components\{9459C573-B17A-45AE-9F64-1857B5D58CEE}" /f 2>&1 | Write-Log
        reg delete "HKLM\zSOFTWARE\WOW6432Node\Microsoft\Edge" /f 2>&1 | Write-Log
        reg delete "HKLM\zSOFTWARE\WOW6432Node\Microsoft\EdgeUpdate" /f 2>&1 | Write-Log
        reg delete "HKLM\zSYSTEM\CurrentControlSet\Services\edgeupdate" /f 2>&1 | Write-Log
        reg delete "HKLM\zSYSTEM\ControlSet001\Services\edgeupdate" /f 2>&1 | Write-Log
        reg delete "HKLM\zSYSTEM\CurrentControlSet\Services\edgeupdatem" /f 2>&1 | Write-Log
        reg delete "HKLM\zSYSTEM\ControlSet001\Services\edgeupdatem" /f 2>&1 | Write-Log
        reg delete "HKLM\zSOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft Edge" /f 2>&1 | Write-Log
        reg delete "HKLM\zSOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft Edge Update" /f 2>&1 | Write-Log
        reg add "HKLM\zSOFTWARE\Microsoft\MicrosoftEdge\Main" /v "AllowPrelaunch" /t REG_DWORD /d "1" /f 2>&1 | Write-Log
        reg add "HKLM\zSOFTWARE\Policies\Microsoft\MicrosoftEdge\Main" /v "AllowPrelaunch" /t REG_DWORD /d "1" /f 2>&1 | Write-Log
        reg add "HKLM\zNTUSER\Software\Microsoft\MicrosoftEdge\Main" /v "AllowPrelaunch" /t REG_DWORD /d "1" /f 2>&1 | Write-Log
        reg add "HKLM\zNTUSER\Software\Policies\Microsoft\MicrosoftEdge\Main" /v "AllowPrelaunch" /t REG_DWORD /d "1" /f 2>&1 | Write-Log
        reg add "HKLM\zSOFTWARE\Microsoft\MicrosoftEdge\TabPreloader" /v "AllowTabPreloading" /t REG_DWORD /d "1" /f 2>&1 | Write-Log
        reg add "HKLM\zSOFTWARE\Policies\Microsoft\MicrosoftEdge\TabPreloader" /v "AllowTabPreloading" /t REG_DWORD /d "1" /f 2>&1 | Write-Log
        reg add "HKLM\zNTUSER\Software\Microsoft\MicrosoftEdge\TabPreloader" /v "AllowTabPreloading" /t REG_DWORD /d "1" /f 2>&1 | Write-Log
        reg add "HKLM\zNTUSER\Software\Policies\Microsoft\MicrosoftEdge\TabPreloader" /v "AllowTabPreloading" /t REG_DWORD /d "1" /f 2>&1 | Write-Log
        reg add "HKLM\zSOFTWARE\Policies\Microsoft\EdgeUpdate" /v "UpdateDefault" /t REG_DWORD /d "0" /f 2>&1 | Write-Log
        
        # Disable Edge updates and installation
        $registryKeys = @(
            "HKLM\zSOFTWARE\Microsoft\EdgeUpdate",
            "HKLM\zSOFTWARE\Policies\Microsoft\EdgeUpdate",
            "HKLM\zSOFTWARE\WOW6432Node\Microsoft\EdgeUpdate",
            "HKLM\zNTUSER\Software\Microsoft\EdgeUpdate",
            "HKLM\zNTUSER\Software\Policies\Microsoft\EdgeUpdate"
        )
        foreach ($key in $registryKeys) {
            reg add "$key" /v "DoNotUpdateToEdgeWithChromium" /t REG_DWORD /d "1" /f 2>&1 | Write-Log
            reg add "$key" /v "UpdaterExperimentationAndConfigurationServiceControl" /t REG_DWORD /d "1" /f 2>&1 | Write-Log
            reg add "$key" /v "InstallDefault" /t REG_DWORD /d "1" /f 2>&1 | Write-Log
        }
    }
    catch {
        Write-Log -msg "Error modifying registry: $_"
    }
    finally {
        # Always unload registry hives regardless of errors
        reg unload HKLM\zSOFTWARE 2>&1 | Write-Log
        reg unload HKLM\zSYSTEM 2>&1 | Write-Log
        reg unload HKLM\zNTUSER 2>&1 | Write-Log
        reg unload HKLM\zDEFAULT 2>&1 | Write-Log
    }

    # Remove EDGE files
    Remove-Item -Path "$installMountDir\Program Files\Microsoft\Edge" -Recurse -Force 2>&1 | Write-Log
    Remove-Item -Path "$installMountDir\Program Files\Microsoft\EdgeCore" -Recurse -Force 2>&1 | Write-Log
    Remove-Item -Path "$installMountDir\Program Files\Microsoft\EdgeUpdate" -Recurse -Force 2>&1 | Write-Log
    Remove-Item -Path "$installMountDir\Program Files\Microsoft\EdgeWebView" -Recurse -Force 2>&1 | Write-Log
    Remove-Item -Path "$installMountDir\Program Files (x86)\Microsoft\Edge" -Recurse -Force 2>&1 | Write-Log
    Remove-Item -Path "$installMountDir\Program Files (x86)\Microsoft\EdgeCore" -Recurse -Force 2>&1 | Write-Log
    Remove-Item -Path "$installMountDir\Program Files (x86)\Microsoft\EdgeUpdate" -Recurse -Force 2>&1 | Write-Log
    Remove-Item -Path "$installMountDir\Program Files (x86)\Microsoft\EdgeWebView" -Recurse -Force 2>&1 | Write-Log
    Remove-Item -Path "$installMountDir\ProgramData\Microsoft\EdgeUpdate" -Recurse -Force 2>&1 | Write-Log
    Get-ChildItem "$installMountDir\ProgramData\Microsoft\Windows\AppRepository\Packages\Microsoft.MicrosoftEdge.Stable*" -Directory | ForEach-Object { Set-OwnAndRemove -Path $_.FullName } 2>&1 | Write-Log
    Get-ChildItem "$installMountDir\ProgramData\Microsoft\Windows\AppRepository\Packages\Microsoft.MicrosoftEdgeDevToolsClient*" -Directory | ForEach-Object { Set-OwnAndRemove -Path $_.FullName } 2>&1 | Write-Log
    # Get-ChildItem "$installMountDir\Windows\WinSxS\*microsoft-edge-webview*" -Directory | ForEach-Object { Set-OwnAndRemove -Path $_.FullName } 2>&1 | Write-Log
    Set-OwnAndRemove -Path (Join-Path -Path $installMountDir -ChildPath 'Windows\System32\Microsoft-Edge-WebView') | Out-Null
    Get-Item (Join-Path -Path $installMountDir -ChildPath 'Windows\SystemApps\Microsoft.Win32WebViewHost*') -ErrorAction SilentlyContinue | ForEach-Object { Set-OwnAndRemove -Path $_.FullName | Out-Null }

    # Removing EDGE-Task
    Get-ChildItem -Path "$installMountDir\Windows\System32\Tasks\MicrosoftEdge*" | Where-Object { $_ } | ForEach-Object { Set-OwnAndRemove -Path $_ } 2>&1 | Write-Log
    
    # For Windows 10 (Legacy EDGE)
    if ($buildNumber -lt 22000) {
        Get-ChildItem -Path "$installMountDir\Windows\SystemApps\Microsoft.MicrosoftEdge*" | Where-Object { $_ } | ForEach-Object { Set-OwnAndRemove -Path $_ } 2>&1 | Write-Log
    }
    
    Show-Bar 100 40 "Complete"
Show-Done "Edge removed"
    Write-Log -msg "Microsoft Edge removal completed"
} else {
    Write-Log -msg "Edge removal cancelled"
}

# Remove AI components
if ($buildNumber -ge 22000) {
    if ($DoAIRemove) {
        Show-Section "Removing AI Components" "+"
        Write-Log -msg "Removing AI components"
        
        # Remove AI Packages
        $AIpatterns = @(
            "Microsoft.Windows.Copilot*",
            "Microsoft.Copilot*"
        )
        foreach ($pattern in $AIpatterns) {
            $matchedPackages = Get-ProvisionedAppxPackage -Path $installMountDir | 
            Where-Object { $_.PackageName -like $pattern }
            foreach ($package in $matchedPackages) {
                Invoke-DismFailsafe {Remove-ProvisionedAppxPackage -Path $installMountDir -PackageName $package.PackageName} {dism /image:$installMountDir /Remove-ProvisionedAppxPackage /PackageName:$($package.PackageName)}
            }
        }

        # Disable AI DLLs
        $dllfiles = @('System32', 'SysWOW64') | ForEach-Object {
            Join-Path $installMountDir "Windows\$_\Windows.AI.MachineLearning.dll"
            Join-Path $installMountDir "Windows\$_\Windows.AI.MachineLearning.Preview.dll"
        }
        $dllfiles += Join-Path $installMountDir "Windows\System32\SettingsHandlers_Copilot.dll"
        $dllfiles | Where-Object { Test-Path $_ } | ForEach-Object {
            Set-Ownership -Path $_ | Out-Null
            try { Rename-Item $_ ($_ + ".bak") -Force -ErrorAction Stop 2>&1 | Write-Log }
            catch {
                Write-Log -msg "Rename failed for $_. Attempting to delete..."
                Set-OwnAndRemove -Path $_ 2>&1 | Write-Log
            }
        }

        # Modifying reg keys
        try {
            reg load HKLM\zSOFTWARE "$installMountDir\Windows\System32\config\SOFTWARE" 2>&1 | Write-Log
            reg load HKLM\zSYSTEM "$installMountDir\Windows\System32\config\SYSTEM" 2>&1 | Write-Log
            reg load HKLM\zNTUSER "$installMountDir\Users\Default\ntuser.dat" 2>&1 | Write-Log

            # Registry operations
            reg add "HKLM\zSOFTWARE\Policies\Microsoft\Windows\Explorer" /v "DisableSearchBoxSuggestions" /t REG_DWORD /d "1" /f 2>&1 | Write-Log
            # Disable AI in Notepad
            reg add "HKLM\zSOFTWARE\Policies\WindowsNotepad" /v "DisableAIFeatures" /t REG_DWORD /d "1" /f 2>&1 | Write-Log
            # Disable AI in Paint
            reg add "HKLM\zSOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Paint" /v "DisableCocreator" /t REG_DWORD /d "1" /f 2>&1 | Write-Log
            reg add "HKLM\zSOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Paint" /v "DisableImageCreator" /t REG_DWORD /d "1" /f 2>&1 | Write-Log
            # Disable AI in other apps
            reg add "HKLM\zSOFTWARE\Policies\Microsoft\Windows\AppPrivacy" /v "LetAppsAccessSystemAIModels" /t REG_DWORD /d "2" /f 2>&1 | Write-Log
            reg add "HKLM\zSOFTWARE\Policies\Microsoft\Windows\AppPrivacy" /v "LetAppsAccessGenerativeAI" /t REG_DWORD /d "2" /f 2>&1 | Write-Log
            # Disable AI access
            reg add "HKLM\zSOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\generativeAI" /v "Value" /t REG_SZ /d "Deny" /f 2>&1 | Write-Log
            # Disable AI in Edge
            reg add "HKLM\zSOFTWARE\Policies\Microsoft\Edge" /v "HubsSidebarEnabled" /t REG_DWORD /d "0" /f 2>&1 | Write-Log
            reg add "HKLM\zSOFTWARE\Policies\Microsoft\Edge" /v "CopilotPageContext" /t REG_DWORD /d "0" /f 2>&1 | Write-Log
            reg add "HKLM\zSOFTWARE\Policies\Microsoft\Edge" /v "CopilotCDPPageContext" /t REG_DWORD /d "0" /f 2>&1 | Write-Log
            # Disable AI in Search
            reg add "HKLM\zSOFTWARE\Policies\Microsoft\Windows\WindowsAI" /v "DisableClickToDo" /t REG_DWORD /d "1" /f 2>&1 | Write-Log
            # Disable WSAIFabricSvc Service on first logon
            reg add "HKLM\zSOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce" /v "DisableWSAIFabricSvc" /t REG_SZ /d 'reg add "HKLM\SYSTEM\CurrentControlSet\Services\WSAIFabricSvc" /v "Start" /t REG_DWORD /d "4" /f'
            reg add "HKLM\zSOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce" /v "StopWSAIFabricSvc" /t REG_SZ /d "net stop WSAIFabricSvc"
            # Hide AI components from Settings
            reg add "HKLM\zSOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" /v "SettingsPageVisibility" /t REG_SZ /d "hide:aicomponents" /f 2>&1 | Write-Log
            # Disable AI from Explorer
            reg add "HKLM\zNTUSER\Software\Microsoft\Windows\CurrentVersion\WindowsCopilot" /v "AllowCopilotRuntime" /t REG_DWORD /d "0" /f 2>&1 | Write-Log
            reg add "HKLM\zNTUSER\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband\AuxilliaryPins" /v "CopilotPWAPin" /t REG_DWORD /d "0" /f 2>&1 | Write-Log
            reg add "HKLM\zNTUSER\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband\AuxilliaryPins" /v "RecallPin" /t REG_DWORD /d "0" /f 2>&1 | Write-Log
            # Disable Copilot and Recall system-wide
            reg add "HKLM\zSOFTWARE\Policies\Microsoft\Windows\WindowsCopilot" /v "TurnOffWindowsCopilot" /t REG_DWORD /d "1" /f 2>&1 | Write-Log
            reg add "HKLM\zSOFTWARE\Policies\Microsoft\Windows\WindowsAI" /v "DisableAIDataAnalysis" /t REG_DWORD /d "1" /f 2>&1 | Write-Log
            reg add "HKLM\zSOFTWARE\Policies\Microsoft\Windows\WindowsAI" /v "AllowRecallEnablement" /t REG_DWORD /d "0" /f 2>&1 | Write-Log
            reg add "HKLM\zSOFTWARE\Policies\Microsoft\Windows\WindowsAI" /v "TurnOffSavingSnapshots" /t REG_DWORD /d "1" /f 2>&1 | Write-Log
            reg add "HKLM\zSOFTWARE\Policies\Microsoft\Windows\WindowsAI" /v "DisableSettingsAgent" /t REG_DWORD /d "1" /f 2>&1 | Write-Log
            reg add "HKLM\zSOFTWARE\Microsoft\Windows\Shell\Copilot" /v "IsCopilotAvailable" /t REG_DWORD /d "0" /f 2>&1 | Write-Log
            reg add "HKLM\zSOFTWARE\Microsoft\Windows\Shell\Copilot" /v "CopilotDisabledReason" /t REG_SZ /d "FeatureIsDisabled" /f 2>&1 | Write-Log
            # Disable Copilot and Recall for New Users
            reg add "HKLM\zNTUSER\Software\Policies\Microsoft\Windows\WindowsCopilot" /v "TurnOffWindowsCopilot" /t REG_DWORD /d "1" /f 2>&1 | Write-Log
            reg add "HKLM\zNTUSER\Software\Policies\Microsoft\Windows\WindowsAI" /v "DisableAIDataAnalysis" /t REG_DWORD /d "1" /f 2>&1 | Write-Log
            reg add "HKLM\zNTUSER\Software\Policies\Microsoft\Windows\WindowsAI" /v "AllowRecallEnablement" /t REG_DWORD /d "0" /f 2>&1 | Write-Log
            reg add "HKLM\zNTUSER\Software\Policies\Microsoft\Windows\WindowsAI" /v "TurnOffSavingSnapshots" /t REG_DWORD /d "1" /f 2>&1 | Write-Log
            reg add "HKLM\zNTUSER\Software\Policies\Microsoft\Windows\WindowsAI" /v "DisableSettingsAgent" /t REG_DWORD /d "1" /f 2>&1 | Write-Log
            reg add "HKLM\zNTUSER\Software\Microsoft\Windows\Shell\Copilot" /v "IsCopilotAvailable" /t REG_DWORD /d "0" /f 2>&1 | Write-Log
            reg add "HKLM\zNTUSER\Software\Microsoft\Windows\Shell\Copilot" /v "CopilotDisabledReason" /t REG_SZ /d "FeatureIsDisabled" /f 2>&1 | Write-Log
            # Remove AI Tasks
            reg delete "HKLM\zSOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree\Microsoft\Windows\WindowsAI" /f 2>&1 | Write-Log
            Set-OwnAndRemove -Path "$installMountDir\Windows\System32\Tasks\Microsoft\Windows\WindowsAI" | Out-Null
            # Disable Recall on first logon
            reg add "HKLM\zSOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce" /v "DisableRecall" /t REG_SZ /d "dism.exe /online /disable-feature /FeatureName:recall" /f 2>&1 | Write-Log
        }
        catch {
            Write-Log -msg "Error modifying registry: $_"
        }
        finally {
            # Always unload registry hives regardless of errors
            reg unload HKLM\zSOFTWARE 2>&1 | Write-Log
            reg unload HKLM\zSYSTEM 2>&1 | Write-Log
            reg unload HKLM\zNTUSER 2>&1 | Write-Log
        }
        Show-Bar 100 40 "Complete"
Show-Done "AI components removed"
        Write-Log -msg "AI Components removal completed"
    } else {
        Write-Log -msg "AI Components removal skipped"
    }
}


# +==========================================================================+
# |   Generate Post-Install Script (registry/services tweaks at runtime)    |
# +==========================================================================+
# Rather than baking all registry and service tweaks into the offline WIM
# (which risks CONFIG_INITIALIZATION_FAILED / 0x67 boot errors), we write a
# postinstall.ps1 that Windows Setup auto-runs on first boot via
# SetupComplete.cmd placed in $OEM$\$$\Setup\Scripts\.
# Tweaks that MUST be offline (package removal, OOBE bypass, task deletion)
# are still applied directly to the WIM above.

Show-Section "Building Post-Install Script" "*"
Write-Log -msg "Generating postinstall.ps1"

# $OEM$ structure: sources\$OEM$\$$\Setup\Scripts\SetupComplete.cmd
#   -> Windows copies $$ contents to %SystemRoot% during setup
#   -> Scripts\SetupComplete.cmd is run by setup after first boot, as SYSTEM
#   -> It launches our postinstall.ps1 which applies all registry/service tweaks

$oemScriptsDir = Join-Path $destinationPath 'sources\$OEM$\$$\Setup\Scripts'
New-Item -ItemType Directory -Path $oemScriptsDir -Force | Out-Null

# SetupComplete.cmd - runs our PS script as SYSTEM on first boot, then deletes itself
$setupCompleteContent = @'
@echo off
PowerShell.exe -NoProfile -ExecutionPolicy Bypass -File "%SystemRoot%\Setup\Scripts\postinstall.ps1" >> "%SystemRoot%\Setup\Scripts\postinstall_log.txt" 2>&1
del "%SystemRoot%\Setup\Scripts\SetupComplete.cmd"
'@
$setupCompleteContent | Set-Content -Path (Join-Path $oemScriptsDir 'SetupComplete.cmd') -Encoding ASCII

# ---------------------------------------------------------------------------
# Build the postinstall.ps1 content
# Each tweak block mirrors what was previously applied offline, but uses
# live HKLM\SOFTWARE, HKCU (via Default User hive load), HKLM\SYSTEM, etc.
# ---------------------------------------------------------------------------

# Collect the enabled-tweak flags so postinstall.ps1 knows what to apply
$piFlags = @{
    DisableSponsoredApps        = $true
    DisableTelemetry            = $true
    DisableMouseAcceleration    = $true
    DisableMeetNow              = $true
    DisableAdsAndStuffs         = $true
    DisableBitlocker            = $true
    DisableOneDriveJunk         = $true
    DisableGameDVR              = $true
    DisableGamebarPopup         = $true
    ActivityHistory             = $DoDisableActivityHistory
    LocationTracking            = $DoDisableLocationTracking
    PS7Telemetry                = $DoDisablePS7Telemetry
    WPBT                        = $DoDisableWPBT
    CrossDeviceResume           = $DoDisableCrossDeviceResume
    Hibernation                 = $DoDisableHibernation
    ServicesManual              = $DoSetServicesManual
    BackgroundApps              = $DoDisableBackgroundApps
    FullscreenOptimizations     = $DoDisableFSO
    ClassicRightClick           = $DoSetClassicRightClick
    EndTask                     = $DoEnableEndTask
    ExplorerAutoDiscovery       = $DoDisableExplorerAutoDisc
    DarkTheme                   = $DoSetDarkTheme
    Teredo                      = $DoDisableTeredo
    PreferIPv4                  = $DoPreferIPv4
    DisableIPv6                 = $DoDisableIPv6
    XboxComponents              = $DoRemoveXboxComponents
    GameBarProtocols            = $DoDisableGameBarProtocols
    CopilotExtra                = $DoDisableCopilotExtra
    BlockAdobe                  = $DoBlockAdobeNetwork
    BlockRazer                  = $DoBlockRazerInstalls
    DisplayPerformance          = $DoSetDisplayPerformance
    DetailedBSoD                = $DoEnableDetailedBSoD
    BingSearch                  = $DoDisableBingSearch
    WER                         = $DoDisableWER
    DeliveryOptimization        = $DoDisableDeliveryOpt
    AutoLogger                  = $DoDisableAutoLogger
    CEIP                        = $DoDisableCEIP
    RedirectNTP                 = $DoRedirectNTP
    AppAccountInfo              = $DoDisableAppAccountInfo
    AppContactsCal              = $DoDisableAppContactsCal
    AppCameraMic                = $DoDisableAppCameraMic
    AppMessaging                = $DoDisableAppMessaging
    ClipboardHistory            = $DoDisableClipboardHistory
    SmartScreenExplorer         = $DoDisableSmartScreenExplorer
    SmartScreenStore            = $DoDisableSmartScreenStore
    DefenderMAPS                = $DoDisableDefenderMAPS
    CloudSearch                 = $DoDisableCloudSearch
    RemoteAssistance            = $DoDisableRemoteAssistance
    AutoRun                     = $DoDisableAutoRun
    DCOM                        = $DoDisableDCOM
}

# Serialize flags into the script as a hashtable literal
$flagsBlock = "`$flags = @{`n"
foreach ($kv in $piFlags.GetEnumerator()) {
    $val = if ($kv.Value) { '$true' } else { '$false' }
    $flagsBlock += "    $($kv.Key) = $val`n"
}
$flagsBlock += "}`n"

$postInstallScript = @'
# NiiX Debloat - Post-Install Script
# Auto-generated by niixdebloat.ps1 - runs on first boot via SetupComplete.cmd
# Applies registry and service tweaks that are safer to set on a live system.

$logPath = "$env:SystemRoot\Setup\Scripts\postinstall_log.txt"
function Write-PI { param([string]$m) $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'; Add-Content $logPath "$ts - $m" }

Write-PI "=== NiiX Post-Install Tweaks starting ==="

# ---------------------------------------------------------------------------
# Helper: load Default user hive so HKCU tweaks apply to new users too
# ---------------------------------------------------------------------------
$defaultHivePath = "$env:SystemRoot\Users\Default\NTUSER.DAT"
$hiveLoaded = $false
if (Test-Path $defaultHivePath) {
    reg load HKLM\piNTUSER $defaultHivePath 2>&1 | ForEach-Object { Write-PI $_ }
    $hiveLoaded = $true
    Write-PI "Default user hive loaded"
}

# Apply to both live HKCU and Default user hive
function Set-UserReg {
    param([string]$SubKey, [string]$Name, [string]$Type, [string]$Value)
    reg add "HKCU\$SubKey"          /v $Name /t $Type /d $Value /f 2>&1 | ForEach-Object { Write-PI $_ }
    if ($hiveLoaded) {
        reg add "HKLM\piNTUSER\$SubKey" /v $Name /t $Type /d $Value /f 2>&1 | ForEach-Object { Write-PI $_ }
    }
}

'@

# Append the flags block
$postInstallScript += $flagsBlock

$postInstallScript += @'

# ---------------------------------------------------------------------------
# Disable Sponsored Apps / Content Delivery Manager
# ---------------------------------------------------------------------------
if ($flags.DisableSponsoredApps) {
    Write-PI "Disabling Sponsored Apps"
    reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /v "OemPreInstalledAppsEnabled" /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "PreInstalledAppsEnabled"         REG_DWORD "0"
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SilentInstalledAppsEnabled"      REG_DWORD "0"
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SubscribedContentEnabled"        REG_DWORD "0"
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SubscribedContent-310093Enabled" REG_DWORD "0"
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SubscribedContent-338388Enabled" REG_DWORD "0"
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SubscribedContent-338389Enabled" REG_DWORD "0"
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SubscribedContent-338393Enabled" REG_DWORD "0"
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SubscribedContent-353694Enabled" REG_DWORD "0"
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SubscribedContent-353696Enabled" REG_DWORD "0"
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SubscribedContent-338387Enabled" REG_DWORD "0"
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "ContentDeliveryAllowed"          REG_DWORD "0"
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "PreInstalledAppsEverEnabled"     REG_DWORD "0"
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SoftLandingEnabled"              REG_DWORD "0"
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SystemPaneSuggestionsEnabled"    REG_DWORD "0"
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\CloudContent" /v "DisableWindowsConsumerFeatures" /t REG_DWORD /d "1" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Microsoft\PolicyManager\current\device\Start" /v "ConfigureStartPins" /t REG_SZ /d '{"pinnedList": [{}]}' /f 2>&1 | ForEach-Object { Write-PI $_ }
}

# ---------------------------------------------------------------------------
# Disable Telemetry
# ---------------------------------------------------------------------------
if ($flags.DisableTelemetry) {
    Write-PI "Disabling Telemetry"
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\DataCollection" /v "AllowTelemetry" /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    Set-UserReg "Software\Microsoft\Personalization\Settings"                       "AcceptedPrivacyPolicy"                       REG_DWORD "0"
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\Privacy"                 "TailoredExperiencesWithDiagnosticDataEnabled" REG_DWORD "0"
    Set-UserReg "Software\Microsoft\Speech_OneCore\Settings\OnlineSpeechPrivacy"    "HasAccepted"                                 REG_DWORD "0"
    Set-UserReg "Software\Microsoft\InputPersonalization"                           "RestrictImplicitInkCollection"               REG_DWORD "1"
    Set-UserReg "Software\Microsoft\InputPersonalization"                           "RestrictImplicitTextCollection"              REG_DWORD "1"
    Set-UserReg "Software\Microsoft\InputPersonalization\TrainedDataStore"          "HarvestContacts"                             REG_DWORD "0"
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo"         "Enabled"                                     REG_DWORD "0"
    reg add "HKLM\SYSTEM\CurrentControlSet\Services\dmwappushservice" /v "Start" /t REG_DWORD /d "4" /f 2>&1 | ForEach-Object { Write-PI $_ }
}

# ---------------------------------------------------------------------------
# Disable Mouse Acceleration
# ---------------------------------------------------------------------------
if ($flags.DisableMouseAcceleration) {
    Write-PI "Disabling Mouse Acceleration"
    Set-UserReg "Control Panel\Mouse" "MouseSpeed"      REG_SZ "0"
    Set-UserReg "Control Panel\Mouse" "MouseThreshold1" REG_SZ "0"
    Set-UserReg "Control Panel\Mouse" "MouseThreshold2" REG_SZ "0"
}

# ---------------------------------------------------------------------------
# Disable Meet Now / Online Tips
# ---------------------------------------------------------------------------
if ($flags.DisableMeetNow) {
    Write-PI "Disabling Meet Now"
    reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" /v "HideSCAMeetNow"  /t REG_DWORD /d "1" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" /v "AllowOnlineTips" /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
}

# ---------------------------------------------------------------------------
# Disable Ads and Stuffs
# ---------------------------------------------------------------------------
if ($flags.DisableAdsAndStuffs) {
    Write-PI "Disabling Ads and Stuffs"
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo"             "Enabled"                   REG_DWORD "0"
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\CloudContent" /v "DisableConsumerAccountStateContent" /t REG_DWORD /d "1" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\CloudContent" /v "DisableCloudOptimizedContent"       /t REG_DWORD /d "1" /f 2>&1 | ForEach-Object { Write-PI $_ }
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"           "Start_IrisRecommendations" REG_DWORD "0"
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Feeds" /v "EnableFeeds"  /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel" "{2cc5ca98-6485-489a-920e-b3e88a6ccce3}" REG_DWORD "1"
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search" /v "AllowCortana" /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    Set-UserReg "Control Panel\Desktop"                                                 "MenuShowDelay"             REG_SZ "200"
    reg add "HKLM\SOFTWARE\Policies\Microsoft\MRT"    /v "DontOfferThroughWUAU"         /t REG_DWORD /d "1" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Teams"  /v "DisableInstallation"          /t REG_DWORD /d "1" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Mail" /v "PreventRun"     /t REG_DWORD /d "1" /f 2>&1 | ForEach-Object { Write-PI $_ }
}

# ---------------------------------------------------------------------------
# Disable BitLocker auto-encryption
# ---------------------------------------------------------------------------
if ($flags.DisableBitlocker) {
    Write-PI "Disabling BitLocker auto-encryption"
    reg add "HKLM\SYSTEM\CurrentControlSet\Control\BitLocker" /v "PreventDeviceEncryption" /t REG_DWORD /d "1" /f 2>&1 | ForEach-Object { Write-PI $_ }
}

# ---------------------------------------------------------------------------
# Remove OneDrive junk
# ---------------------------------------------------------------------------
if ($flags.DisableOneDriveJunk) {
    Write-PI "Removing OneDrive junk"
    reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "OneDriveSetup" /f 2>&1 | ForEach-Object { Write-PI $_ }
    if ($hiveLoaded) { reg delete "HKLM\piNTUSER\Software\Microsoft\Windows\CurrentVersion\Run" /v "OneDriveSetup" /f 2>&1 | ForEach-Object { Write-PI $_ } }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\OneDrive" /v "DisableLibrariesDefaultSaveToOneDrive" /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\OneDrive" /v "DisableFileSyncNGSC"                  /t REG_DWORD /d "1" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\OneDrive"         /v "KFMBlockOptIn"                        /t REG_DWORD /d "1" /f 2>&1 | ForEach-Object { Write-PI $_ }
}

# ---------------------------------------------------------------------------
# Disable GameDVR
# ---------------------------------------------------------------------------
if ($flags.DisableGameDVR) {
    Write-PI "Disabling GameDVR"
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\GameDVR" "AppCaptureEnabled" REG_DWORD "0"
    Set-UserReg "System\GameConfigStore"                             "GameDVR_Enabled"   REG_DWORD "0"
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\GameDVR" /v "AllowGameDVR" /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SYSTEM\CurrentControlSet\Services\BcastDVRUserService"   /v "Start" /t REG_DWORD /d "4" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SYSTEM\CurrentControlSet\Services\GameBarPresenceWriter" /v "Start" /t REG_DWORD /d "4" /f 2>&1 | ForEach-Object { Write-PI $_ }
}

# ---------------------------------------------------------------------------
# Disable GameBar popup
# ---------------------------------------------------------------------------
if ($flags.DisableGamebarPopup) {
    Write-PI "Disabling GameBar popup"
    Set-UserReg "Software\Microsoft\GameBar" "AutoGameModeEnabled"          REG_DWORD "0"
    Set-UserReg "Software\Microsoft\GameBar" "UseNexusForGameBarEnabled"    REG_DWORD "0"
    Set-UserReg "Software\Microsoft\GameBar" "ShowStartupPanel"             REG_DWORD "0"
}

# ---------------------------------------------------------------------------
# 1. Disable Activity History
# ---------------------------------------------------------------------------
if ($flags.ActivityHistory) {
    Write-PI "Disabling Activity History"
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\System" /v "EnableActivityFeed"   /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\System" /v "PublishUserActivities" /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\System" /v "UploadUserActivities"  /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
}

# ---------------------------------------------------------------------------
# 2. Disable Location Tracking
# ---------------------------------------------------------------------------
if ($flags.LocationTracking) {
    Write-PI "Disabling Location Tracking"
    reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location" /v "Value" /t REG_SZ /d "Deny" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Sensor\Overrides\{BFA794E4-F964-4FDB-90F6-51056BFE4B44}" /v "SensorPermissionState" /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SYSTEM\CurrentControlSet\Services\lfsvc\Service\Configuration" /v "Status" /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
}

# ---------------------------------------------------------------------------
# 3. Disable PowerShell 7 Telemetry
# ---------------------------------------------------------------------------
if ($flags.PS7Telemetry) {
    Write-PI "Disabling PS7 Telemetry"
    reg add "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v "POWERSHELL_TELEMETRY_OPTOUT" /t REG_SZ /d "1" /f 2>&1 | ForEach-Object { Write-PI $_ }
    [System.Environment]::SetEnvironmentVariable("POWERSHELL_TELEMETRY_OPTOUT", "1", "Machine")
}

# ---------------------------------------------------------------------------
# 4. Disable WPBT
# ---------------------------------------------------------------------------
if ($flags.WPBT) {
    Write-PI "Disabling WPBT"
    reg add "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager" /v "DisableWpbtExecution" /t REG_DWORD /d "1" /f 2>&1 | ForEach-Object { Write-PI $_ }
}

# ---------------------------------------------------------------------------
# 5. Disable Cross-Device Resume
# ---------------------------------------------------------------------------
if ($flags.CrossDeviceResume) {
    Write-PI "Disabling Cross-Device Resume"
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\System" /v "EnableCdp" /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\CDP" "CdpSessionUserAuthzPolicy"        REG_DWORD "0"
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\CDP" "EnableRemotelyLaunchedActivations" REG_DWORD "0"
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\CDP" "RomeSdkChannelUserAuthzPolicy"     REG_DWORD "0"
}

# ---------------------------------------------------------------------------
# 6. Disable Hibernation
# ---------------------------------------------------------------------------
if ($flags.Hibernation) {
    Write-PI "Disabling Hibernation"
    reg add "HKLM\SYSTEM\CurrentControlSet\Control\Power" /v "HibernateEnabled"       /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SYSTEM\CurrentControlSet\Control\Power" /v "HibernateEnabledDefault" /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    & powercfg /hibernate off 2>&1 | ForEach-Object { Write-PI $_ }
}

# ---------------------------------------------------------------------------
# 7. Set Non-Essential Services to Manual
# ---------------------------------------------------------------------------
if ($flags.ServicesManual) {
    Write-PI "Setting non-essential services to Manual"
    $manualServices = @(
        "ALG","AppMgmt","AppReadiness","Appinfo","AssignedAccessManagerSvc","AxInstSV",
        "BDESVC","BTAGService","CDPSvc","COMSysApp","CertPropSvc","CscService",
        "DevQueryBroker","DeviceAssociationService","DeviceInstall","DisplayEnhancementService",
        "EFS","EapHost","FDResPub","FrameServer","FrameServerMonitor","GraphicsPerfSvc",
        "HvHost","IKEEXT","InstallService","IpxlatCfgSvc","KtmRm","LicenseManager",
        "LxpSvc","MSDTC","MSiSCSI","MapsBroker","McpManagementService","NaturalAuthentication",
        "NcaSvc","NcbService","NcdAutoSetup","NetSetupSvc","Netman","NlaSvc","PcaSvc",
        "PeerDistSvc","PerfHost","PhoneSvc","PolicyAgent","PrintNotify","PushToInstall",
        "QWAVE","RasAuto","RasMan","RetailDemo","RmSvc","RpcLocator","SCPolicySvc",
        "SCardSvr","SDRSVC","SEMgrSvc","SNMPTRAP","SNMPTrap","SSDPSRV","ScDeviceEnum",
        "SensorDataService","SensorService","SensrSvc","SessionEnv","SharedAccess",
        "SmsRouter","SstpSvc","StiSvc","StorSvc","TapiSrv","TermService","TieringEngineService",
        "TokenBroker","TroubleshootingSvc","TrustedInstaller","UmRdpService","UsoSvc","VSS",
        "VaultSvc","W32Time","WEPHOSTSVC","WFDSConMgrSvc","WMPNetworkSvc","WManSvc",
        "WPDBusEnum","WSAIFabricSvc","WalletService","WarpJITSvc","WbioSrvc","WdiServiceHost",
        "WdiSystemHost","WebClient","Wecsvc","WerSvc","WiaRpc","WinRM","WpcMonSvc","WpnService",
        "XblAuthManager","XblGameSave","XboxGipSvc","XboxNetApiSvc","autotimesvc","bthserv",
        "camsvc","cloudidsvc","dcsvc","defragsvc","diagsvc","dmwappushservice","dot3svc",
        "edgeupdate","edgeupdatem","fdPHost","fhsvc","hidserv","icssvc","lfsvc","lltdsvc",
        "lmhosts","netprofm","perceptionsimulation","pla","seclogon","smphost","svsvc","swprv",
        "upnphost","vds","vmicguestinterface","vmicheartbeat","vmickvpexchange","vmicrdv",
        "vmicshutdown","vmictimesync","vmicvmsession","vmicvss","wbengine","wcncsvc",
        "webthreatdefsvc","wercplsupport","wisvc","wlidsvc","wlpasvc","wmiApSrv",
        "workfolderssvc","wuauserv"
    )
    $disabledServices = @("AppVClient","AssignedAccessManagerSvc","DialogBlockingService",
        "NetTcpPortSharing","RemoteAccess","RemoteRegistry","UevAgentService","shpamsvc","ssh-agent","tzautoupdate")
    foreach ($svc in $manualServices) {
        reg add "HKLM\SYSTEM\CurrentControlSet\Services\$svc" /v "Start" /t REG_DWORD /d "3" /f 2>&1 | ForEach-Object { Write-PI $_ }
    }
    foreach ($svc in $disabledServices) {
        reg add "HKLM\SYSTEM\CurrentControlSet\Services\$svc" /v "Start" /t REG_DWORD /d "4" /f 2>&1 | ForEach-Object { Write-PI $_ }
    }
    reg add "HKLM\SYSTEM\CurrentControlSet\Services\DiagTrack" /v "Start" /t REG_DWORD /d "4" /f 2>&1 | ForEach-Object { Write-PI $_ }
    foreach ($svc in @("InventorySvc","PcaSvc","StorSvc","UsoSvc","WpnService","camsvc","WSAIFabricSvc")) {
        reg add "HKLM\SYSTEM\CurrentControlSet\Services\$svc" /v "Start" /t REG_DWORD /d "3" /f 2>&1 | ForEach-Object { Write-PI $_ }
    }
}

# ---------------------------------------------------------------------------
# 9. Disable Background Apps
# ---------------------------------------------------------------------------
if ($flags.BackgroundApps) {
    Write-PI "Disabling Background Apps"
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\BackgroundAccessApplications" "GlobalUserDisabled" REG_DWORD "1"
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" /v "LetAppsRunInBackground" /t REG_DWORD /d "2" /f 2>&1 | ForEach-Object { Write-PI $_ }
}

# ---------------------------------------------------------------------------
# 10. Disable Fullscreen Optimizations
# ---------------------------------------------------------------------------
if ($flags.FullscreenOptimizations) {
    Write-PI "Disabling Fullscreen Optimizations"
    Set-UserReg "System\GameConfigStore" "GameDVR_DXGIHonorFSEWindowsCompatible" REG_DWORD "1"
    Set-UserReg "System\GameConfigStore" "GameDVR_FSEBehavior"                   REG_DWORD "2"
    Set-UserReg "System\GameConfigStore" "GameDVR_FSEBehaviorMode"               REG_DWORD "2"
    Set-UserReg "System\GameConfigStore" "GameDVR_HonorUserFSEBehaviorMode"      REG_DWORD "1"
}

# ---------------------------------------------------------------------------
# 12. Classic Right-Click Menu (Win11)
# ---------------------------------------------------------------------------
if ($flags.ClassicRightClick) {
    Write-PI "Setting Classic Right-Click Menu"
    Set-UserReg "Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32" "" REG_SZ ""
}

# ---------------------------------------------------------------------------
# 15. Enable End Task on Taskbar Right-Click
# ---------------------------------------------------------------------------
if ($flags.EndTask) {
    Write-PI "Enabling End Task on Taskbar"
    reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\TaskbarDeveloperSettings" /v "TaskbarEndTask" /t REG_DWORD /d "1" /f 2>&1 | ForEach-Object { Write-PI $_ }
}

# ---------------------------------------------------------------------------
# 16. Disable Explorer Auto Folder Discovery
# ---------------------------------------------------------------------------
if ($flags.ExplorerAutoDiscovery) {
    Write-PI "Disabling Explorer Auto Folder Discovery"
    Set-UserReg "Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\Bags\AllFolders\Shell" "FolderType" REG_SZ "NotSpecified"
}

# ---------------------------------------------------------------------------
# 19. Dark Theme
# ---------------------------------------------------------------------------
if ($flags.DarkTheme) {
    Write-PI "Setting Dark Theme"
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" "AppsUseLightTheme"    REG_DWORD "0"
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" "SystemUsesLightTheme" REG_DWORD "0"
    reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" /v "AppsUseLightTheme"    /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" /v "SystemUsesLightTheme" /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
}

# ---------------------------------------------------------------------------
# 20. Disable Teredo
# ---------------------------------------------------------------------------
if ($flags.Teredo) {
    Write-PI "Disabling Teredo"
    reg add "HKLM\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters" /v "DisabledComponents" /t REG_DWORD /d "1" /f 2>&1 | ForEach-Object { Write-PI $_ }
    & netsh interface teredo set state disabled 2>&1 | ForEach-Object { Write-PI $_ }
}

# ---------------------------------------------------------------------------
# 21. Prefer IPv4 over IPv6
# ---------------------------------------------------------------------------
if ($flags.PreferIPv4 -and -not $flags.DisableIPv6) {
    Write-PI "Preferring IPv4 over IPv6"
    reg add "HKLM\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters" /v "DisabledComponents" /t REG_DWORD /d "32" /f 2>&1 | ForEach-Object { Write-PI $_ }
}

# ---------------------------------------------------------------------------
# 22. Fully Disable IPv6
# ---------------------------------------------------------------------------
if ($flags.DisableIPv6) {
    Write-PI "Fully disabling IPv6"
    reg add "HKLM\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters" /v "DisabledComponents" /t REG_DWORD /d "255" /f 2>&1 | ForEach-Object { Write-PI $_ }
}

# ---------------------------------------------------------------------------
# 23. Remove Xbox Registry Components
# ---------------------------------------------------------------------------
if ($flags.XboxComponents) {
    Write-PI "Removing Xbox registry components"
    reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" /v "XboxAutosave" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\GameDVR" /v "AllowGameDVR" /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Microsoft\PolicyManager\default\ApplicationManagement\AllowGameDVR" /v "value" /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    foreach ($svc in @("XblAuthManager","XblGameSave","XboxGipSvc","XboxNetApiSvc")) {
        reg add "HKLM\SYSTEM\CurrentControlSet\Services\$svc" /v "Start" /t REG_DWORD /d "4" /f 2>&1 | ForEach-Object { Write-PI $_ }
    }
}

# ---------------------------------------------------------------------------
# 24. Disable GameBar Protocol Handlers
# ---------------------------------------------------------------------------
if ($flags.GameBarProtocols) {
    Write-PI "Disabling GameBar protocol handlers"
    foreach ($proto in @("ms-gamebar","ms-gamebarservices","ms-gamingoverlay")) {
        reg add "HKLM\SOFTWARE\Classes\$proto"                    /v "NoOpenWith"    /t REG_SZ /d ""                                 /f 2>&1 | ForEach-Object { Write-PI $_ }
        reg add "HKLM\SOFTWARE\Classes\$proto"                    /ve                /t REG_SZ /d "URL:$proto"                        /f 2>&1 | ForEach-Object { Write-PI $_ }
        reg add "HKLM\SOFTWARE\Classes\$proto"                    /v "URL Protocol"  /t REG_SZ /d ""                                 /f 2>&1 | ForEach-Object { Write-PI $_ }
        reg add "HKLM\SOFTWARE\Classes\$proto\shell\open\command" /ve                /t REG_SZ /d "%SystemRoot%\System32\systray.exe" /f 2>&1 | ForEach-Object { Write-PI $_ }
    }
    reg add "HKLM\SOFTWARE\Microsoft\WindowsRuntime\Server\Windows.Gaming.GameBar.Internal.PresenceWriterServer" /v "ExePath" /t REG_SZ /d "%SystemRoot%\System32\systray.exe" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Microsoft\WindowsRuntime\ActivatableClassId\Windows.Gaming.GameBar.PresenceServer.Internal.PresenceWriter" /v "ActivationType" /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
}

# ---------------------------------------------------------------------------
# 25. Extra Copilot Disable Policies
# ---------------------------------------------------------------------------
if ($flags.CopilotExtra) {
    Write-PI "Applying extra Copilot disable policies"
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot" /v "TurnOffWindowsCopilot" /t REG_DWORD /d "1" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Edge"                   /v "HubsSidebarEnabled"    /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Edge"                   /v "CopilotPageContext"     /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    Set-UserReg "Software\Policies\Microsoft\Windows\WindowsCopilot"       "TurnOffWindowsCopilot"  REG_DWORD "1"
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\WindowsCopilot" "AllowCopilotRuntime"    REG_DWORD "0"
    reg add "HKLM\SOFTWARE\Microsoft\Windows\Shell\Copilot" /v "IsCopilotAvailable"    /t REG_DWORD /d "0"                  /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Microsoft\Windows\Shell\Copilot" /v "CopilotDisabledReason" /t REG_SZ    /d "FeatureIsDisabled"  /f 2>&1 | ForEach-Object { Write-PI $_ }
}

# ---------------------------------------------------------------------------
# 26. Block Adobe Telemetry Network
# ---------------------------------------------------------------------------
if ($flags.BlockAdobe) {
    Write-PI "Blocking Adobe telemetry network"
    $hostsPath = "$env:SystemRoot\System32\drivers\etc\hosts"
    try {
        $adobeHosts = (Invoke-WebRequest "https://github.com/Ruddernation-Designs/Adobe-URL-Block-List/raw/refs/heads/master/hosts" -UseBasicParsing -ErrorAction Stop).Content
        $existing   = if (Test-Path $hostsPath) { Get-Content $hostsPath -Raw } else { "" }
        $newEntries = ($adobeHosts -split "`n") | Where-Object { $_ -notmatch "^127\.0\.0\.1\s+localhost" -and $_.Trim() }
        "$existing`n# Adobe URL Block List`n$($newEntries -join "`n")" | Set-Content $hostsPath -Encoding UTF8
        Write-PI "Adobe network block applied"
    } catch { Write-PI "Adobe hosts download failed: $_" }
}

# ---------------------------------------------------------------------------
# 27. Block Razer Auto-Installs
# ---------------------------------------------------------------------------
if ($flags.BlockRazer) {
    Write-PI "Blocking Razer auto-installs"
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\DeviceInstall\Settings" /v "AllowInstallationOfMatchingDeviceSetupClasses" /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\DriverSearching"         /v "SearchOrderConfig"                           /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SYSTEM\CurrentControlSet\Services\Razer Chroma SDK Service" /v "Start" /t REG_DWORD /d "4" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SYSTEM\CurrentControlSet\Services\RazerCentralService"      /v "Start" /t REG_DWORD /d "4" /f 2>&1 | ForEach-Object { Write-PI $_ }
}

# ---------------------------------------------------------------------------
# 28. Set Display for Performance
# ---------------------------------------------------------------------------
if ($flags.DisplayPerformance) {
    Write-PI "Setting Display for Performance"
    Set-UserReg "Control Panel\Desktop"                                            "DragFullWindows"       REG_SZ    "0"
    Set-UserReg "Control Panel\Desktop"                                            "MenuShowDelay"         REG_SZ    "200"
    Set-UserReg "Control Panel\Desktop\WindowMetrics"                              "MinAnimate"            REG_SZ    "0"
    Set-UserReg "Control Panel\Keyboard"                                           "KeyboardDelay"         REG_DWORD "0"
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"      "ListviewAlphaSelect"   REG_DWORD "0"
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"      "ListviewShadow"        REG_DWORD "0"
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"      "TaskbarAnimations"     REG_DWORD "0"
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects" "VisualFXSetting"       REG_DWORD "3"
    Set-UserReg "Software\Microsoft\Windows\DWM"                                   "EnableAeroPeek"        REG_DWORD "0"
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"      "TaskbarMn"             REG_DWORD "0"
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"      "TaskbarDa"             REG_DWORD "0"
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"      "ShowTaskViewButton"    REG_DWORD "0"
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\Search"                 "SearchboxTaskbarMode"  REG_DWORD "0"
}

# ---------------------------------------------------------------------------
# 29. Enable Detailed BSoD
# ---------------------------------------------------------------------------
if ($flags.DetailedBSoD) {
    Write-PI "Enabling Detailed BSoD"
    reg add "HKLM\SYSTEM\CurrentControlSet\Control\CrashControl" /v "DisplayParameters" /t REG_DWORD /d "1" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SYSTEM\CurrentControlSet\Control\CrashControl" /v "DisableEmoticon"   /t REG_DWORD /d "1" /f 2>&1 | ForEach-Object { Write-PI $_ }
}

# ---------------------------------------------------------------------------
# 32. Disable Bing Search in Start Menu
# ---------------------------------------------------------------------------
if ($flags.BingSearch) {
    Write-PI "Disabling Bing Search in Start Menu"
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\Search" "BingSearchEnabled" REG_DWORD "0"
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\Search" "CortanaConsent"    REG_DWORD "0"
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search" /v "DisableWebSearch"                            /t REG_DWORD /d "1" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search" /v "ConnectedSearchUseWeb"                       /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search" /v "ConnectedSearchUseWebOverMeteredConnections"  /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
}

# ---------------------------------------------------------------------------
# A1. Disable Windows Error Reporting
# ---------------------------------------------------------------------------
if ($flags.WER) {
    Write-PI "Disabling Windows Error Reporting"
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Error Reporting" /v "Disabled"               /t REG_DWORD /d "1" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Error Reporting" /v "DontSendAdditionalData" /t REG_DWORD /d "1" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Error Reporting" /v "LoggingDisabled"        /t REG_DWORD /d "1" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Microsoft\Windows\Windows Error Reporting"          /v "Disabled"               /t REG_DWORD /d "1" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Microsoft\Windows\Windows Error Reporting"          /v "DontSendAdditionalData" /t REG_DWORD /d "1" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Microsoft\Windows\Windows Error Reporting"          /v "LoggingDisabled"        /t REG_DWORD /d "1" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SYSTEM\CurrentControlSet\Services\WerSvc"        /v "Start" /t REG_DWORD /d "4" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SYSTEM\CurrentControlSet\Services\wercplsupport" /v "Start" /t REG_DWORD /d "4" /f 2>&1 | ForEach-Object { Write-PI $_ }
}

# ---------------------------------------------------------------------------
# A3. Disable Delivery Optimization
# ---------------------------------------------------------------------------
if ($flags.DeliveryOptimization) {
    Write-PI "Disabling Delivery Optimization"
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\DeliveryOptimization" /v "DODownloadMode"             /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\DeliveryOptimization" /v "DODisallowCacheServer"      /t REG_DWORD /d "1" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\DeliveryOptimization" /v "SystemSettingsDownloadMode" /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\DeliveryOptimization\Settings" /v "DownloadMode" /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SYSTEM\CurrentControlSet\Services\DoSvc" /v "Start" /t REG_DWORD /d "4" /f 2>&1 | ForEach-Object { Write-PI $_ }
}

# ---------------------------------------------------------------------------
# A4. Disable ETW AutoLogger Sessions
# ---------------------------------------------------------------------------
if ($flags.AutoLogger) {
    Write-PI "Disabling ETW AutoLogger sessions"
    reg add "HKLM\SYSTEM\CurrentControlSet\Control\WMI\Autologger\AutoLogger-Diagtrack-Listener" /v "Start"          /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SYSTEM\CurrentControlSet\Control\WMI\Autologger\AutoLogger-Diagtrack-Listener" /v "EnableProperty" /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SYSTEM\CurrentControlSet\Control\WMI\Autologger\SQMLogger"                     /v "Start"          /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SYSTEM\CurrentControlSet\Control\WMI\Autologger\SQMLogger"                     /v "EnableProperty" /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SYSTEM\CurrentControlSet\Control\WMI\Autologger\WiFiSession"                   /v "Start"          /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
}

# ---------------------------------------------------------------------------
# A5. Disable CEIP
# ---------------------------------------------------------------------------
if ($flags.CEIP) {
    Write-PI "Disabling CEIP"
    reg add "HKLM\SOFTWARE\Policies\Microsoft\SQMClient\Windows" /v "CEIPEnable" /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Microsoft\SQMClient\Windows"          /v "CEIPEnable" /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Messenger\Client"  /v "CEIP"       /t REG_DWORD /d "2" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Microsoft\VSCommon\15.0\SQM"          /v "OptIn"      /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
}

# ---------------------------------------------------------------------------
# B5. Redirect NTP
# ---------------------------------------------------------------------------
if ($flags.RedirectNTP) {
    Write-PI "Redirecting NTP to pool.ntp.org"
    reg add "HKLM\SOFTWARE\Policies\Microsoft\W32time\Parameters"          /v "NtpServer" /t REG_SZ /d "0.pool.ntp.org,1.pool.ntp.org,2.pool.ntp.org,3.pool.ntp.org" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\W32time\Parameters"          /v "Type"      /t REG_SZ /d "NTP" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SYSTEM\CurrentControlSet\Services\W32Time\Parameters"    /v "NtpServer" /t REG_SZ /d "0.pool.ntp.org,0x9" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SYSTEM\CurrentControlSet\Services\W32Time\Parameters"    /v "Type"      /t REG_SZ /d "NTP" /f 2>&1 | ForEach-Object { Write-PI $_ }
}

# ---------------------------------------------------------------------------
# C1. Block App Access to Account Info
# ---------------------------------------------------------------------------
if ($flags.AppAccountInfo) {
    Write-PI "Blocking app access to Account Info"
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" /v "LetAppsAccessAccountInfo" /t REG_DWORD /d "2" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\userAccountInformation" /v "Value" /t REG_SZ /d "Deny" /f 2>&1 | ForEach-Object { Write-PI $_ }
}

# ---------------------------------------------------------------------------
# C2. Block App Access to Contacts & Calendar
# ---------------------------------------------------------------------------
if ($flags.AppContactsCal) {
    Write-PI "Blocking app access to Contacts & Calendar"
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" /v "LetAppsAccessContacts"    /t REG_DWORD /d "2" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" /v "LetAppsAccessCalendar"    /t REG_DWORD /d "2" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" /v "LetAppsAccessPhone"       /t REG_DWORD /d "2" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" /v "LetAppsAccessCallHistory" /t REG_DWORD /d "2" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" /v "LetAppsAccessEmail"       /t REG_DWORD /d "2" /f 2>&1 | ForEach-Object { Write-PI $_ }
    foreach ($store in @("contacts","appointments","phoneCall","phoneCallHistory","email")) {
        reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\$store" /v "Value" /t REG_SZ /d "Deny" /f 2>&1 | ForEach-Object { Write-PI $_ }
    }
}

# ---------------------------------------------------------------------------
# C3. Block App Access to Camera & Microphone
# ---------------------------------------------------------------------------
if ($flags.AppCameraMic) {
    Write-PI "Blocking app access to Camera & Mic"
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" /v "LetAppsAccessCamera"     /t REG_DWORD /d "2" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" /v "LetAppsAccessMicrophone" /t REG_DWORD /d "2" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\webcam"     /v "Value" /t REG_SZ /d "Deny" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone" /v "Value" /t REG_SZ /d "Deny" /f 2>&1 | ForEach-Object { Write-PI $_ }
}

# ---------------------------------------------------------------------------
# C4. Block App Access to Messaging
# ---------------------------------------------------------------------------
if ($flags.AppMessaging) {
    Write-PI "Blocking app access to Messaging"
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" /v "LetAppsAccessMessaging"           /t REG_DWORD /d "2" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" /v "LetAppsAccessNotifications"       /t REG_DWORD /d "2" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" /v "LetAppsAccessUnpublishedCalendar" /t REG_DWORD /d "2" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\chat"                     /v "Value" /t REG_SZ /d "Deny" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\userNotificationListener" /v "Value" /t REG_SZ /d "Deny" /f 2>&1 | ForEach-Object { Write-PI $_ }
}

# ---------------------------------------------------------------------------
# C5. Disable Clipboard History & Cross-Device Sync
# ---------------------------------------------------------------------------
if ($flags.ClipboardHistory) {
    Write-PI "Disabling Clipboard History & Sync"
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\System" /v "AllowClipboardHistory"     /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\System" /v "AllowCrossDeviceClipboard" /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    Set-UserReg "Software\Microsoft\Clipboard" "EnableClipboardHistory"        REG_DWORD "0"
    Set-UserReg "Software\Microsoft\Clipboard" "CloudClipboardAutomaticUpload" REG_DWORD "0"
}

# ---------------------------------------------------------------------------
# D1. Disable SmartScreen for Explorer
# ---------------------------------------------------------------------------
if ($flags.SmartScreenExplorer) {
    Write-PI "Disabling SmartScreen for Explorer"
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\System"               /v "EnableSmartScreen"          /t REG_DWORD /d "0"        /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer"       /v "SmartScreenEnabled"         /t REG_SZ    /d "Off"      /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\SmartScreen" /v "ConfigureAppInstallControl" /t REG_SZ    /d "Anywhere" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\SmartScreen" /v "EnableSmartScreen"          /t REG_DWORD /d "0"        /f 2>&1 | ForEach-Object { Write-PI $_ }
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\AppHost" "EnableWebContentEvaluation" REG_DWORD "0"
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\AppHost" "PreventOverride"            REG_DWORD "0"
}

# ---------------------------------------------------------------------------
# D2. Disable SmartScreen for Store Apps
# ---------------------------------------------------------------------------
if ($flags.SmartScreenStore) {
    Write-PI "Disabling SmartScreen for Store"
    reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\AppHost" /v "EnableWebContentEvaluation" /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\AppHost" /v "PreventOverride"            /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
}

# ---------------------------------------------------------------------------
# D3. Disable Defender MAPS / Cloud Protection
# ---------------------------------------------------------------------------
if ($flags.DefenderMAPS) {
    Write-PI "Disabling Defender MAPS/Cloud Protection"
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Spynet" /v "SpynetReporting"         /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Spynet" /v "SubmitSamplesConsent"    /t REG_DWORD /d "2" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Spynet" /v "DisableBlockAtFirstSeen" /t REG_DWORD /d "1" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Microsoft\Windows Defender\Spynet"          /v "SpynetReporting"         /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Microsoft\Windows Defender\Spynet"          /v "SubmitSamplesConsent"    /t REG_DWORD /d "2" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Microsoft\Windows Defender\Features"        /v "TamperProtection"        /t REG_DWORD /d "4" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\MpEngine" /v "MpCloudBlockLevel"     /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
}

# ---------------------------------------------------------------------------
# D4. Disable Windows Search Cloud Indexing
# ---------------------------------------------------------------------------
if ($flags.CloudSearch) {
    Write-PI "Disabling Windows Search Cloud Indexing"
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search" /v "AllowCloudSearch"          /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search" /v "AllowSearchToUseLocation"  /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search" /v "ConnectedSearchPrivacy"    /t REG_DWORD /d "3" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search" /v "ConnectedSearchSafeSearch" /t REG_DWORD /d "3" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search" /v "DoNotUseWebResults"        /t REG_DWORD /d "1" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search" /v "EnableDynamicContentInWSB" /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\SearchSettings" "IsDeviceSearchHistoryEnabled" REG_DWORD "0"
}

# ---------------------------------------------------------------------------
# F1. Disable Remote Assistance
# ---------------------------------------------------------------------------
if ($flags.RemoteAssistance) {
    Write-PI "Disabling Remote Assistance"
    reg add "HKLM\SYSTEM\CurrentControlSet\Control\Remote Assistance"       /v "fAllowToGetHelp"         /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SYSTEM\CurrentControlSet\Control\Remote Assistance"       /v "fAllowFullControl"       /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services" /v "fAllowToGetHelp"         /t REG_DWORD /d "0" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\System"               /v "DisableRemoteAssistance" /t REG_DWORD /d "1" /f 2>&1 | ForEach-Object { Write-PI $_ }
}

# ---------------------------------------------------------------------------
# F2. Disable AutoRun / AutoPlay
# ---------------------------------------------------------------------------
if ($flags.AutoRun) {
    Write-PI "Disabling AutoRun & AutoPlay"
    reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" /v "NoDriveTypeAutoRun"     /t REG_DWORD /d "255" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" /v "NoAutorun"              /t REG_DWORD /d "1"   /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" /v "NoAutoplayfornonVolume" /t REG_DWORD /d "1"   /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\Explorer"                /v "NoAutoplayfornonVolume" /t REG_DWORD /d "1"   /f 2>&1 | ForEach-Object { Write-PI $_ }
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer" "NoDriveTypeAutoRun" REG_DWORD "255"
    Set-UserReg "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer" "NoAutorun"          REG_DWORD "1"
    reg add "HKLM\SYSTEM\CurrentControlSet\Services\ShellHWDetection" /v "Start" /t REG_DWORD /d "4" /f 2>&1 | ForEach-Object { Write-PI $_ }
}

# ---------------------------------------------------------------------------
# F3. Disable Network DCOM
# ---------------------------------------------------------------------------
if ($flags.DCOM) {
    Write-PI "Disabling Network DCOM"
    reg add "HKLM\SOFTWARE\Microsoft\Ole" /v "EnableDCOM"                /t REG_SZ    /d "N" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Microsoft\Ole" /v "EnableRemoteConnect"       /t REG_SZ    /d "N" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Microsoft\Ole" /v "LegacyAuthenticationLevel" /t REG_DWORD /d "6" /f 2>&1 | ForEach-Object { Write-PI $_ }
    reg add "HKLM\SOFTWARE\Microsoft\Ole" /v "LegacyImpersonationLevel"  /t REG_DWORD /d "2" /f 2>&1 | ForEach-Object { Write-PI $_ }
}

# ---------------------------------------------------------------------------
# Unload Default user hive
# ---------------------------------------------------------------------------
if ($hiveLoaded) {
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Start-Sleep -Seconds 2
    reg unload HKLM\piNTUSER 2>&1 | ForEach-Object { Write-PI $_ }
    Write-PI "Default user hive unloaded"
}

Write-PI "=== NiiX Post-Install Tweaks complete ==="
'@

# Write postinstall.ps1 to the OEM scripts directory
$postInstallPath = Join-Path $oemScriptsDir 'postinstall.ps1'
$postInstallScript | Set-Content -Path $postInstallPath -Encoding UTF8
Write-Log -msg "postinstall.ps1 written to $postInstallPath"

Show-Bar 100 40 "Post-install script ready"
Show-Done "postinstall.ps1 + SetupComplete.cmd placed in ISO"
Write-Log -msg "OEM post-install scripts created successfully"

# +==========================================================================+
# |   Minimal Offline Registry Tweaks (WIM)                                 |
# |   Only tweaks that MUST be baked offline:                               |
# |   - Sponsored apps / OOBE bypass (needed before first-run experience)   |
# |   - Scheduled task removal (tasks run before SetupComplete)             |
# |   - BitLocker pre-boot prevention                                       |
# |   All other tweaks are handled by postinstall.ps1 at first boot.        |
# +==========================================================================+

Show-Section "Loading Registry Hives (Minimal Offline Tweaks)" "*"
Write-Log -msg "Loading registry for minimal offline tweaks"
reg load HKLM\zCOMPONENTS "$installMountDir\Windows\System32\config\COMPONENTS" 2>&1 | Write-Log
reg load HKLM\zDEFAULT "$installMountDir\Windows\System32\config\default" 2>&1 | Write-Log
reg load HKLM\zNTUSER "$installMountDir\Users\Default\ntuser.dat" 2>&1 | Write-Log
reg load HKLM\zSOFTWARE "$installMountDir\Windows\System32\config\SOFTWARE" 2>&1 | Write-Log
reg load HKLM\zSYSTEM "$installMountDir\Windows\System32\config\SYSTEM" 2>&1 | Write-Log

Set-Ownership -Registry @("zSOFTWARE\Microsoft\Windows\CurrentVersion\Communications", "zSOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks", "zSOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree\Microsoft\Windows", "zSOFTWARE\Microsoft\WindowsRuntime\Server\Windows.Gaming.GameBar.Internal.PresenceWriterServer") | Out-Null

Show-Bar 100 40 "Hives mounted"
Show-Done "Registry hives loaded"

Show-Section "Applying Minimal Offline Registry Tweaks" "*"
$statusColumn = 60

# -- Disable Sponsored Apps (CDM keys must be offline so OOBE doesn't re-enable them) --
Write-Host -NoNewline ("  Disabling Sponsored Apps (offline)".PadRight($statusColumn)) -ForegroundColor DarkGray
reg add "HKLM\zSOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /v "OemPreInstalledAppsEnabled" /t REG_DWORD /d "0" /f 2>&1 | Write-Log
reg add "HKLM\zNTUSER\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /v "PreInstalledAppsEnabled"   /t REG_DWORD /d "0" /f 2>&1 | Write-Log
reg add "HKLM\zNTUSER\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /v "SilentInstalledAppsEnabled" /t REG_DWORD /d "0" /f 2>&1 | Write-Log
reg add "HKLM\zSOFTWARE\Policies\Microsoft\Windows\CloudContent" /v "DisableWindowsConsumerFeatures" /t REG_DWORD /d "1" /f 2>&1 | Write-Log
reg add "HKLM\zSOFTWARE\Microsoft\PolicyManager\current\device\Start" /v "ConfigureStartPins" /t REG_SZ /d '{\"pinnedList\": [{}]}' /f 2>&1 | Write-Log
reg add "HKLM\zNTUSER\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /v "ContentDeliveryAllowed"      /t REG_DWORD /d "0" /f 2>&1 | Write-Log
reg add "HKLM\zNTUSER\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /v "PreInstalledAppsEverEnabled"  /t REG_DWORD /d "0" /f 2>&1 | Write-Log
reg delete "HKLM\zNTUSER\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager\Subscriptions" /f 2>&1 | Write-Log
reg delete "HKLM\zNTUSER\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager\SuggestedApps" /f 2>&1 | Write-Log
Write-Host " v" -ForegroundColor Green

# -- Disable Telemetry (policy must be 0 before first boot) --
Write-Host -NoNewline ("  Disabling Telemetry (offline)".PadRight($statusColumn)) -ForegroundColor DarkGray
reg add "HKLM\zSOFTWARE\Policies\Microsoft\Windows\DataCollection" /v "AllowTelemetry" /t REG_DWORD /d "0" /f 2>&1 | Write-Log
reg add "HKLM\zSYSTEM\ControlSet001\Services\dmwappushservice" /v "Start" /t REG_DWORD /d "4" /f 2>&1 | Write-Log
Write-Host " v" -ForegroundColor Green

# -- OOBE Bypass (must be offline) --
Write-Host -NoNewline ("  Tweaking OOBE Settings".PadRight($statusColumn)) -ForegroundColor DarkGray
reg add "HKLM\zSOFTWARE\Policies\Microsoft\Windows\OOBE" /v "DisablePrivacyExperience"    /t REG_DWORD /d "1" /f 2>&1 | Write-Log
reg add "HKLM\zSOFTWARE\Microsoft\Windows\CurrentVersion\OOBE" /v "BypassNRO"             /t REG_DWORD /d "1" /f 2>&1 | Write-Log
reg add "HKLM\zSOFTWARE\Microsoft\Windows\CurrentVersion\OOBE" /v "BypassNROGatherOptions" /t REG_DWORD /d "1" /f 2>&1 | Write-Log
if (Test-Path -Path $autounattendXmlPath) {
    Write-Log -msg "Copying Autounattend.xml"
    Copy-Item -Path $autounattendXmlPath -Destination $destinationPath -Force
} else {
    Write-Warning "Autounattend.xml not found at $autounattendXmlPath"
    Write-Log -msg "Warning: Autounattend.xml not found at $autounattendXmlPath"
}
Write-Host " v" -ForegroundColor Green

# -- Prevent orchestrator auto-installs (runs before SetupComplete) --
Write-Host -NoNewline ("  Blocking orchestrator auto-installs".PadRight($statusColumn)) -ForegroundColor DarkGray
reg delete "HKLM\zSOFTWARE\Microsoft\WindowsUpdate\Orchestrator\UScheduler_Oobe\DevHomeUpdate" /f 2>&1 | Write-Log
reg add "HKLM\zSOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Orchestrator\UScheduler\DevHomeUpdate" /v "workCompleted" /t REG_DWORD /d "1" /f 2>&1 | Write-Log
reg delete "HKLM\zSOFTWARE\Microsoft\WindowsUpdate\Orchestrator\UScheduler_Oobe\OutlookUpdate" /f 2>&1 | Write-Log
reg add "HKLM\zSOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Orchestrator\UScheduler\OutlookUpdate" /v "workCompleted" /t REG_DWORD /d "1" /f 2>&1 | Write-Log
reg add "HKLM\zSOFTWARE\Microsoft\Windows\CurrentVersion\Communications" /v "ConfigureChatAutoInstall" /t REG_DWORD /d "0" /f 2>&1 | Write-Log
reg add "HKLM\zSOFTWARE\Policies\Microsoft\Windows\Windows Chat" /v "ChatIcon" /t REG_DWORD /d "3" /f 2>&1 | Write-Log
Write-Host " v" -ForegroundColor Green

# -- Disable BitLocker pre-boot (must be offline) --
Write-Host -NoNewline ("  Disabling BitLocker Encryption (offline)".PadRight($statusColumn)) -ForegroundColor DarkGray
reg add "HKLM\zSYSTEM\ControlSet001\Control\BitLocker" /v "PreventDeviceEncryption" /t REG_DWORD /d "1" /f 2>&1 | Write-Log
Write-Host " v" -ForegroundColor Green

# -- Remove Scheduled Tasks (must be offline - some run before SetupComplete.cmd) --
Write-Host -NoNewline ("  Removing Scheduled Tasks (offline)".PadRight($statusColumn)) -ForegroundColor DarkGray
$win24H2 = (Get-ItemProperty -Path 'Registry::HKLM\zSOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name DisplayVersion -ErrorAction SilentlyContinue).DisplayVersion -eq '24H2'
if ($win24H2) {
    reg delete "HKLM\zSOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks\{780E487D-C62F-4B55-AF84-0E38116AFE07}" /f 2>&1 | Write-Log
    reg delete "HKLM\zSOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks\{FD607F42-4541-418A-B812-05C32EBA8626}" /f 2>&1 | Write-Log
    reg delete "HKLM\zSOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks\{E4FED5BC-D567-4044-9642-2EDADF7DE108}" /f 2>&1 | Write-Log
    Set-OwnAndRemove -Path "$installMountDir\Windows\System32\Tasks\Microsoft\Windows\Customer Experience Improvement Program" | Out-Null
    reg delete "HKLM\zSOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks\{E292525C-72F1-482C-8F35-C513FAA98DAE}" /f 2>&1 | Write-Log
    Set-OwnAndRemove -Path "$installMountDir\Windows\System32\Tasks\Microsoft\Windows\Application Experience\ProgramDataUpdater" | Out-Null
    reg delete "HKLM\zSOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks\{3047C197-66F1-4523-BA92-6C955FEF9E4E}" /f 2>&1 | Write-Log
    reg delete "HKLM\zSOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks\{A0C71CB8-E8F0-498A-901D-4EDA09E07FF4}" /f 2>&1 | Write-Log
    Set-OwnAndRemove -Path "$installMountDir\Windows\System32\Tasks\Microsoft\Windows\Application Experience\Microsoft Compatibility Appraiser" | Out-Null
} else {
    reg delete "HKLM\zSOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks\{4738DE7A-BCC1-4E2D-B1B0-CADB044BFA81}" /f 2>&1 | Write-Log
    reg delete "HKLM\zSOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks\{6FAC31FA-4A85-4E64-BFD5-2154FF4594B3}" /f 2>&1 | Write-Log
    reg delete "HKLM\zSOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks\{FC931F16-B50A-472E-B061-B6F79A71EF59}" /f 2>&1 | Write-Log
    reg delete "HKLM\zSOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree\Microsoft\Windows\Customer Experience Improvement Program" /f 2>&1 | Write-Log
    Set-OwnAndRemove -Path "$installMountDir\Windows\System32\Tasks\Microsoft\Windows\Customer Experience Improvement Program" | Out-Null
    reg delete "HKLM\zSOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks\{0671EB05-7D95-4153-A32B-1426B9FE61DB}" /f 2>&1 | Write-Log
    Set-OwnAndRemove -Path "$installMountDir\Windows\System32\Tasks\Microsoft\Windows\Application Experience\ProgramDataUpdater" | Out-Null
    reg delete "HKLM\zSOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks\{0600DD45-FAF2-4131-A006-0B17509B9F78}" /f 2>&1 | Write-Log
    Set-OwnAndRemove -Path "$installMountDir\Windows\System32\Tasks\Microsoft\Windows\Application Experience\Microsoft Compatibility Appraiser" | Out-Null
}
reg delete "HKLM\zSOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree\Microsoft\Windows\Application Experience\PcaPatchDbTask" /f 2>&1 | Write-Log
reg delete "HKLM\zSOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree\Microsoft\Windows\Application Experience\MareBackup" /f 2>&1 | Write-Log
reg delete "HKLM\zSOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree\Microsoft\Windows\Application Experience\Microsoft Compatibility Appraiser" /f 2>&1 | Write-Log
reg delete "HKLM\zSOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree\Microsoft\Windows\Autochk\Proxy" /f 2>&1 | Write-Log
reg delete "HKLM\zSOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree\Microsoft\Windows\Customer Experience Improvement Program\Consolidator" /f 2>&1 | Write-Log
reg delete "HKLM\zSOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree\Microsoft\Windows\Customer Experience Improvement Program\UsbCeip" /f 2>&1 | Write-Log
if ($DoRemoveDeviceCensusTask) {
    Set-OwnAndRemove -Path "$installMountDir\Windows\System32\Tasks\Microsoft\Windows\Device Information\Device" | Out-Null
    reg delete "HKLM\zSOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree\Microsoft\Windows\Device Information\Device" /f 2>&1 | Write-Log
    reg add "HKLM\zSOFTWARE\Policies\Microsoft\Windows\AppCompat" /v "DisablePCA" /t REG_DWORD /d "1" /f 2>&1 | Write-Log
}
if ($DoRemoveStartupAppTask) {
    Set-OwnAndRemove -Path "$installMountDir\Windows\System32\Tasks\Microsoft\Windows\Application Experience\StartupAppTask" | Out-Null
    reg delete "HKLM\zSOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree\Microsoft\Windows\Application Experience\StartupAppTask" /f 2>&1 | Write-Log
    reg delete "HKLM\zSOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree\Microsoft\Windows\Application Experience\ProgramDataUpdater" /f 2>&1 | Write-Log
}
if ($DoDisableWER) {
    Set-OwnAndRemove -Path "$installMountDir\Windows\System32\Tasks\Microsoft\Windows\Windows Error Reporting\QueueReporting" | Out-Null
    reg delete "HKLM\zSOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree\Microsoft\Windows\Windows Error Reporting\QueueReporting" /f 2>&1 | Write-Log
}
Write-Host " v" -ForegroundColor Green

Show-Bar 100 40 "Minimal offline tweaks applied"
Show-Done "Offline registry tweaks complete"
Write-Log -msg "Minimal offline registry tweaks complete"

Show-Section "Unloading Registry Hives" "*"
Write-Log -msg "Unloading registry"
reg unload HKLM\zCOMPONENTS 2>&1 | Write-Log
reg unload HKLM\zDEFAULT 2>&1 | Write-Log
reg unload HKLM\zNTUSER 2>&1 | Write-Log
reg unload HKLM\zSOFTWARE 2>&1 | Write-Log
reg unload HKLM\zSYSTEM 2>&1 | Write-Log
Show-Done "Success"


# Integrate Intel RST/VMD Drivers
if ($DoDriverIntegrate) {
    Show-Section "Integrating Intel RST/VMD Drivers..."
    Write-Host ("  This may take some time") -ForegroundColor DarkGray
    Write-Log -msg "Starting Intel RST/VMD driver integration"
    
    Test-InternetConnection | Out-Null
    
    $driverTempPath = Join-Path $scriptDirectory "WIDTemp\drivers"
    $driverZipPath = "$driverTempPath\drivers.zip"
    $driverExtractPath = "$driverTempPath\extracted"
    $DriverURL = "https://github.com/itsNileshHere/Windows-ISO-Debloater/archive/refs/heads/main.zip"
    
    try {
        # Create temp directories
        New-Item -ItemType Directory -Path $driverTempPath -Force 2>&1 | Write-Log
        New-Item -ItemType Directory -Path $driverExtractPath -Force 2>&1 | Write-Log
        
        # Download drivers
        Write-Host "  - Downloading drivers..."  -ForegroundColor DarkGray
        $ProgressPreference = 'SilentlyContinue'
        try {
            Invoke-WebRequest -Uri $DriverURL -OutFile $driverZipPath -UseBasicParsing -ErrorAction Stop
        }
        catch {
            Write-Host "Failed to download drivers" -ForegroundColor Red
            Write-Log -msg "Driver download failed: $_"
            return
        }
        finally {
            $ProgressPreference = 'Continue'
        }
        
        # Verify download
        if (-not (Test-Path $driverZipPath)) {
            Write-Host "Driver download failed - file not found" -ForegroundColor Red
            Write-Log -msg "Driver zip file not found at: $driverZipPath"
            return
        }
        Write-Log -msg "Drivers downloaded to $driverZipPath"
        
        # Extract drivers
        Write-Host "  - Extracting drivers..." -ForegroundColor DarkGray
        try {
            Expand-Archive -Path $driverZipPath -DestinationPath $driverExtractPath -Force -ErrorAction Stop
        }
        catch {
            Write-Host "Failed to extract drivers" -ForegroundColor Red
            Write-Log -msg "Driver extraction failed: $_"
            return
        }
        Write-Log -msg "Drivers extracted to $driverExtractPath"
        
        # Get and verify driver path
        $driverSourcePath = Join-Path $driverExtractPath "Windows-ISO-Debloater-main\Drivers"
        if (-not (Test-Path $driverSourcePath)) {
            Write-Host "Driver folder not found in extracted files" -ForegroundColor Red
            Write-Log -msg "Driver folder not found at: $driverSourcePath"
            return
        }
        Write-Log -msg "Driver source path verified: $driverSourcePath"
        
        # Add drivers to install.wim
        Write-Host "  - Adding drivers to install.wim..." -ForegroundColor DarkGray
        Invoke-DismFailsafe {Add-WindowsDriver -Path $installMountDir -Driver $driverSourcePath -Recurse -ForceUnsigned} {dism /image:$installMountDir /Add-Driver /driver:$driverSourcePath /recurse /ForceUnsigned}
        Write-Log -msg "Drivers added to install.wim"
        
        # Add drivers to boot.wim
        Write-Host "  - Adding drivers to boot.wim..." -ForegroundColor DarkGray
        $bootWimPath = Join-Path $destinationPath "sources\boot.wim"
        $bootMountDir = Join-Path $scriptDirectory "WIDTemp\mountdir\bootWIM"
        New-Item -ItemType Directory -Path $bootMountDir -Force 2>&1 | Write-Log
        
        # Mount boot.wim, Add drivers, and unmount
        Invoke-DismFailsafe {Mount-WindowsImage -ImagePath $bootWimPath -Index 2 -Path $bootMountDir} {dism /mount-image /imagefile:$bootWimPath /index:2 /mountdir:$bootMountDir}
        Invoke-DismFailsafe {Add-WindowsDriver -Path $bootMountDir -Driver $driverSourcePath -Recurse -ForceUnsigned} {dism /image:$bootMountDir /Add-Driver /driver:$driverSourcePath /recurse /ForceUnsigned}
        Invoke-DismFailsafe {Dismount-WindowsImage -Path $bootMountDir -Save} {dism /unmount-image /mountdir:$bootMountDir /commit}
        
        Write-Log -msg "Drivers added to boot.wim"
        
        Show-Done "Driver integration completed"
        Write-Log -msg "Driver integration completed"
    }
    catch {
        Write-Host "Driver integration failed - skipping" -ForegroundColor Red
        Write-Log -msg "Driver integration failed: $_"
    }
    finally {
        Remove-Item -Path $driverTempPath -Recurse -Force -ErrorAction SilentlyContinue 2>&1 | Write-Log
    }
}
else {
    Write-Log -msg "Driver integration skipped"
}

# Unmounting and cleaning up the image
Show-Section "Cleaning up image..."
Write-Log -msg "Cleaning up image"
Invoke-DismFailsafe {Repair-WindowsImage -Path $installMountDir -StartComponentCleanup -ResetBase} {dism /image:$installMountDir /Cleanup-Image /StartComponentCleanup /ResetBase}

Show-Section "Unmounting and Exporting image..."
Write-Log -msg "Unmounting image"
try {
    Invoke-DismFailsafe {Dismount-WindowsImage -Path $installMountDir -Save} {dism /unmount-image /mountdir:$installMountDir /commit}
    Write-Log -msg "Image unmounted successfully"
}
catch {
    Write-Host "`n`nFailed to Unmount the Image. Check Logs for more info." -ForegroundColor Red
    Write-Host "Close all the Folders opened in the mountdir to complete the Script."
    Write-Host "Run the following code in Powershell(as admin) to unmount the broken image: "
    Write-Host "Dismount-WindowsImage -Path $installMountDir -Discard" -ForegroundColor Yellow
    Write-Log -msg "Failed to unmount image: $_"
    Pause
    Exit
}

Write-Log -msg "Exporting image"
$tempWimPath = "$destinationPath\sources\install_temp.wim"
$exportSuccess = $false

if ($DoESDConvert) {
    Show-Section "Compressing image to esd..."
    Write-Log -msg "Compressing image to esd"
    try {        
        $process = Start-Process -FilePath "dism.exe" -ArgumentList "/Export-Image /SourceImageFile:`"$destinationPath\sources\install.wim`" /SourceIndex:$sourceIndex /DestinationImageFile:`"$tempWimPath`" /Compress:Recovery /CheckIntegrity" -Wait -NoNewWindow -PassThru
        if ($process.ExitCode -eq 0 -and (Test-Path $tempWimPath)) {
            $exportSuccess = $true
            Show-Bar 100 40 "Compression done"
Show-Done "Image compressed"
            Write-Log -msg "Compression completed"
        } else {
            Write-Host "Compression failed with exit code: $($process.ExitCode)" -ForegroundColor Red
            Write-Log -msg "Compression failed with exit code: $($process.ExitCode)"
        }
    } catch {
        Write-Host "Compression failed with error: $_" -ForegroundColor Red
        Write-Log -msg "Compression failed with error: $_"
    }
}
else {
    Show-Section "Exporting image to wim..."
    Write-Log -msg "Exporting image to wim"
    try {
        Invoke-DismFailsafe {Export-WindowsImage -SourceImagePath "$destinationPath\sources\install.wim" -SourceIndex $sourceIndex -DestinationImagePath $tempWimPath -CompressionType Maximum -CheckIntegrity} {dism /Export-Image /SourceImageFile:$destinationPath\sources\install.wim /SourceIndex:$sourceIndex /DestinationImageFile:$tempWimPath /compress:max}
        if (Test-Path $tempWimPath) {
            $exportSuccess = $true
            Show-Bar 100 40 "Export done"
Show-Done "WIM exported successfully"
            Write-Log -msg "Export completed successfully"
        } else {
            Write-Host "Export failed - temp WIM not found" -ForegroundColor Red
            Write-Log -msg "Export failed - temp WIM not found"
        }
    } catch {
        Write-Host "Export failed with error: $_" -ForegroundColor Red
        Write-Log -msg "Export failed with error: $_"
    }
}

if ($exportSuccess) {
    Remove-Item -Path "$destinationPath\sources\install.wim" -Force
    Move-Item -Path $tempWimPath -Destination "$destinationPath\sources\install.wim" -Force
   
    if (-not (Test-Path "$destinationPath\sources\install.wim")) {
        Write-Host "Error: Unable to create the WIM file. Check logs for details." -ForegroundColor Red
        Write-Log -msg "Final install.wim missing"
        Pause
        Exit
    } else {
        Write-Log -msg "WIM file successfully replaced"
    }
} else {
    Write-Host "Error: Unable to export modified WIM file. Check logs for details." -ForegroundColor Red
    Write-Log -msg "WIM export failed, original WIM file preserved"
    Pause
    Exit
}

# Verify the WIM file is accessible and valid
try {
    $wimPath = Get-WindowsImage -ImagePath "$destinationPath\sources\install.wim" -ErrorAction Stop
    if ($wimPath) {
        Show-Done "WIM file validation successful: $($wimPath.Count) images found"
        Write-Log -msg "WIM validation passed: $($wimPath.Count) images found"
        
        # Force a filesystem sync to ensure all changes are written to disk
        [System.IO.File]::OpenWrite("$destinationPath\sources\install.wim").Close()
        # Add a small delay to ensure file operations are complete
        Start-Sleep -Seconds 3
    } else {
        Write-Warning "WIM file validation returned no images"
        Write-Log -msg "WIM validation warning: No images returned"
    }
} catch {
    Write-Host "Error: WIM file validation failed - $($_)" -ForegroundColor Red
    Write-Log -msg "WIM validation failed: $_"
}

Write-Log -msg "Checking required files"
if ($outputISO) {
    $ISOFileName = ($ISOFileName -replace '[<>:"/\\|?*\x00-\x1F\s]', '').Trim()
    $ISOFileName = [System.IO.Path]::GetFileNameWithoutExtension($outputISO)
} else {
    do {
        $ISOFileName = Read-Host -Prompt "`nEnter the name for the ISO file (without extension)"

        # Remove invalid characters
        $ISOFileName = ($ISOFileName -replace '[<>:"/\\|?*\x00-\x1F\s]', '').Trim()
        if ([string]::IsNullOrWhiteSpace($ISOFileName)) {
            Write-Warning "Filename is empty or invalid. Please enter a valid name."
        }
    } while ([string]::IsNullOrWhiteSpace($ISOFileName))
}
$ISOFile = Join-Path -Path $scriptDirectory -ChildPath "$ISOFileName.iso"
Write-Log -msg "ISO file name set to: $ISOFileName.iso"

if ($DoUseOscdimg) {
    if (-not (Test-Path -Path $Oscdimg)) {
        Write-Log -msg "Oscdimg.exe not found at '$Oscdimg'"
        Write-Host "`nOscdimg.exe not found at '$Oscdimg'." -ForegroundColor Red
        Write-Host "`nTrying to Download oscdimg.exe..." -ForegroundColor Cyan
        
        Test-InternetConnection | Out-Null

        # Downloading Oscdimg.exe
        # Courtesy: https://github.com/p0w3rsh3ll/ADK
        $ADKfolder = "$scriptDirectory\ADKDownload"
        $CabFileName = "5d984200acbde182fd99cbfbe9bad133.cab"
        $ExtractedFileName = "fil720cc132fbb53f3bed2e525eb77bdbc1"

        New-Item -ItemType Directory -Path $OscdimgPath -Force 2>&1 | Write-Log
        New-Item -ItemType Directory -Path $ADKfolder -Force 2>&1 | Write-Log
        
        # Resolve the URL
        $RedirectResponse = Invoke-WebRequest -Uri "https://go.microsoft.com/fwlink/?linkid=2290227" -MaximumRedirection 0 -UseBasicParsing -ErrorAction SilentlyContinue
        if ($RedirectResponse.StatusCode -eq 302) {
            $BaseURL = $RedirectResponse.Headers.Location.TrimEnd('/') + "/"
            $CabURL = "$BaseURL`Installers/$CabFileName"
            $CabFilePath = "$ADKfolder\$CabFileName"
        
            Write-Log -msg "Downloading CAB file from: $CabURL"
            Invoke-WebRequest -Uri $CabURL -OutFile $CabFilePath -UseBasicParsing
        
            # Extract the CAB file
            Write-Log -msg "Extracting CAB file..."
            expand.exe -F:* $CabFilePath $ADKfolder 2>&1 | Write-Log
        
            # Move the required file
            $ExtractedFilePath = "$ADKfolder\$ExtractedFileName"
            $FinalFilePath = "$OscdimgPath\oscdimg.exe"
        
            if (Test-Path $ExtractedFilePath) {
                Move-Item -Path $ExtractedFilePath -Destination $FinalFilePath -Force 2>&1 | Write-Log
                Write-Host "Oscdimg.exe downloaded successfully" -ForegroundColor Green
                Write-Log -msg "Oscdimg.exe successfully placed in: $OscdimgPath"
            }
            else {
                Write-Log -msg "Error: Extracted file not found!"
            }
        }
        else {
            Write-Host "Error: Failed to download Oscdimg.exe" -ForegroundColor Red
            Write-Log -msg "Failed to resolve ADK download link. HTTP Status: $($RedirectResponse.StatusCode)"
            Remove-TempFiles
            Pause
            Exit
        }
    }

    # Generate ISO
    Show-Section "Generating ISO..."
    Write-Log -msg "Generating ISO using OSCDIMG"
    try {
        $etfsbootPath = "$destinationPath\boot\etfsboot.com"
        $efisysPath = "$destinationPath\efi\Microsoft\boot\efisys.bin"
        $bootData = "2#p0,e,b`"$etfsbootPath`"#pEF,e,b`"$efisysPath`""
        Write-Log -msg "Boot data set: $bootData"
        
        $oscdimgArgs = @(
            "-bootdata:$bootData",
            "-m",               # Ignore maximum size limit
            "-o",               # Optimize for space
            "-h",               # Show hidden files
            "-u2",              # UDF 2.0
            "-udfver102",       # UDF version 1.02
            "-l$ISOFileName",   # Set volume label
            "`"$destinationPath`"",
            "`"$ISOFile`""
        )
        
        Write-Log -msg "OSCDIMG command: $Oscdimg $($oscdimgArgs -join ' ')"
        $oscdimgProcess = Start-Process -FilePath "$Oscdimg" -ArgumentList $oscdimgArgs -PassThru -Wait -NoNewWindow
        
        if ($oscdimgProcess.ExitCode -eq 0) {
            Show-Done "ISO creation successful"
            Write-Log -msg "ISO successfully created with exit code 0"
        } else {
            Write-Warning "ISO creation finished with errors"
            Write-Log -msg "OSCDIMG exited with code: $($oscdimgProcess.ExitCode)"
        }
    }
    catch {
        Write-Log -msg "Failed to generate ISO with exit code: $_"
    }
}
else {
    Show-Section "Preparing ISO creation..."
    Write-Log -msg "Preparing ISO creation"

    # ISOWriter class
    # More Here: https://learn.microsoft.com/en-us/windows/win32/api/_imapi/
    if (!('ISOWriter' -as [Type])) {
        Add-Type -TypeDefinition @'
        using System;
        using System.Runtime.InteropServices;
        using System.Runtime.InteropServices.ComTypes;

        public class ISOWriter {
            [DllImport("shlwapi.dll", CharSet = CharSet.Unicode, ExactSpelling = true, PreserveSig = false)]
            private static extern void SHCreateStreamOnFileEx(string fileName, uint mode, uint attributes, bool create, IStream streamNull, out IStream stream);
            public static bool Create(string filePath, ref object imageStream, int blockSize, int totalBlocks) {IStream resultStream = (IStream)imageStream, imageFile; SHCreateStreamOnFileEx(filePath, 0x1001, 0x80, true, null, out imageFile); const int bufferSize = 1024; int remainingBlocks = totalBlocks;
                while (remainingBlocks > 0) { int blocksToWrite = Math.Min(remainingBlocks, bufferSize); resultStream.CopyTo(imageFile, blocksToWrite * blockSize, IntPtr.Zero, IntPtr.Zero); remainingBlocks -= blocksToWrite;}
                imageFile.Commit(0);
                return true;}
        }
'@
    }

    try {
        $comObjects = @()

        # Initialize boot configuration
        $bootStream = New-Object -ComObject ADODB.Stream -Property @{ Type = 1 }
        $comObjects += $bootStream
        $bootStream.Open()
        $bootStream.LoadFromFile("$destinationPath\efi\Microsoft\boot\efisys.bin")
        # $bootStream.LoadFromFile("$destinationPath\efi\Microsoft\boot\efisys_noprompt.bin")

        # Configure boot and filesystem
        $bootOptions = New-Object -ComObject IMAPI2FS.BootOptions -Property @{
            PlatformId = 0xEF
            Manufacturer = "Microsoft"
            Emulation = 0
        }
        $comObjects += $bootOptions
        $bootOptions.AssignBootImage($bootStream)

        $FSImage = New-Object -ComObject IMAPI2FS.MsftFileSystemImage -Property @{
            FileSystemsToCreate = 4
            UDFRevision = 0x102
            FreeMediaBlocks = 0
            VolumeName = $ISOFileName
        }
        $comObjects += $FSImage
        
        Write-Log -msg "Creating ISO structure"
        $FSImage.Root.AddTree($destinationPath, $false)
        $FSImage.BootImageOptions = $bootOptions
        
        Write-Host "[INFO] Generating ISO..." -ForegroundColor Cyan
        Write-Log -msg "Generating ISO using ISOWriter"
        $resultImage = $FSImage.CreateResultImage()
        $comObjects += $resultImage

        [ISOWriter]::Create($ISOFile, [ref]$resultImage.ImageStream, $resultImage.BlockSize, $resultImage.TotalBlocks) | Out-Null
        
        if ((Get-Item $ISOFile).Length -eq ($resultImage.BlockSize * $resultImage.TotalBlocks)) {
            Write-Log -msg "ISO successfully created at: $ISOFile"
        }
    }
    catch {
        Write-Log -msg "ISO creation failed: $_" -Type Error
    }
    finally {
        foreach ($obj in $comObjects) {
            if ($obj) { 
                while ([Runtime.InteropServices.Marshal]::ReleaseComObject($obj) -gt 0) { }
            }
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        Show-Done "ISO creation successful"
    }
}

# ISO verification using 7-Zip (no mount required)
if (Test-Path -Path $ISOFile) {
    try {
        Write-Log -msg "Verifying output ISO with 7-Zip..."
$missingFiles = $reqFiles | Where-Object {
    $pattern = $_ -replace '/', '\\'   # use Windows path separator
    $verifyOutput -notmatch [regex]::Escape($pattern)
}

        if ($missingFiles) {
            Write-Host "`nError: Created ISO is missing critical files" -ForegroundColor Red
            Write-Log -msg "ISO verification failed - missing files: $($missingFiles -join ', ')"
        }
        else {
            Write-Host ""
Write-Host "  $("="*58)" -ForegroundColor Magenta
Write-Host "  v  ALL DONE" -ForegroundColor Green
Write-Host "  $("="*58)" -ForegroundColor Magenta
Show-Done "ISO saved to: $scriptDirectory"
Show-Bar 100 40 "Complete"
Write-Host ""
            Write-Log -msg "ISO verification successful"
        }
    }
    catch {
        Write-Warning "`nUnable to verify ISO integrity"
        Write-Log -msg "Failed to verify ISO: $_"
    }
} else {
    Write-Host "`nError: ISO file wasn't created" -ForegroundColor Red
    Write-Log -msg "ISO file wasn't created"
}

# Remove temporary files
Write-Log -msg "Removing temporary files"
try {
    Remove-TempFiles
}
catch {
    Write-Log -msg "Failed to remove temporary files: $_"
}
finally {
    Write-Log -msg "Script completed"
}

Write-Host
Pause
