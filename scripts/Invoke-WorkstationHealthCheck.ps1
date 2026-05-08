# ArcForge Studio IT Toolkit
# Workstation Health Check v0.6

$ReportDate = Get-Date
$ComputerName = $env:COMPUTERNAME
$CurrentUser = $env:USERNAME

$ProjectRoot = Split-Path -Parent $PSScriptRoot
$ReportFolder = Join-Path $ProjectRoot "reports"
$Timestamp = Get-Date -Format "yyyy-MM-dd-HHmmss"
$ReportFile = Join-Path $ReportFolder "$ComputerName-healthcheck-$Timestamp.txt"
$ReportLines = New-Object System.Collections.Generic.List[string]

$CheckCounts = @{
    OK = 0
    WARN = 0
    FAIL = 0
}

if (-not (Test-Path $ReportFolder)) {
    New-Item -Path $ReportFolder -ItemType Directory | Out-Null
}

function Add-ReportLine {
    param (
        [string]$Line = ""
    )

    $script:ReportLines.Add($Line) | Out-Null
}

function Write-Result {
    param (
        [string]$Status,
        [string]$Label,
        [string]$Value,
        [bool]$CountResult = $true
    )

    $StatusUpper = $Status.ToUpper()
    $StatusFieldWidth = 6
    $CurrentStatusWidth = $StatusUpper.Length + 2
    $StatusPadding = " " * ($StatusFieldWidth - $CurrentStatusWidth)
    $LabelPadded = "{0,-18}" -f $Label

    if ($CountResult -and $script:CheckCounts.ContainsKey($StatusUpper)) {
        $script:CheckCounts[$StatusUpper]++
    }

    Write-Host "[" -NoNewline -ForegroundColor Gray

    switch ($StatusUpper) {
        "OK"   { Write-Host "OK" -NoNewline -ForegroundColor Green }
        "WARN" { Write-Host "WARN" -NoNewline -ForegroundColor Yellow }
        "FAIL" { Write-Host "FAIL" -NoNewline -ForegroundColor Red }
        default { Write-Host $StatusUpper -NoNewline -ForegroundColor Gray }
    }

    Write-Host "]$StatusPadding  " -NoNewline -ForegroundColor Gray
    Write-Host "$LabelPadded $Value" -ForegroundColor Gray

    Add-ReportLine -Line ("[{0}]{1}  {2} {3}" -f $StatusUpper, $StatusPadding, $LabelPadded, $Value)
}

function Write-Section {
    param (
        [string]$Title
    )

    Write-Host ""
    Write-Host "[$Title]" -ForegroundColor Gray

    Add-ReportLine
    Add-ReportLine -Line "[$Title]"
}

function Write-Summary {
    Write-Section -Title "SUMMARY"

    $PassedChecks = $script:CheckCounts.OK
    $Warnings = $script:CheckCounts.WARN
    $Failures = $script:CheckCounts.FAIL
    $TotalChecks = $PassedChecks + $Warnings + $Failures

    Write-Result -Status "OK" -Label "Total Checks:" -Value $TotalChecks -CountResult:$false
    Write-Result -Status "OK" -Label "Passed Checks:" -Value $PassedChecks -CountResult:$false

    if ($Warnings -gt 0) {
        Write-Result -Status "WARN" -Label "Warnings:" -Value $Warnings -CountResult:$false
    }
    else {
        Write-Result -Status "OK" -Label "Warnings:" -Value $Warnings -CountResult:$false
    }

    if ($Failures -gt 0) {
        Write-Result -Status "FAIL" -Label "Failures:" -Value $Failures -CountResult:$false
    }
    else {
        Write-Result -Status "OK" -Label "Failures:" -Value $Failures -CountResult:$false
    }

    if ($Failures -gt 0) {
        Write-Result -Status "FAIL" -Label "Overall Status:" -Value "Action required" -CountResult:$false
    }
    elseif ($Warnings -gt 0) {
        Write-Result -Status "WARN" -Label "Overall Status:" -Value "Attention recommended" -CountResult:$false
    }
    else {
        Write-Result -Status "OK" -Label "Overall Status:" -Value "Healthy" -CountResult:$false
    }
}

function Test-SoftwareInstalled {
    param (
        [string[]]$Commands = @(),
        [string[]]$DisplayNamePatterns = @(),
        [string[]]$CommonPaths = @()
    )

    foreach ($Command in $Commands) {
        if (Get-Command $Command -ErrorAction SilentlyContinue) {
            return $true
        }
    }

    foreach ($Path in $CommonPaths) {
        if ($Path -and (Test-Path $Path)) {
            return $true
        }
    }

    $RegistryPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    foreach ($RegistryPath in $RegistryPaths) {
        $InstalledApps = Get-ItemProperty $RegistryPath -ErrorAction SilentlyContinue

        foreach ($Pattern in $DisplayNamePatterns) {
            if ($InstalledApps.DisplayName -like $Pattern) {
                return $true
            }
        }
    }

    return $false
}

Write-Host "========================================" -ForegroundColor Gray
Write-Host " ArcForge Studio Workstation Health Check" -ForegroundColor Gray
Write-Host "========================================" -ForegroundColor Gray
Write-Host ""

Add-ReportLine -Line "========================================"
Add-ReportLine -Line " ArcForge Studio Workstation Health Check"
Add-ReportLine -Line "========================================"
Add-ReportLine

Write-Result -Status "OK" -Label "Computer Name:" -Value $ComputerName
Write-Result -Status "OK" -Label "Current User:" -Value $CurrentUser
Write-Result -Status "OK" -Label "Report Date:" -Value $ReportDate

# System Checks
Write-Section -Title "SYSTEM"

try {
    $OS = Get-CimInstance Win32_OperatingSystem

    Write-Result -Status "OK" -Label "OS Name:" -Value $OS.Caption
    Write-Result -Status "OK" -Label "OS Version:" -Value $OS.Version
    Write-Result -Status "OK" -Label "Architecture:" -Value $OS.OSArchitecture
}
catch {
    Write-Result -Status "FAIL" -Label "System Info:" -Value $_.Exception.Message
}

# Uptime Check
Write-Section -Title "UPTIME"

try {
    $LastBoot = $OS.LastBootUpTime
    $Uptime = (Get-Date) - $LastBoot
    $UptimeDays = [math]::Round($Uptime.TotalDays, 2)

    Write-Result -Status "OK" -Label "Last Boot:" -Value $LastBoot

    if ($UptimeDays -ge 14) {
        Write-Result -Status "WARN" -Label "Uptime Days:" -Value "$UptimeDays days - reboot recommended"
    }
    else {
        Write-Result -Status "OK" -Label "Uptime Days:" -Value "$UptimeDays days"
    }
}
catch {
    Write-Result -Status "FAIL" -Label "Uptime:" -Value $_.Exception.Message
}

# Storage Check
Write-Section -Title "STORAGE"

try {
    $Disk = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
    $FreeGB = [math]::Round($Disk.FreeSpace / 1GB, 2)
    $TotalGB = [math]::Round($Disk.Size / 1GB, 2)
    $FreePercent = [math]::Round(($Disk.FreeSpace / $Disk.Size) * 100, 2)

    Write-Result -Status "OK" -Label "Drive:" -Value "C:"
    Write-Result -Status "OK" -Label "Total Size:" -Value "$TotalGB GB"

    if ($FreePercent -lt 10) {
        Write-Result -Status "FAIL" -Label "Free Space:" -Value "$FreeGB GB free ($FreePercent%) - critically low"
    }
    elseif ($FreePercent -lt 20) {
        Write-Result -Status "WARN" -Label "Free Space:" -Value "$FreeGB GB free ($FreePercent%) - low disk space"
    }
    else {
        Write-Result -Status "OK" -Label "Free Space:" -Value "$FreeGB GB free ($FreePercent%)"
    }
}
catch {
    Write-Result -Status "FAIL" -Label "Storage:" -Value $_.Exception.Message
}

# Network Checks
Write-Section -Title "NETWORK"

try {
    $NetworkConfig = Get-CimInstance Win32_NetworkAdapterConfiguration |
        Where-Object {
            $_.IPEnabled -eq $true -and
            $_.DefaultIPGateway -ne $null
        } |
        Select-Object -First 1

    if ($NetworkConfig) {
        $Gateway = $NetworkConfig.DefaultIPGateway[0]
        $IPAddress = $NetworkConfig.IPAddress | Where-Object { $_ -match '^\d{1,3}(\.\d{1,3}){3}$' } | Select-Object -First 1
        $DNSServers = $NetworkConfig.DNSServerSearchOrder -join ", "

        Write-Result -Status "OK" -Label "IPv4 Address:" -Value $IPAddress
        Write-Result -Status "OK" -Label "Gateway:" -Value $Gateway
        Write-Result -Status "OK" -Label "DNS Servers:" -Value $DNSServers

        if (Test-Connection -ComputerName $Gateway -Count 2 -Quiet) {
            Write-Result -Status "OK" -Label "Gateway Ping:" -Value "Reachable"
        }
        else {
            Write-Result -Status "FAIL" -Label "Gateway Ping:" -Value "Unreachable"
        }
    }
    else {
        Write-Result -Status "FAIL" -Label "Network Config:" -Value "No active adapter with default gateway found"
    }
}
catch {
    Write-Result -Status "FAIL" -Label "Network Config:" -Value $_.Exception.Message
}

if (Test-Connection -ComputerName "1.1.1.1" -Count 2 -Quiet) {
    Write-Result -Status "OK" -Label "Internet Ping:" -Value "1.1.1.1 reachable"
}
else {
    Write-Result -Status "FAIL" -Label "Internet Ping:" -Value "1.1.1.1 unreachable"
}

try {
    Resolve-DnsName "github.com" -ErrorAction Stop | Out-Null
    Write-Result -Status "OK" -Label "DNS Resolution:" -Value "github.com resolved"
}
catch {
    Write-Result -Status "FAIL" -Label "DNS Resolution:" -Value "Failed to resolve github.com"
}

# Software Checks
Write-Section -Title "SOFTWARE - CORE IT"

$ProgramFilesX86 = ${env:ProgramFiles(x86)}

$CoreTools = @(
    @{
        Name = "Git"
        Commands = @("git")
        DisplayNamePatterns = @("Git*")
        CommonPaths = @(
            "$env:ProgramFiles\Git\cmd\git.exe"
        )
    },
    @{
        Name = "VS Code"
        Commands = @("code")
        DisplayNamePatterns = @("Microsoft Visual Studio Code*", "Visual Studio Code*")
        CommonPaths = @(
            "$env:LOCALAPPDATA\Programs\Microsoft VS Code\Code.exe",
            "$env:ProgramFiles\Microsoft VS Code\Code.exe"
        )
    },
    @{
        Name = "PowerShell 7"
        Commands = @("pwsh")
        DisplayNamePatterns = @("PowerShell 7*", "Microsoft PowerShell*")
        CommonPaths = @(
            "$env:ProgramFiles\PowerShell\7\pwsh.exe"
        )
    },
    @{
        Name = "Notepad++"
        Commands = @("notepad++")
        DisplayNamePatterns = @("Notepad++*")
        CommonPaths = @(
            "$env:ProgramFiles\Notepad++\notepad++.exe",
            "$ProgramFilesX86\Notepad++\notepad++.exe"
        )
    }
)

foreach ($Tool in $CoreTools) {
    $Installed = Test-SoftwareInstalled `
        -Commands $Tool.Commands `
        -DisplayNamePatterns $Tool.DisplayNamePatterns `
        -CommonPaths $Tool.CommonPaths

    if ($Installed) {
        Write-Result -Status "OK" -Label "$($Tool.Name):" -Value "Installed"
    }
    else {
        Write-Result -Status "WARN" -Label "$($Tool.Name):" -Value "Not found"
    }
}

Write-Section -Title "SOFTWARE - GAME DEV"

$GameDevTools = @(
    @{
        Name = "Python"
        Commands = @("python", "py")
        DisplayNamePatterns = @("Python*")
        CommonPaths = @(
            "$env:LOCALAPPDATA\Programs\Python\Python*\python.exe",
            "$env:ProgramFiles\Python*\python.exe"
        )
    },
    @{
        Name = "Visual Studio"
        Commands = @("devenv")
        DisplayNamePatterns = @("Microsoft Visual Studio*")
        CommonPaths = @(
            "$env:ProgramFiles\Microsoft Visual Studio\2022\Community\Common7\IDE\devenv.exe",
            "$env:ProgramFiles\Microsoft Visual Studio\2022\Professional\Common7\IDE\devenv.exe",
            "$env:ProgramFiles\Microsoft Visual Studio\2022\Enterprise\Common7\IDE\devenv.exe"
        )
    },
    @{
        Name = "Perforce"
        Commands = @("p4", "p4v")
        DisplayNamePatterns = @("Perforce*", "Helix*")
        CommonPaths = @(
            "$env:ProgramFiles\Perforce\p4.exe",
            "$env:ProgramFiles\Perforce\p4v.exe",
            "$ProgramFilesX86\Perforce\p4.exe",
            "$ProgramFilesX86\Perforce\p4v.exe"
        )
    },
    @{
        Name = "Unreal Engine"
        Commands = @()
        DisplayNamePatterns = @("Unreal Engine*", "Epic Games Launcher*")
        CommonPaths = @(
            "$env:ProgramFiles\Epic Games\UE_*",
            "$env:ProgramFiles\Epic Games\Launcher\Portal\Binaries\Win64\EpicGamesLauncher.exe"
        )
    }
)

foreach ($Tool in $GameDevTools) {
    $Installed = Test-SoftwareInstalled `
        -Commands $Tool.Commands `
        -DisplayNamePatterns $Tool.DisplayNamePatterns `
        -CommonPaths $Tool.CommonPaths

    if ($Installed) {
        Write-Result -Status "OK" -Label "$($Tool.Name):" -Value "Installed"
    }
    else {
        Write-Result -Status "WARN" -Label "$($Tool.Name):" -Value "Not found"
    }
}

Write-Section -Title "SOFTWARE - ART/PIPELINE"

$ArtPipelineTools = @(
    @{
        Name = "Blender"
        Commands = @("blender")
        DisplayNamePatterns = @("Blender*")
        CommonPaths = @(
            "$env:ProgramFiles\Blender Foundation\Blender*\blender.exe"
        )
    },
    @{
        Name = "Maya"
        Commands = @("maya")
        DisplayNamePatterns = @("Autodesk Maya*")
        CommonPaths = @(
            "$env:ProgramFiles\Autodesk\Maya*\bin\maya.exe"
        )
    },
    @{
        Name = "Houdini"
        Commands = @("houdini")
        DisplayNamePatterns = @("Houdini*", "SideFX Houdini*")
        CommonPaths = @(
            "$env:ProgramFiles\Side Effects Software\Houdini*\bin\houdini.exe"
        )
    },
    @{
        Name = "ZBrush"
        Commands = @()
        DisplayNamePatterns = @("ZBrush*", "Maxon ZBrush*")
        CommonPaths = @(
            "$env:ProgramFiles\Maxon ZBrush*\ZBrush.exe",
            "$env:ProgramFiles\Pixologic\ZBrush*\ZBrush.exe"
        )
    }
)

foreach ($Tool in $ArtPipelineTools) {
    $Installed = Test-SoftwareInstalled `
        -Commands $Tool.Commands `
        -DisplayNamePatterns $Tool.DisplayNamePatterns `
        -CommonPaths $Tool.CommonPaths

    if ($Installed) {
        Write-Result -Status "OK" -Label "$($Tool.Name):" -Value "Installed"
    }
    else {
        Write-Result -Status "WARN" -Label "$($Tool.Name):" -Value "Not found"
    }
}

Write-Section -Title "SOFTWARE - LAUNCHERS/QA"

$LauncherQATools = @(
    @{
        Name = "Riot Client"
        Commands = @()
        DisplayNamePatterns = @("Riot Client*", "Riot Games*")
        CommonPaths = @(
            "C:\Riot Games\Riot Client\RiotClientServices.exe"
        )
    },
    @{
        Name = "Steam"
        Commands = @("steam")
        DisplayNamePatterns = @("Steam*")
        CommonPaths = @(
            "$ProgramFilesX86\Steam\steam.exe",
            "$env:ProgramFiles\Steam\steam.exe"
        )
    },
    @{
        Name = "Epic Launcher"
        Commands = @()
        DisplayNamePatterns = @("Epic Games Launcher*")
        CommonPaths = @(
            "$env:ProgramFiles\Epic Games\Launcher\Portal\Binaries\Win64\EpicGamesLauncher.exe",
            "$ProgramFilesX86\Epic Games\Launcher\Portal\Binaries\Win64\EpicGamesLauncher.exe"
        )
    }
)

foreach ($Tool in $LauncherQATools) {
    $Installed = Test-SoftwareInstalled `
        -Commands $Tool.Commands `
        -DisplayNamePatterns $Tool.DisplayNamePatterns `
        -CommonPaths $Tool.CommonPaths

    if ($Installed) {
        Write-Result -Status "OK" -Label "$($Tool.Name):" -Value "Installed"
    }
    else {
        Write-Result -Status "WARN" -Label "$($Tool.Name):" -Value "Not found"
    }
}

# Security Checks
Write-Section -Title "SECURITY"

try {
    $FirewallProfiles = Get-NetFirewallProfile -ErrorAction Stop
    $DisabledProfiles = $FirewallProfiles | Where-Object { $_.Enabled -eq $false }

    if ($DisabledProfiles.Count -eq 0) {
        Write-Result -Status "OK" -Label "Firewall:" -Value "Enabled for all profiles"
    }
    else {
        $DisabledNames = ($DisabledProfiles.Name -join ", ")
        Write-Result -Status "WARN" -Label "Firewall:" -Value "Disabled profile(s): $DisabledNames"
    }
}
catch {
    try {
        $FirewallState = netsh advfirewall show allprofiles state

        if ($FirewallState -match "State\s+OFF") {
            Write-Result -Status "WARN" -Label "Firewall:" -Value "One or more profiles may be disabled"
        }
        elseif ($FirewallState -match "State\s+ON") {
            Write-Result -Status "OK" -Label "Firewall:" -Value "Enabled - verified with netsh"
        }
        else {
            Write-Result -Status "WARN" -Label "Firewall:" -Value "Unable to determine firewall state"
        }
    }
    catch {
        Write-Result -Status "WARN" -Label "Firewall:" -Value "Unable to query firewall status"
    }
}

try {
    $DefenderStatus = Get-MpComputerStatus -ErrorAction Stop

    if ($DefenderStatus.AntivirusEnabled) {
        Write-Result -Status "OK" -Label "Defender AV:" -Value "Enabled"
    }
    else {
        Write-Result -Status "FAIL" -Label "Defender AV:" -Value "Disabled"
    }

    if ($DefenderStatus.RealTimeProtectionEnabled) {
        Write-Result -Status "OK" -Label "Real-Time Protect:" -Value "Enabled"
    }
    else {
        Write-Result -Status "FAIL" -Label "Real-Time Protect:" -Value "Disabled"
    }
}
catch {
    Write-Result -Status "WARN" -Label "Defender:" -Value "Unable to query Defender status - verify manually"
}

try {
    $LocalAdmins = Get-LocalGroupMember -Group "Administrators" -ErrorAction Stop
    $AdminCount = $LocalAdmins.Count

    if ($AdminCount -le 1) {
        Write-Result -Status "OK" -Label "Local Admins:" -Value "$AdminCount member"
    }
    else {
        Write-Result -Status "WARN" -Label "Local Admins:" -Value "$AdminCount members - review recommended"
    }
}
catch {
    Write-Result -Status "WARN" -Label "Local Admins:" -Value "Unable to query local administrators"
}

Write-Summary

Write-Host ""
Write-Host "Health check complete." -ForegroundColor Gray
Write-Host "Report saved to: $ReportFile" -ForegroundColor Gray

Add-ReportLine
Add-ReportLine -Line "Health check complete."
Add-ReportLine -Line "Report saved to: $ReportFile"

$ReportLines | Out-File -FilePath $ReportFile -Encoding UTF8