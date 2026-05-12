# ArcForge First Response
# ArcForge First Response Report v0.13

param (
    [ValidateSet("General", "Gaming", "Creator", "Developer", "Homelab", "Secure")]
    [string]$BattlestationProfile = "General"
)

$ReportDate = Get-Date
$ComputerName = $env:COMPUTERNAME
$CurrentUser = $env:USERNAME

$ProjectRoot = Split-Path -Parent $PSScriptRoot
$ReportFolder = Join-Path $ProjectRoot "reports"
$Timestamp = Get-Date -Format "yyyy-MM-dd-HHmmss"
$ProfileNameForFile = $BattlestationProfile.ToLower()
$ReportFile = Join-Path $ReportFolder "$ComputerName-$ProfileNameForFile-first-response-$Timestamp.txt"
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
        [string]$SoftwareName = "",
        [string[]]$Commands = @(),
        [string[]]$DisplayNamePatterns = @(),
        [string[]]$CommonPaths = @(),
        [string[]]$Services = @()
    )

    # Known-app fallback: VS Code
    if ($SoftwareName -eq "VS Code") {
        $VsCodePaths = @(
            "$env:ProgramFiles\Microsoft VS Code\Code.exe",
            "${env:ProgramFiles(x86)}\Microsoft VS Code\Code.exe",
            "$env:LOCALAPPDATA\Programs\Microsoft VS Code\Code.exe"
        )

        if (Get-Command "code" -ErrorAction SilentlyContinue) {
            return $true
        }

        foreach ($Path in $VsCodePaths) {
            if (Test-Path $Path) {
                return $true
            }
        }

        $RegistryMatches = Get-ItemProperty `
            "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*", `
            "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*", `
            "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" `
            -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -like "*Visual Studio Code*" }

        if ($RegistryMatches) {
            return $true
        }
    }

    # Known software detection normalization.
    # This supplements the CSV catalog when Detection Target is too human-readable.
    $KnownSoftwareDetections = @{
        "Chrome" = @{
            Commands = @("chrome")
            DisplayNamePatterns = @("Google Chrome*")
            CommonPaths = @(
                "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
                "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"
            )
        }

        "PowerShell" = @{
            Commands = @("pwsh")
            DisplayNamePatterns = @("PowerShell*", "Microsoft PowerShell*")
            CommonPaths = @(
                "$env:ProgramFiles\PowerShell\7\pwsh.exe"
            )
        }

        "Notepad++" = @{
            Commands = @("notepad++")
            DisplayNamePatterns = @("Notepad++*")
            CommonPaths = @(
                "$env:ProgramFiles\Notepad++\notepad++.exe",
                "${env:ProgramFiles(x86)}\Notepad++\notepad++.exe"
            )
        }

        "Windows Terminal" = @{
            Commands = @("wt")
            DisplayNamePatterns = @("Windows Terminal*", "Microsoft Windows Terminal*")
            CommonPaths = @(
                "$env:LOCALAPPDATA\Microsoft\WindowsApps\wt.exe"
            )
        }

        "7-Zip" = @{
            Commands = @("7z")
            DisplayNamePatterns = @("7-Zip*", "7zip*")
            CommonPaths = @(
                "$env:ProgramFiles\7-Zip\7z.exe",
                "$env:ProgramFiles\7-Zip\7zFM.exe",
                "${env:ProgramFiles(x86)}\7-Zip\7z.exe",
                "${env:ProgramFiles(x86)}\7-Zip\7zFM.exe"
            )
        }

        "OpenHashTab" = @{
            Commands = @()
            DisplayNamePatterns = @("OpenHashTab*")
            CommonPaths = @()
        }

        "FFmpeg (full)" = @{
            Commands = @("ffmpeg")
            DisplayNamePatterns = @("FFmpeg*", "Gyan.FFmpeg*", "Gyan FFmpeg*")
            CommonPaths = @(
                "$env:ProgramFiles\ffmpeg\bin\ffmpeg.exe",
                "${env:ProgramFiles(x86)}\ffmpeg\bin\ffmpeg.exe"
            )
        }

        "Python3" = @{
            Commands = @("python", "py")
            DisplayNamePatterns = @("Python*")
            CommonPaths = @(
                "$env:LOCALAPPDATA\Programs\Python\Python*\python.exe",
                "$env:ProgramFiles\Python*\python.exe",
                "${env:ProgramFiles(x86)}\Python*\python.exe"
            )
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($SoftwareName)) {
        if ($KnownSoftwareDetections.ContainsKey($SoftwareName)) {
            $KnownDetection = $KnownSoftwareDetections[$SoftwareName]

            $Commands += $KnownDetection.Commands
            $DisplayNamePatterns += $KnownDetection.DisplayNamePatterns
            $CommonPaths += $KnownDetection.CommonPaths
        }
    }

    foreach ($Command in $Commands) {
        if (-not [string]::IsNullOrWhiteSpace($Command)) {
            if (Get-Command $Command -ErrorAction SilentlyContinue) {
                return $true
            }
        }
    }

    foreach ($Service in $Services) {
        if (-not [string]::IsNullOrWhiteSpace($Service)) {
            if (Get-Service -Name $Service -ErrorAction SilentlyContinue) {
                return $true
            }

            if (Get-CimInstance Win32_Service -Filter "Name='$Service'" -ErrorAction SilentlyContinue) {
                return $true
            }
        }
    }

    foreach ($Path in $CommonPaths) {
        if (-not [string]::IsNullOrWhiteSpace($Path)) {
            if (Test-Path $Path) {
                return $true
            }
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
            if (-not [string]::IsNullOrWhiteSpace($Pattern)) {
                if ($InstalledApps.DisplayName -like $Pattern) {
                    return $true
                }
            }
        }
    }

    # Generic fallback: match installed programs by software name.
    # This helps catch normal desktop apps when catalog detection targets are too human-readable.
    if (-not [string]::IsNullOrWhiteSpace($SoftwareName)) {
        $SafeSoftwareName = $SoftwareName.Trim()

        $NameFallbackPatterns = @(
            "*$SafeSoftwareName*"
        )

        # Small normalization helpers for common catalog display names.
        switch -Wildcard ($SafeSoftwareName) {
            "Chrome" {
                $NameFallbackPatterns += "*Google Chrome*"
            }
            "PowerShell" {
                $NameFallbackPatterns += "*PowerShell*"
                $NameFallbackPatterns += "*Microsoft PowerShell*"
            }
            "Windows Terminal" {
                $NameFallbackPatterns += "*Windows Terminal*"
            }
            "7-Zip" {
                $NameFallbackPatterns += "*7-Zip*"
                $NameFallbackPatterns += "*7zip*"
            }
            "Notepad++" {
                $NameFallbackPatterns += "*Notepad++*"
            }
            "OpenHashTab" {
                $NameFallbackPatterns += "*OpenHashTab*"
            }
            "Firefox" {
                $NameFallbackPatterns += "*Mozilla Firefox*"
            }
            "Discord" {
                $NameFallbackPatterns += "*Discord*"
            }
        }

        foreach ($RegistryPath in $RegistryPaths) {
            $InstalledApps = Get-ItemProperty $RegistryPath -ErrorAction SilentlyContinue

            foreach ($Pattern in $NameFallbackPatterns) {
                if ($InstalledApps.DisplayName -like $Pattern) {
                    return $true
                }
            }
        }
    }

    return $false
}

function Test-YesValue {
    param (
        [object]$Value
    )

    return (([string]$Value).Trim().ToLower() -eq "yes")
}

function Get-CatalogValue {
    param (
        [pscustomobject]$Row,
        [string]$ColumnName
    )

    $Property = $Row.PSObject.Properties[$ColumnName]

    if ($Property) {
        return ([string]$Property.Value).Trim()
    }

    return ""
}

function Get-DisplayNamePatterns {
    param (
        [string]$SoftwareName,
        [string]$DetectionTarget
    )

    $Patterns = New-Object System.Collections.Generic.List[string]

    if (-not [string]::IsNullOrWhiteSpace($SoftwareName)) {
        $Patterns.Add("*$SoftwareName*") | Out-Null

        $BaseName = ($SoftwareName -replace "\s*\(.*?\)", "").Trim()
        if ($BaseName -and $BaseName -ne $SoftwareName) {
            $Patterns.Add("*$BaseName*") | Out-Null
        }
    }

    if ($DetectionTarget -match "matching\s+(.+)$") {
        $TargetName = $Matches[1].Trim()
        if ($TargetName) {
            $Patterns.Add("*$TargetName*") | Out-Null

            $TargetBaseName = ($TargetName -replace "\s*\(.*?\)", "").Trim()
            if ($TargetBaseName -and $TargetBaseName -ne $TargetName) {
                $Patterns.Add("*$TargetBaseName*") | Out-Null
            }
        }
    }

    return @($Patterns | Sort-Object -Unique)
}

function Split-DetectionCandidates {
    param (
        [string]$DetectionTarget
    )

    if ([string]::IsNullOrWhiteSpace($DetectionTarget)) {
        return @()
    }

    return @(
        $DetectionTarget -split "\s*/\s*|;|,|\s+or\s+" |
            ForEach-Object { $_.Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )
}

function Get-SoftwareDetectionConfig {
    param (
        [pscustomobject]$CatalogRow
    )

    $SoftwareName = Get-CatalogValue -Row $CatalogRow -ColumnName "Software Name"
    $DetectionMethod = Get-CatalogValue -Row $CatalogRow -ColumnName "Detection Method"
    $DetectionTarget = Get-CatalogValue -Row $CatalogRow -ColumnName "Detection Target"

    $Commands = New-Object System.Collections.Generic.List[string]
    $CommonPaths = New-Object System.Collections.Generic.List[string]
    $Services = New-Object System.Collections.Generic.List[string]

    $Candidates = Split-DetectionCandidates -DetectionTarget $DetectionTarget

    foreach ($Candidate in $Candidates) {
        $CleanCandidate = $Candidate.Trim()

        if ([string]::IsNullOrWhiteSpace($CleanCandidate)) {
            continue
        }

        if ($DetectionMethod -match "Command") {
            if ($CleanCandidate -match "^[A-Za-z0-9_.+\-\*]+(\.exe)?$") {
                $Commands.Add(($CleanCandidate -replace "\.exe$", "")) | Out-Null
            }
        }

        if ($CleanCandidate -match "\.exe$") {
            $ExeName = Split-Path $CleanCandidate -Leaf
            $CommandName = $ExeName -replace "\.exe$", ""

            if ($CommandName) {
                $Commands.Add($CommandName) | Out-Null
            }

            $CommonPaths.Add((Join-Path $env:ProgramFiles "*\$ExeName")) | Out-Null

            if (${env:ProgramFiles(x86)}) {
                $CommonPaths.Add((Join-Path ${env:ProgramFiles(x86)} "*\$ExeName")) | Out-Null
            }

            if ($env:LOCALAPPDATA) {
                $CommonPaths.Add((Join-Path $env:LOCALAPPDATA "Programs\*\$ExeName")) | Out-Null
            }
        }

        if ($DetectionMethod -match "Service") {
            if ($CleanCandidate -match "(?i)^(.+?)\s+service$") {
                $Services.Add($Matches[1].Trim()) | Out-Null
            }
            elseif ($CleanCandidate -match "^[A-Za-z0-9_.\-]+$") {
                $Services.Add($CleanCandidate) | Out-Null
            }
        }
    }

    $DisplayNamePatterns = Get-DisplayNamePatterns -SoftwareName $SoftwareName -DetectionTarget $DetectionTarget

    [pscustomobject]@{
        Commands = @($Commands | Sort-Object -Unique)
        DisplayNamePatterns = @($DisplayNamePatterns | Sort-Object -Unique)
        CommonPaths = @($CommonPaths | Sort-Object -Unique)
        Services = @($Services | Sort-Object -Unique)
    }
}

Write-Host "=========================" -ForegroundColor Gray
Write-Host " ArcForge First Response" -ForegroundColor Gray
Write-Host "=========================" -ForegroundColor Gray
Write-Host ""

Add-ReportLine -Line "========================="
Add-ReportLine -Line " ArcForge First Response"
Add-ReportLine -Line "========================="
Add-ReportLine

Write-Result -Status "OK" -Label "Computer Name:" -Value $ComputerName
Write-Result -Status "OK" -Label "Current User:" -Value $CurrentUser
Write-Result -Status "OK" -Label "Report Date:" -Value $ReportDate
Write-Result -Status "OK" -Label "Active Profile:" -Value $BattlestationProfile

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

# Process Readiness Checks
Write-Section -Title "PROCESSES"

try {
    $HungProcesses = Get-Process -ErrorAction Stop |
        Where-Object {
            $_.MainWindowTitle -and
            $_.Responding -eq $false
        }

    if ($HungProcesses.Count -gt 0) {
        $HungNames = ($HungProcesses.ProcessName | Sort-Object -Unique) -join ", "
        Write-Result -Status "WARN" -Label "Hung Apps:" -Value "$($HungProcesses.Count) non-responding app(s): $HungNames"
    }
    else {
        Write-Result -Status "OK" -Label "Hung Apps:" -Value "None detected"
    }
}
catch {
    Write-Result -Status "WARN" -Label "Hung Apps:" -Value "Unable to query process responsiveness"
}

try {
    $TopMemoryProcesses = Get-Process -ErrorAction Stop |
        Sort-Object WorkingSet64 -Descending |
        Select-Object -First 5

    foreach ($Process in $TopMemoryProcesses) {
        $MemoryMB = [math]::Round($Process.WorkingSet64 / 1MB, 2)
        Write-Result -Status "OK" -Label "Top Memory:" -Value "$($Process.ProcessName) - $MemoryMB MB" -CountResult:$false
    }
}
catch {
    Write-Result -Status "WARN" -Label "Memory Usage:" -Value "Unable to query top memory processes"
}

# Service Readiness Checks
Write-Section -Title "SERVICES"

# Core workstation services only.
# Antivirus/security provider service validation will be handled separately later
# so third-party AV products do not trigger false Defender warnings.
$CoreServices = @(
    @{
        Name = "EventLog"
        Label = "Event Log:"
    },
    @{
        Name = "Winmgmt"
        Label = "WMI:"
    },
    @{
        Name = "LanmanWorkstation"
        Label = "Workstation:"
    },
    @{
        Name = "Dnscache"
        Label = "DNS Client:"
    }
)

foreach ($Service in $CoreServices) {
    try {
        $ServiceInfo = Get-CimInstance Win32_Service -Filter "Name='$($Service.Name)'" -ErrorAction Stop

        if (-not $ServiceInfo) {
            Write-Result -Status "WARN" -Label $Service.Label -Value "Service not found"
        }
        elseif ($ServiceInfo.StartMode -eq "Disabled") {
            Write-Result -Status "WARN" -Label $Service.Label -Value "Disabled"
        }
        elseif ($ServiceInfo.State -eq "Running") {
            Write-Result -Status "OK" -Label $Service.Label -Value "$($ServiceInfo.StartMode) / Running"
        }
        else {
            Write-Result -Status "WARN" -Label $Service.Label -Value "$($ServiceInfo.StartMode) / $($ServiceInfo.State)"
        }
    }
    catch {
        Write-Result -Status "WARN" -Label $Service.Label -Value "Unable to query service"
    }
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
Write-Section -Title "SOFTWARE"

$CatalogFolder = Join-Path $ProjectRoot "catalog"
$CatalogFile = Join-Path $CatalogFolder "arcforge-software-catalog.csv"

if ($BattlestationProfile -eq "General") {
    Write-Result -Status "OK" -Label "Profile Tools:" -Value "No profile-specific software checks for General profile" -CountResult:$false
}
elseif (-not (Test-Path $CatalogFile)) {
    Write-Result -Status "WARN" -Label "Catalog File:" -Value "Not found at $CatalogFile"
    Write-Result -Status "WARN" -Label "Profile Tools:" -Value "Unable to run catalog-based software checks for $BattlestationProfile"
}
else {
    try {
        $SoftwareCatalog = Import-Csv -Path $CatalogFile

        $SelectedSoftwareTools = @(
            $SoftwareCatalog | Where-Object {
                (Test-YesValue (Get-CatalogValue -Row $_ -ColumnName $BattlestationProfile)) -and
                ((Get-CatalogValue -Row $_ -ColumnName "Priority") -eq "Recommended")
            }
        )

        if (-not $SelectedSoftwareTools -or $SelectedSoftwareTools.Count -eq 0) {
            Write-Result -Status "OK" -Label "Profile Tools:" -Value "No recommended software checks for $BattlestationProfile profile" -CountResult:$false
        }
        else {
            Write-Result -Status "OK" -Label "Profile Tools:" -Value "$($SelectedSoftwareTools.Count) recommended software check(s) selected for $BattlestationProfile" -CountResult:$false

            $SoftwareCategories = @(
                $SelectedSoftwareTools |
                    Select-Object -ExpandProperty Category -Unique |
                    Sort-Object
            )

            foreach ($Category in $SoftwareCategories) {
                $CategoryTools = @($SelectedSoftwareTools | Where-Object { $_.Category -eq $Category })

                if (-not $CategoryTools -or $CategoryTools.Count -eq 0) {
                    continue
                }

                Add-ReportLine
                Add-ReportLine -Line "[$Category]"
                Write-Host ""
                Write-Host "[$Category]" -ForegroundColor Gray

                foreach ($Tool in $CategoryTools) {
                    $ToolName = Get-CatalogValue -Row $Tool -ColumnName "Software Name"
                    $DetectionConfig = Get-SoftwareDetectionConfig -CatalogRow $Tool

                    $Installed = Test-SoftwareInstalled `
                        -SoftwareName $Tool."Software Name" `
                        -Commands $Commands `
                        -DisplayNamePatterns $DisplayNamePatterns `
                        -CommonPaths $CommonPaths `
                        -Services $Services

                    if ($Installed) {
                        Write-Result -Status "OK" -Label "$($ToolName):" -Value "Installed"
                    }
                    else {
                        Write-Result -Status "WARN" -Label "$($ToolName):" -Value "Recommended for $BattlestationProfile profile but not found"
                    }
                }
            }
        }
    }
    catch {
        Write-Result -Status "WARN" -Label "Catalog File:" -Value "Unable to read $CatalogFile"
        Write-Result -Status "WARN" -Label "Catalog Error:" -Value $_.Exception.Message
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
    $AntivirusProducts = Get-CimInstance -Namespace "root\SecurityCenter2" -ClassName AntiVirusProduct -ErrorAction Stop

    if ($AntivirusProducts) {
        $AntivirusNames = ($AntivirusProducts.displayName | Sort-Object -Unique) -join ", "
        Write-Result -Status "OK" -Label "Antivirus:" -Value "$AntivirusNames registered"
    }
    else {
        Write-Result -Status "WARN" -Label "Antivirus:" -Value "No registered antivirus provider found"
    }
}
catch {
    Write-Result -Status "WARN" -Label "Antivirus:" -Value "Unable to query antivirus provider"
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

# Windows Update Checks
Write-Section -Title "UPDATES"

try {
    $WindowsUpdateService = Get-CimInstance Win32_Service -Filter "Name='wuauserv'" -ErrorAction Stop

    if ($WindowsUpdateService.StartMode -eq "Disabled") {
        Write-Result -Status "WARN" -Label "Update Service:" -Value "Disabled"
    }
    elseif ($WindowsUpdateService.State -eq "Running") {
        Write-Result -Status "OK" -Label "Update Service:" -Value "$($WindowsUpdateService.StartMode) / Running"
    }
    else {
        Write-Result -Status "OK" -Label "Update Service:" -Value "$($WindowsUpdateService.StartMode) / $($WindowsUpdateService.State) - available on demand"
    }
}
catch {
    Write-Result -Status "WARN" -Label "Update Service:" -Value "Unable to query Windows Update service"
}

try {
    $BitsService = Get-CimInstance Win32_Service -Filter "Name='BITS'" -ErrorAction Stop

    if ($BitsService.StartMode -eq "Disabled") {
        Write-Result -Status "WARN" -Label "BITS Service:" -Value "Disabled"
    }
    elseif ($BitsService.State -eq "Running") {
        Write-Result -Status "OK" -Label "BITS Service:" -Value "$($BitsService.StartMode) / Running"
    }
    else {
        Write-Result -Status "OK" -Label "BITS Service:" -Value "$($BitsService.StartMode) / $($BitsService.State) - available on demand"
    }
}
catch {
    Write-Result -Status "WARN" -Label "BITS Service:" -Value "Unable to query BITS service"
}

try {
    $PendingRebootPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired",
        "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager"
    )

    $PendingReboot = $false

    foreach ($Path in $PendingRebootPaths) {
        if ($Path -eq "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager") {
            $PendingFileRename = Get-ItemProperty -Path $Path -Name "PendingFileRenameOperations" -ErrorAction SilentlyContinue

            if ($PendingFileRename) {
                $PendingReboot = $true
            }
        }
        elseif (Test-Path $Path) {
            $PendingReboot = $true
        }
    }

    if ($PendingReboot) {
        Write-Result -Status "WARN" -Label "Pending Reboot:" -Value "Detected - reboot recommended"
    }
    else {
        Write-Result -Status "OK" -Label "Pending Reboot:" -Value "Not detected"
    }
}
catch {
    Write-Result -Status "WARN" -Label "Pending Reboot:" -Value "Unable to determine reboot status"
}

try {
    $LatestHotFix = Get-HotFix |
        Sort-Object InstalledOn -Descending |
        Select-Object -First 1

    if ($LatestHotFix) {
        Write-Result -Status "OK" -Label "Last Hotfix:" -Value "$($LatestHotFix.HotFixID) installed on $($LatestHotFix.InstalledOn.ToShortDateString())"
    }
    else {
        Write-Result -Status "WARN" -Label "Last Hotfix:" -Value "No hotfix history found"
    }
}
catch {
    Write-Result -Status "WARN" -Label "Last Hotfix:" -Value "Unable to query hotfix history"
}

Write-Summary

Write-Host ""
Write-Host "Health check complete." -ForegroundColor Gray
Write-Host "Report saved to: $ReportFile" -ForegroundColor Gray

Add-ReportLine
Add-ReportLine -Line "Health check complete."
Add-ReportLine -Line "Report saved to: $ReportFile"

$ReportLines | Out-File -FilePath $ReportFile -Encoding UTF8