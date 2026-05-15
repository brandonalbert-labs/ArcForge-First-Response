# ArcForge First Response
# ArcForge First Response Report v0.21

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
$HtmlReportFile = Join-Path $ReportFolder "$ComputerName-$ProfileNameForFile-first-response-$Timestamp.html"
$ReportId = "AFR-$ComputerName-$ProfileNameForFile-$Timestamp"

$CheckCounts = @{
    OK = 0
    WARN = 0
    FAIL = 0
}

$script:ReportLines = [System.Collections.Generic.List[string]]::new()

if (-not (Test-Path $ReportFolder)) {
    New-Item -Path $ReportFolder -ItemType Directory | Out-Null
}

# Adds a line to the in-memory TXT report buffer.
#
# ArcForge writes results to the console immediately, but it also needs to save
# the same information to a TXT report at the end of the run. This helper keeps
# report-writing consistent by appending text to $script:ReportLines.
#
# Input:
# - Line: The exact text to add. Defaults to a blank line when omitted.
# Output:
# - No direct output. Updates the script-scoped ReportLines list.
function Add-ReportLine {
    param (
        [string]$Line = ""
    )

    $script:ReportLines.Add($Line) | Out-Null
}

# Prints one check result to the console and records it in the TXT report.
#
# This is the main output helper used throughout the script. It standardizes the
# [OK] / [WARN] / [FAIL] format, applies console colors, aligns labels, and
# increments the summary counters unless CountResult is disabled.
#
# Input:
# - Status: OK, WARN, FAIL, or another status string.
# - Label: The left-side label shown beside the status.
# - Value: The result details shown after the label.
# - CountResult: Whether this line should affect the final summary totals.
# Output:
# - Writes to the console.
# - Adds the same formatted line to $script:ReportLines.
# - Updates $script:CheckCounts when CountResult is true.
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

# Starts a new report section in both the console and TXT report.
#
# Sections are simple bracketed headings like [SYSTEM], [NETWORK], or [SUMMARY].
# The HTML report later uses these headings to group raw findings into cards.
#
# Input:
# - Title: The section name to display.
# Output:
# - Writes a blank line and section heading to the console.
# - Adds the same section marker to $script:ReportLines.
function Write-Section {
    param (
        [string]$Title
    )

    Write-Host ""
    Write-Host "[$Title]" -ForegroundColor Gray

    Add-ReportLine
    Add-ReportLine -Line "[$Title]"
}

# Builds the final summary section from the accumulated check counters.
#
# This function does not run new health checks. It reads the OK/WARN/FAIL totals
# collected by Write-Result during the script run, then prints an overall status.
#
# Output:
# - Adds the [SUMMARY] section.
# - Writes total checks, passed checks, warnings, failures, and overall status.
# - Uses CountResult:$false so summary lines do not inflate their own totals.
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

# Determines whether a catalog software item appears to be installed.
#
# ArcForge supports several detection methods because Windows software can be
# discovered in different ways: command availability, installed services, common
# executable paths, and uninstall-registry display names. This helper combines
# catalog-provided detection data with a few hand-tuned known-app fallbacks.
#
# Input:
# - SoftwareName: Friendly name from the catalog, such as Chrome or VS Code.
# - Commands: CLI commands to test with Get-Command.
# - DisplayNamePatterns: Registry DisplayName wildcard patterns.
# - CommonPaths: File paths to check with Test-Path.
# - Services: Windows service names to check.
# Output:
# - $true when any detection method finds the software.
# - $false when none of the checks find it.
function Test-SoftwareInstalled {
    param (
        [string]$SoftwareName = "",
        [string[]]$Commands = @(),
        [string[]]$DisplayNamePatterns = @(),
        [string[]]$CommonPaths = @(),
        [string[]]$Services = @()
    )

    # VS Code is handled explicitly because it is commonly installed per-user,
    # system-wide, or exposed through the "code" command. The generic catalog
    # detection can miss one of those install styles.
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
    #
    # This supplements the CSV catalog when Detection Target is too human-readable.
    # Keep this small and focused: it is a fallback layer, not a replacement for
    # maintaining good catalog data.
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

    # Detection pass 1: command lookup.
    # Example: "pwsh", "code", "git", or "ffmpeg" exists in PATH.
    foreach ($Command in $Commands) {
        if (-not [string]::IsNullOrWhiteSpace($Command)) {
            if (Get-Command $Command -ErrorAction SilentlyContinue) {
                return $true
            }
        }
    }

    # Detection pass 2: service lookup.
    # Useful for apps/components that register a Windows service.
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

    # Detection pass 3: common executable paths.
    # Useful when the app exists on disk but is not available as a command.
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

    # Detection pass 4: uninstall registry DisplayName patterns.
    # This catches traditional desktop apps listed in Programs and Features.
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

# Normalizes a catalog cell into a simple yes/no decision.
#
# The software catalog stores profile membership as text. This helper treats a
# value as selected only when it trims down to "yes", case-insensitively.
#
# Input:
# - Value: Any catalog cell value.
# Output:
# - $true if the value is "yes" after trimming/lowercasing; otherwise $false.
function Test-YesValue {
    param (
        [object]$Value
    )

    return (([string]$Value).Trim().ToLower() -eq "yes")
}

# Safely reads one column from a software catalog CSV row.
#
# This avoids errors when a column is missing or blank. It also trims whitespace
# so downstream comparisons are not thrown off by accidental spaces.
#
# Input:
# - Row: One Import-Csv row object.
# - ColumnName: The column/property name to read.
# Output:
# - Trimmed string value when the column exists.
# - Empty string when the column is missing.
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

# Builds registry DisplayName wildcard patterns for software detection.
#
# Windows uninstall entries often use names that are close to, but not exactly,
# the catalog software name. This helper creates a small set of flexible patterns
# from the friendly software name and the detection target.
#
# Examples:
# - "FFmpeg (full)" also produces a pattern for "FFmpeg".
# - A detection target like "matching Google Chrome" produces "*Google Chrome*".
#
# Output:
# - A unique array of wildcard patterns suitable for DisplayName -like checks.
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

# Splits a catalog detection target into individual candidates.
#
# Catalog cells may contain multiple possible detections separated by slashes,
# commas, semicolons, or the word "or". This helper turns that single string into
# a clean list the detection parser can inspect one item at a time.
#
# Input:
# - DetectionTarget: Raw detection target text from the CSV.
# Output:
# - Array of trimmed candidate strings.
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

# Converts one software catalog row into concrete detection instructions.
#
# The CSV is human-readable, while Test-SoftwareInstalled needs structured lists
# of commands, services, paths, and registry patterns. This helper bridges that
# gap by parsing the row and returning a standardized detection config object.
#
# Input:
# - CatalogRow: One Import-Csv software catalog row.
# Output:
# - PSCustomObject with these arrays:
#   - Commands
#   - DisplayNamePatterns
#   - CommonPaths
#   - Services
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


# Groups raw TXT report lines into known report sections.
#
# The HTML report does not run checks again. Instead, it reads the lines already
# captured in $ReportLines and sorts them under major section names. Only known
# sections are captured so accidental bracketed lines do not create random cards.
#
# Input:
# - ReportLines: The full collected report output.
# Output:
# - Hashtable where each known section name maps to a list of lines.
function Get-ArcForgeReportSections {
    param (
        [string[]]$ReportLines
    )

    $KnownSections = @(
        "SYSTEM",
        "UPTIME",
        "PROCESSES",
        "SERVICES",
        "STORAGE",
        "NETWORK",
        "SOFTWARE",
        "SECURITY",
        "UPDATES",
        "SUMMARY"
    )

    $Sections = @{}

    foreach ($Section in $KnownSections) {
        $Sections[$Section] = [System.Collections.Generic.List[string]]::new()
    }

    $CurrentSection = $null

    foreach ($Line in $ReportLines) {
        if ($Line -match '^\[(?<Section>[^\]]+)\]\s*$') {
            $CandidateSection = $Matches.Section.Trim().ToUpperInvariant()

            if ($KnownSections -contains $CandidateSection) {
                $CurrentSection = $CandidateSection
                continue
            }
        }

        if ($CurrentSection -and -not [string]::IsNullOrWhiteSpace($Line)) {
            $Sections[$CurrentSection].Add($Line) | Out-Null
        }
    }

    return $Sections
}

# Generates the self-contained static HTML report.
#
# This function turns the raw report lines and summary counters into a polished
# local HTML file. It does not use JavaScript, external CSS, or external assets.
# The HTML acts as a future GUI prototype while keeping ArcForge simple and local.
#
# Input:
# - OutputPath: Destination .html file path.
# - ReportId, ComputerName, CurrentUser, BattlestationProfile, GeneratedAt:
#   Metadata displayed in the report header/cards.
# - CheckCounts: Final OK/WARN/FAIL totals.
# - ReportLines: Raw TXT report lines used to build sections and findings.
# Output:
# - Writes a complete HTML document to OutputPath.
function New-ArcForgeHtmlReport {
    param (
        [string]$OutputPath,
        [string]$ReportId,
        [string]$ComputerName,
        [string]$CurrentUser,
        [string]$BattlestationProfile,
        [datetime]$GeneratedAt,
        [hashtable]$CheckCounts,
        [string[]]$ReportLines
    )

    # Encodes text before placing it into HTML.
    #
    # This prevents report values containing characters like <, >, or & from
    # breaking the HTML structure or being interpreted as markup.
    function ConvertTo-HtmlSafeText {
        param (
            [string]$Text
        )

        return [System.Net.WebUtility]::HtmlEncode($Text)
    }

    # Converts raw finding lines into an HTML <li> list.
    #
    # Some section line collections can be nested arrays, especially when multiple
    # report sections are combined into one HTML card. This helper flattens them,
    # removes blanks, HTML-encodes every line, and wraps each finding in <code>.
    #
    # Output:
    # - A string containing one or more <li> elements.
    # - A muted placeholder <li> when the section has no findings.
    function ConvertTo-ArcForgeHtmlFindingList {
        param (
            [object[]]$Lines,
            [string]$EmptyMessage = "No findings captured for this section. See Raw Findings for the complete report output."
        )

        $FlattenedLines = @(
            foreach ($Line in $Lines) {
                if ($null -eq $Line) {
                    continue
                }

                if ($Line -is [System.Collections.IEnumerable] -and $Line -isnot [string]) {
                    foreach ($Item in $Line) {
                        if ($null -ne $Item) {
                            [string]$Item
                        }
                    }
                }
                else {
                    [string]$Line
                }
            }
        )

        $CleanLines = @(
            $FlattenedLines |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                ForEach-Object { ConvertTo-HtmlSafeText $_ }
        )

        if (-not $CleanLines -or $CleanLines.Count -eq 0) {
            return "<li class=`"muted`">$(ConvertTo-HtmlSafeText $EmptyMessage)</li>"
        }

        return ($CleanLines | ForEach-Object {
            "<li><code>$_</code></li>"
        }) -join "`n"
    }

    # Flattens nested line arrays into a simple string array.
    #
    # Used by readiness scoring so combined sections like System can be counted
    # the same way as single sections like Network or Security.
    function Get-ArcForgeFlattenedLines {
        param (
            [object[]]$Lines
        )

        return @(
            foreach ($Line in $Lines) {
                if ($null -eq $Line) {
                    continue
                }

                if ($Line -is [System.Collections.IEnumerable] -and $Line -isnot [string]) {
                    foreach ($Item in $Line) {
                        if ($null -ne $Item) {
                            [string]$Item
                        }
                    }
                }
                else {
                    [string]$Line
                }
            }
        )
    }

    # Scores one report area for the Readiness Overview cards.
    #
    # This function counts how many [OK], [WARN], and [FAIL] lines exist in a
    # section, then assigns the card status shown in the HTML dashboard.
    #
    # Output:
    # - PSCustomObject containing name, status label, CSS class, counts, and summary.
    function Get-ArcForgeSectionReadiness {
        param (
            [string]$Name,
            [object[]]$Lines
        )

        $FlattenedLines = Get-ArcForgeFlattenedLines -Lines $Lines

        $OkCount = @($FlattenedLines | Where-Object { $_ -match '^\[OK\]' }).Count
        $WarnCount = @($FlattenedLines | Where-Object { $_ -match '^\[WARN\]' }).Count
        $FailCount = @($FlattenedLines | Where-Object { $_ -match '^\[FAIL\]' }).Count

        if ($FailCount -gt 0) {
            $Status = "Critical"
            $StatusClass = "readiness-critical"
            $Summary = "Critical findings require attention."
        }
        elseif ($WarnCount -gt 0) {
            $Status = "Attention"
            $StatusClass = "readiness-attention"
            $Summary = "Warnings found. Review recommended actions."
        }
        elseif ($OkCount -gt 0) {
            $Status = "OK"
            $StatusClass = "readiness-ok"
            $Summary = "All checks passed."
        }
        else {
            $Status = "No Data"
            $StatusClass = "readiness-neutral"
            $Summary = "No findings detected in this section."
        }

        [pscustomobject]@{
            Name        = $Name
            Status      = $Status
            StatusClass = $StatusClass
            OkCount     = $OkCount
            WarnCount   = $WarnCount
            FailCount   = $FailCount
            Summary     = $Summary
        }
    }

    # Builds the HTML block for the Readiness Overview dashboard cards.
    #
    # The card data is prepared by Get-ArcForgeSectionReadiness. This helper only
    # converts those objects into HTML markup for the final report.
    function New-ArcForgeReadinessOverviewHtml {
        param (
            [object[]]$ReadinessCards
        )

        $CardBlocks = @()

        foreach ($Card in $ReadinessCards) {
            $SafeName = ConvertTo-HtmlSafeText $Card.Name
            $SafeStatus = ConvertTo-HtmlSafeText $Card.Status
            $SafeSummary = ConvertTo-HtmlSafeText $Card.Summary

            $CardBlocks += @"
            <article class="readiness-card $($Card.StatusClass)">
                <div class="readiness-card-header">
                    <h3>$SafeName</h3>
                    <span class="readiness-status">$SafeStatus</span>
                </div>
                <div class="readiness-counts">
                    <span><strong>$($Card.OkCount)</strong> OK</span>
                    <span><strong>$($Card.WarnCount)</strong> WARN</span>
                    <span><strong>$($Card.FailCount)</strong> FAIL</span>
                </div>
                <p>$SafeSummary</p>
            </article>
"@
        }

        $CardsHtml = $CardBlocks -join "`n"

        return @"
        <section id="readiness-overview" class="card section">
            <div class="section-title">
                <h2>Readiness Overview</h2>
                <p>Dashboard-style summary of major battlestation readiness areas.</p>
            </div>
            <div class="readiness-grid">
$CardsHtml
            </div>
        </section>
"@
    }

    # Builds the compact v0.20 Endpoint Summary badge shown at the top of the
    # HTML report.
    #
    # Why this exists:
    # - The first screen of the HTML report should feel like a quick endpoint
    #   identity card, similar to what a tech might scan in an RMM tool.
    # - A tech should be able to immediately identify the computer, current user,
    #   selected Battlestation Profile, generated time, overall status, and check
    #   totals without reading a large grid of separate cards.
    #
    # Input:
    # - Report metadata that was already being displayed in the Report Summary.
    # - Existing OK/WARN/FAIL/Total counters and the existing OverallStatus value.
    #
    # Output:
    # - A presentation-only HTML block inserted into the Report Summary section.
    #
    # Important:
    # - This does not run checks.
    # - This does not change scoring.
    # - This does not change console output.
    # - This does not change TXT report output.
    #
    # Troubleshooting rule:
    # - If the top badge looks wrong, inspect the endpoint-summary CSS classes and
    #   the values passed into this helper before changing health-check logic.
    function New-ArcForgeEndpointSummaryHtml {
        param (
            [string]$ReportId,
            [string]$ComputerName,
            [string]$CurrentUser,
            [string]$BattlestationProfile,
            [datetime]$GeneratedAt,
            [string]$OverallStatus,
            [string]$StatusClass,
            [int]$OkCount,
            [int]$WarnCount,
            [int]$FailCount,
            [int]$TotalChecks
        )

        $SafeReportId = ConvertTo-HtmlSafeText $ReportId
        $SafeComputerName = ConvertTo-HtmlSafeText $ComputerName
        $SafeCurrentUser = ConvertTo-HtmlSafeText $CurrentUser
        $SafeBattlestationProfile = ConvertTo-HtmlSafeText $BattlestationProfile
        $SafeGeneratedAt = ConvertTo-HtmlSafeText $GeneratedAt
        $SafeOverallStatus = ConvertTo-HtmlSafeText $OverallStatus

        return @"
        <section class="endpoint-summary card" aria-label="Endpoint Summary">
            <div class="endpoint-main">
                <div class="endpoint-identity">
                    <div class="endpoint-kicker">Endpoint Summary</div>
                    <div class="endpoint-title-row">
                        <h2>$SafeComputerName</h2>
                        <span class="endpoint-status-pill $StatusClass">$SafeOverallStatus</span>
                    </div>
                    <p class="endpoint-description">Static local triage snapshot for this workstation.</p>

                    <div class="endpoint-meta-stack">
                        <div class="endpoint-meta-item">
                            <span class="endpoint-meta-label">Current User</span>
                            <strong>$SafeCurrentUser</strong>
                        </div>
                        <div class="endpoint-meta-item">
                            <span class="endpoint-meta-label">Battlestation Profile</span>
                            <strong>$SafeBattlestationProfile</strong>
                        </div>
                    </div>
                </div>

                <div class="endpoint-status-panel">
                    <div class="endpoint-status-heading">Check Totals</div>
                    <div class="endpoint-count-strip" aria-label="ArcForge check result totals">
                        <div class="endpoint-count endpoint-count-ok">
                            <span class="endpoint-count-number">$OkCount</span>
                            <span class="endpoint-count-label">OK</span>
                        </div>
                        <div class="endpoint-count endpoint-count-warn">
                            <span class="endpoint-count-number">$WarnCount</span>
                            <span class="endpoint-count-label">WARN</span>
                        </div>
                        <div class="endpoint-count endpoint-count-fail">
                            <span class="endpoint-count-number">$FailCount</span>
                            <span class="endpoint-count-label">FAIL</span>
                        </div>
                        <div class="endpoint-count endpoint-count-total">
                            <span class="endpoint-count-number">$TotalChecks</span>
                            <span class="endpoint-count-label">TOTAL</span>
                        </div>
                    </div>
                </div>
            </div>

            <div class="endpoint-bottom-row">
                <div class="endpoint-meta-item">
                    <span class="endpoint-meta-label">Report ID</span>
                    <strong>$SafeReportId</strong>
                </div>
                <div class="endpoint-meta-item">
                    <span class="endpoint-meta-label">Generated At</span>
                    <strong>$SafeGeneratedAt</strong>
                </div>
            </div>
        </section>
"@
    }


    # Assigns one WARN/FAIL finding to the triage bucket shown in Recommended Actions.
    #
    # Why this exists:
    # - The raw report lines are useful, but a long flat list is hard to scan.
    # - This helper lets the HTML report group problems into practical buckets,
    #   almost like a small ticket queue.
    # - It only affects the HTML Recommended Actions section. It does not change
    #   the checks themselves, the console output, or the TXT report.
    #
    # How it works:
    # - PowerShell's -match operator checks whether the finding contains certain
    #   words or phrases.
    # - The first matching bucket wins because each match immediately returns.
    # - Anything that does not match a known pattern falls back to General Findings.
    #
    # Input:
    # - Finding: One raw report line such as:
    #   [WARN]    DNS: Resolution failed
    #
    # Output:
    # - A category name used as a heading in the Recommended Actions panel.
    function Get-ArcForgeActionCategory {
        param (
            [string]$Finding
        )

        # Missing profile software and catalog problems belong together because
        # they affect whether the selected Battlestation Profile is fully ready.
        if ($Finding -match 'recommended for .+ profile but not found|Profile Tools|Catalog File|Catalog Error') {
            return "Profile Readiness"
        }

        # Network-related findings are grouped together so internet, gateway, DNS,
        # and adapter issues are easy to review in one place.
        if ($Finding -match 'Gateway|Internet Ping|DNS|IP Address|Network Adapter') {
            return "Network Connectivity"
        }

        # Security posture findings are grouped separately because they usually
        # need a manual review rather than an immediate technical repair.
        if ($Finding -match 'Local Admins|Firewall|Antivirus|Defender') {
            return "Security Review"
        }

        # Pending reboot could also be considered system stability, but for this
        # report it is more useful under Update Readiness because reboots often
        # block patching, software installs, and troubleshooting.
        if ($Finding -match 'Pending Reboot|Windows Update|Update Service|BITS Service|Last Hotfix|Hotfix') {
            return "Update Readiness"
        }

        # System-level findings are things that affect the local workstation's
        # stability or day-to-day readiness.
        if ($Finding -match 'Uptime|Hung Apps|Processes|Services|Storage|Disk') {
            return "System Stability"
        }

        # Safe fallback. This prevents a finding from disappearing just because
        # it did not match one of the known patterns above.
        return "General Findings"
    }

    # Gives each action item a short, beginner-friendly next step.
    #
    # Why this exists:
    # - A finding tells the user what ArcForge noticed.
    # - A suggested action tells the user what to do next.
    # - Keeping this in one helper makes the wording easier to improve later.
    #
    # Important:
    # - These are intentionally simple Tier-1 style suggestions.
    # - The detailed technical evidence still lives in the normal section cards
    #   and in Raw Findings.
    # - This helper does not fix anything automatically.
    #
    # Input:
    # - Finding: One raw WARN/FAIL report line.
    # - BattlestationProfile: The selected profile, such as Developer or Gaming.
    #
    # Output:
    # - A plain English remediation suggestion displayed under the action item.
    function Get-ArcForgeSuggestedAction {
        param (
            [string]$Finding,
            [string]$BattlestationProfile
        )

        # Long uptime can make troubleshooting noisy because a reboot may clear
        # pending updates, stale services, driver weirdness, or old hung processes.
        if ($Finding -match 'Uptime') {
            return "Reboot during the next maintenance window, then rerun ArcForge."
        }

        # Hung apps are usually best handled by closing/restarting the affected
        # apps before assuming the whole workstation is unhealthy.
        if ($Finding -match 'Hung Apps') {
            return "Close or restart the affected apps, then rerun ArcForge."
        }

        # A network warning could come from several checks, so this suggestion
        # points the user toward the most likely basic troubleshooting areas.
        if ($Finding -match 'Gateway|Internet Ping|DNS|IP Address|Network Adapter') {
            return "Verify network connectivity, adapter configuration, DNS settings, and gateway reachability."
        }

        # Local administrator membership is not automatically bad, but it should
        # be intentional. This keeps the action worded as a review item.
        if ($Finding -match 'Local Admins') {
            return "Confirm all local administrator accounts are expected."
        }

        # Firewall checks can be satisfied by Windows Firewall or a valid third-
        # party firewall, so the suggestion avoids assuming Defender is the only answer.
        if ($Finding -match 'Firewall') {
            return "Verify Windows Firewall or the active third-party firewall is enabled."
        }

        # Antivirus can also be provided by third-party tools, so this suggestion
        # asks the user to confirm the expected provider is healthy.
        if ($Finding -match 'Antivirus|Defender') {
            return "Confirm an expected antivirus provider is installed, enabled, and reporting healthy."
        }

        # Pending reboot is often the first thing to resolve before troubleshooting
        # Windows Update, installers, or other system changes.
        if ($Finding -match 'Pending Reboot') {
            return "Reboot before continuing patching, installs, or troubleshooting."
        }

        # Windows Update-related warnings are grouped around patch readiness.
        if ($Finding -match 'Windows Update|Update Service|BITS Service|Last Hotfix|Hotfix') {
            return "Review Windows Update readiness before patching or installing additional software."
        }

        # Catalog-level issues are different from missing apps. They may mean the
        # runtime CSV is missing, broken, or not matching the selected profile.
        if ($Finding -match 'Profile Tools|Catalog File|Catalog Error') {
            return "Review the ArcForge Software Catalog and confirm the selected Battlestation Profile is using the expected runtime catalog."
        }

        # Safe fallback for any current or future WARN/FAIL line that does not
        # match the more specific patterns above.
        return "Review the related section for details, then rerun ArcForge after remediation."
    }

    # Converts raw WARN/FAIL report lines into grouped action objects.
    #
    # Why this exists:
    # - The checks currently write human-readable report lines, not structured
    #   objects. That is fine for now.
    # - This function acts as a thin presentation adapter. It reads those existing
    #   lines and prepares clean objects for the HTML action panel.
    # - This avoids refactoring the check engine during v0.17.
    #
    # What gets filtered out:
    # - The SUMMARY section includes lines like "Warnings:" and "Failures:".
    #   Those are totals, not actionable findings, so they are excluded.
    # - "Overall Status" is also excluded because it is a summary label, not a
    #   specific repair item.
    #
    # Special software handling:
    # - A profile like Developer can have many missing recommended tools.
    # - Listing every missing tool inside Recommended Actions makes the action
    #   queue noisy.
    # - This function collects those missing-tool warnings and replaces them with
    #   one summarized Profile Readiness action.
    # - The full missing-tools list is still preserved in Software Readiness and
    #   Raw Findings.
    #
    # Input:
    # - ReportLines: The complete in-memory TXT report lines.
    # - BattlestationProfile: The selected profile name.
    #
    # Output:
    # - PSCustomObject items with Category, Severity, Title, Detail, and SuggestedAction.
    function Get-ArcForgeActionItems {
        param (
            [string[]]$ReportLines,
            [string]$BattlestationProfile
        )

        # Start with every WARN/FAIL line from the finished report.
        # Then remove summary counters so the action queue only shows real issues.
        $FindingLines = @(
            $ReportLines |
                Where-Object {
                    $_ -match '^\[(WARN|FAIL)\]' -and
                    $_ -notmatch '^\[(WARN|FAIL)\]\s+Warnings:' -and
                    $_ -notmatch '^\[(WARN|FAIL)\]\s+Failures:' -and
                    $_ -notmatch '^\[(WARN|FAIL)\]\s+Overall Status:'
                }
        )

        # ActionItems will hold the final objects that become cards in HTML.
        $ActionItems = [System.Collections.Generic.List[object]]::new()

        # MissingProfileTools temporarily stores software warnings that should be
        # summarized into one action instead of displayed one-by-one.
        $MissingProfileTools = [System.Collections.Generic.List[string]]::new()

        foreach ($Finding in $FindingLines) {
            # Detect missing recommended profile tools.
            #
            # Example matched line:
            # [WARN]    Git: Recommended for Developer profile but not found
            #
            # Named regex groups like (?<Severity>WARN|FAIL) make the pattern
            # easier to understand later, even though this specific branch only
            # needs to count the matching lines.
            if ($Finding -match '^\[(?<Severity>WARN|FAIL)\]\s+(?<ToolName>.+?):\s+Recommended for (?<Profile>.+?) profile but not found$') {
                $MissingProfileTools.Add($Finding) | Out-Null
                continue
            }

            # Default to WARN so a malformed line still has a safe visual style.
            # If the line begins with [WARN] or [FAIL], use that real severity.
            $Severity = "WARN"
            if ($Finding -match '^\[(?<Severity>WARN|FAIL)\]') {
                $Severity = $Matches.Severity
            }

            # Remove the leading [WARN] or [FAIL] tag from the title because the
            # visual badge already shows severity in the HTML action card.
            $Title = $Finding -replace '^\[(WARN|FAIL)\]\s+', ''

            # Build one normalized action object for the HTML renderer.
            # The renderer does not need to know how the item was detected; it
            # only needs these simple fields.
            $ActionItems.Add([pscustomobject]@{
                Category        = Get-ArcForgeActionCategory -Finding $Finding
                Severity        = $Severity
                Title           = $Title
                Detail          = ""
                SuggestedAction = Get-ArcForgeSuggestedAction -Finding $Finding -BattlestationProfile $BattlestationProfile
            }) | Out-Null
        }

        # Add one summarized software readiness action after all findings have
        # been scanned. This keeps Recommended Actions readable while preserving
        # the complete details elsewhere in the report.
        if ($MissingProfileTools.Count -gt 0) {
            $ActionItems.Add([pscustomobject]@{
                Category        = "Profile Readiness"
                Severity        = "WARN"
                Title           = "$($MissingProfileTools.Count) recommended $BattlestationProfile profile tools are missing."
                Detail          = "Software Readiness contains the full missing-tools list."
                SuggestedAction = "Review Software Readiness before using this workstation for $BattlestationProfile work."
            }) | Out-Null
        }

        # Return as an array so the caller can safely count/filter the result,
        # even when there is only one action item.
        return @($ActionItems)
    }

    # Builds the grouped Recommended Actions triage panel.
    #
    # Why this exists:
    # - Get-ArcForgeActionItems prepares clean action objects.
    # - This function turns those objects into static HTML.
    # - Separating data preparation from HTML generation makes future polishing
    #   easier without changing the report parsing logic.
    #
    # Important:
    # - The output is static HTML only.
    # - No JavaScript is used.
    # - No external dependencies are used.
    # - This only changes the HTML Recommended Actions section.
    #
    # Input:
    # - ActionItems: Objects returned by Get-ArcForgeActionItems.
    #
    # Output:
    # - A string containing the complete Recommended Actions <section> block.
    function New-ArcForgeRecommendedActionsHtml {
        param (
            [object[]]$ActionItems
        )

        # If there are no WARN/FAIL action items, show a clean healthy-state card
        # instead of leaving the section blank.
        if (-not $ActionItems -or $ActionItems.Count -eq 0) {
            return @"
        <section id="recommended-actions" class="card section">
            <div class="section-title">
                <h2>Recommended Actions</h2>
                <p>No immediate recommended actions. System appears healthy based on current checks.</p>
            </div>
        </section>
"@
        }

        # Controls the order of groups in the HTML report.
        #
        # The order is intentionally not alphabetical. It follows a practical
        # triage flow: local stability, network, profile tools, security, updates,
        # then anything uncategorized.
        $CategoryOrder = @(
            "System Stability",
            "Network Connectivity",
            "Profile Readiness",
            "Security Review",
            "Update Readiness",
            "General Findings"
        )

        $GroupBlocks = @()

        foreach ($Category in $CategoryOrder) {
            # Pull only the items for the current category.
            # Wrapping the result in @() makes .Count reliable even if there is
            # only one matching item.
            $CategoryItems = @($ActionItems | Where-Object { $_.Category -eq $Category })

            # Skip empty categories so the report only shows useful groups.
            if (-not $CategoryItems -or $CategoryItems.Count -eq 0) {
                continue
            }

            $ItemBlocks = @()

            foreach ($Item in $CategoryItems) {
                # Build a CSS class from the severity, such as action-warn or
                # action-fail. ToLowerInvariant avoids locale-specific casing.
                $SeverityClass = "action-$($Item.Severity.ToLowerInvariant())"

                # Always HTML-encode values before inserting them into the HTML.
                # This protects the report if a finding contains characters like
                # <, >, &, quotes, or other markup-looking text.
                $SafeSeverity = ConvertTo-HtmlSafeText $Item.Severity
                $SafeTitle = ConvertTo-HtmlSafeText $Item.Title
                $SafeDetail = ConvertTo-HtmlSafeText $Item.Detail
                $SafeSuggestedAction = ConvertTo-HtmlSafeText $Item.SuggestedAction

                # Detail is optional. Most findings do not need a second detail
                # line, but the summarized software action uses it to point back
                # to the full Software Readiness list.
                $DetailHtml = ""
                if (-not [string]::IsNullOrWhiteSpace($Item.Detail)) {
                    $DetailHtml = "<p class=`"action-detail`">$SafeDetail</p>"
                }

                # This is the individual ticket-style action card.
                # The severity badge and left border provide quick visual context.
                $ItemBlocks += @"
                    <article class="action-item $SeverityClass">
                        <div class="action-header">
                            <span class="action-severity">$SafeSeverity</span>
                            <strong>$SafeTitle</strong>
                        </div>
                        $DetailHtml
                        <p class="action-suggestion"><strong>Suggested Action:</strong> $SafeSuggestedAction</p>
                    </article>
"@
            }

            $SafeCategory = ConvertTo-HtmlSafeText $Category
            $ItemsHtml = $ItemBlocks -join "`n"

            # This is the group wrapper, such as System Stability or Security Review.
            $GroupBlocks += @"
                <div class="action-group">
                    <h3>$SafeCategory</h3>
$ItemsHtml
                </div>
"@
        }

        $GroupsHtml = $GroupBlocks -join "`n"

        # Final Recommended Actions section inserted into the main HTML template.
        return @"
        <section id="recommended-actions" class="card section">
            <div class="section-title">
                <h2>Recommended Actions</h2>
                <p>Grouped WARN/FAIL findings prioritized as a local triage queue.</p>
            </div>
            <div class="action-queue">
$GroupsHtml
            </div>
        </section>
"@
    }

    # Builds the three compact status segments shown beside readiness-domain
    # links in the HTML report sidebar.
    #
    # Why this exists:
    # - v0.19 is still a presentation-layer release.
    # - The sidebar segments give the report a quick dashboard-style glance
    #   without adding JavaScript, external dependencies, or a new GUI layer.
    # - This helper only converts an existing readiness status into small HTML
    #   spans. It does not inspect the computer or rerun any health checks.
    #
    # Input:
    # - A readiness status from Get-ArcForgeSectionReadiness:
    #   Critical, Attention, OK, or No Data.
    #
    # Output:
    # - A string containing three small <span> elements.
    #
    # Important:
    # - This is presentation-only.
    # - This does not change check logic, scoring, console output, or TXT output.
    # - The sidebar status should always match the Readiness Overview card that
    #   was built from the same readiness object.
    function New-ArcForgeSidebarStatusSegmentsHtml {
        param (
            [string]$Status
        )

        $SegmentClass = "sidebar-segment-empty"
        $FilledSegments = 0

        switch ($Status) {
            "Critical" {
                $SegmentClass = "sidebar-segment-critical"
                $FilledSegments = 1
            }
            "Attention" {
                $SegmentClass = "sidebar-segment-attention"
                $FilledSegments = 2
            }
            "OK" {
                $SegmentClass = "sidebar-segment-ok"
                $FilledSegments = 3
            }
            default {
                $SegmentClass = "sidebar-segment-empty"
                $FilledSegments = 0
            }
        }

        $Segments = @()

        for ($Index = 1; $Index -le 3; $Index++) {
            if ($Index -le $FilledSegments) {
                $Segments += "<span class=""sidebar-segment $SegmentClass""></span>"
            }
            else {
                $Segments += "<span class=""sidebar-segment sidebar-segment-empty""></span>"
            }
        }

        return ($Segments -join "")
    }

    # Builds the static sidebar navigation used by the HTML report.
    #
    # Why this exists:
    # - v0.18 added quick-jump navigation for the major report sections.
    # - v0.19 reuses the existing Readiness Overview data to add small status
    #   segments beside the five primary readiness domains only.
    # - This helper keeps the navigation markup in one small place instead of
    #   scattering repeated <a> tags throughout the main HTML template.
    #
    # Input:
    # - ReadinessCards are the same objects used by New-ArcForgeReadinessOverviewHtml.
    # - The cards are calculated once, then reused by both the Readiness Overview
    #   and this sidebar navigation. That keeps both views in sync.
    #
    # Important:
    # - These are normal internal anchor links like href="#network".
    # - The status segments are visual/presentation-only.
    # - No JavaScript is used.
    # - No external dependencies are used.
    # - This does not change any check logic, console output, or TXT output.
    #
    # Troubleshooting rule:
    # - Every href="#section-name" in this helper must match an id="section-name"
    #   somewhere in the HTML template below.
    # - Sidebar status segments should only appear for System, Network,
    #   Software Readiness, Security, and Updates.
    # - If a segment does not match the Readiness Overview card, inspect the
    #   readiness card Name values first.
    #
    # Output:
    # - A string containing the complete sidebar <aside> block.
    function New-ArcForgeReportNavigationHtml {
        param (
            [object[]]$ReadinessCards
        )

        $ReadinessByName = @{}

        foreach ($Card in $ReadinessCards) {
            $ReadinessByName[$Card.Name] = $Card
        }

        $NavigationItems = @(
            [pscustomobject]@{ Label = "Report Summary";       Anchor = "report-summary";       ShowStatus = $false }
            [pscustomobject]@{ Label = "Incident Summary";     Anchor = "incident-summary";     ShowStatus = $false }
            [pscustomobject]@{ Label = "Readiness Overview";   Anchor = "readiness-overview";   ShowStatus = $false }
            [pscustomobject]@{ Label = "System";               Anchor = "system";               ShowStatus = $true  }
            [pscustomobject]@{ Label = "Network";              Anchor = "network";              ShowStatus = $true  }
            [pscustomobject]@{ Label = "Software Readiness";   Anchor = "software-readiness";   ShowStatus = $true  }
            [pscustomobject]@{ Label = "Security";             Anchor = "security";             ShowStatus = $true  }
            [pscustomobject]@{ Label = "Updates";              Anchor = "updates";              ShowStatus = $true  }
            [pscustomobject]@{ Label = "Recommended Actions";  Anchor = "recommended-actions";  ShowStatus = $false }
            [pscustomobject]@{ Label = "Raw Findings";         Anchor = "raw-findings";         ShowStatus = $false }
        )

        $NavigationLinks = @()

        foreach ($Item in $NavigationItems) {
            $SafeLabel = ConvertTo-HtmlSafeText $Item.Label
            $SafeAnchor = ConvertTo-HtmlSafeText $Item.Anchor

            if ($Item.ShowStatus -and $ReadinessByName.ContainsKey($Item.Label)) {
                $Card = $ReadinessByName[$Item.Label]
                $SafeStatus = ConvertTo-HtmlSafeText $Card.Status
                $SegmentsHtml = New-ArcForgeSidebarStatusSegmentsHtml -Status $Card.Status

                $NavigationLinks += @"
                <a class="sidebar-link sidebar-link-with-status" href="#$SafeAnchor" title="$SafeLabel readiness: $SafeStatus" aria-label="$SafeLabel readiness: $SafeStatus">
                    <span class="sidebar-link-label">$SafeLabel</span>
                    <span class="sidebar-status-segments" aria-hidden="true">$SegmentsHtml</span>
                </a>
"@
            }
            else {
                $NavigationLinks += "                <a class=""sidebar-link"" href=""#$SafeAnchor"">$SafeLabel</a>"
            }
        }

        $NavigationLinksHtml = $NavigationLinks -join "`n"

        return @"
        <aside class="report-sidebar">
            <div class="sidebar-title">Report Navigation</div>
            <div class="sidebar-subtitle">Jump to a major report section.</div>
            <nav class="sidebar-nav" aria-label="ArcForge report sections">
$NavigationLinksHtml
            </nav>
        </aside>
"@
    }

    $ReportSections = Get-ArcForgeReportSections -ReportLines $ReportLines

    $SystemLines = @(
        $ReportSections["SYSTEM"]
        $ReportSections["UPTIME"]
        $ReportSections["PROCESSES"]
        $ReportSections["SERVICES"]
        $ReportSections["STORAGE"]
    )

    $NetworkLines = $ReportSections["NETWORK"]
    $SoftwareLines = $ReportSections["SOFTWARE"]
    $SecurityLines = $ReportSections["SECURITY"]
    $UpdatesLines = $ReportSections["UPDATES"]

    $SystemFindingsHtml = ConvertTo-ArcForgeHtmlFindingList -Lines $SystemLines
    $NetworkFindingsHtml = ConvertTo-ArcForgeHtmlFindingList -Lines $NetworkLines
    $SoftwareFindingsHtml = ConvertTo-ArcForgeHtmlFindingList -Lines $SoftwareLines
    $SecurityFindingsHtml = ConvertTo-ArcForgeHtmlFindingList -Lines $SecurityLines
    $UpdatesFindingsHtml = ConvertTo-ArcForgeHtmlFindingList -Lines $UpdatesLines

    # Build the shared readiness card data once.
    #
    # Why this exists:
    # - The Readiness Overview cards and the v0.19 sidebar readiness segments
    #   should describe the same five readiness domains.
    # - Calculating these objects once prevents the sidebar and overview cards
    #   from drifting out of sync later.
    #
    # Important:
    # - This reuses the existing readiness helper.
    # - This does not rerun health checks.
    # - This does not change console output or TXT report output.
    $ReadinessCards = @(
        Get-ArcForgeSectionReadiness -Name "System" -Lines $SystemLines
        Get-ArcForgeSectionReadiness -Name "Network" -Lines $NetworkLines
        Get-ArcForgeSectionReadiness -Name "Software Readiness" -Lines $SoftwareLines
        Get-ArcForgeSectionReadiness -Name "Security" -Lines $SecurityLines
        Get-ArcForgeSectionReadiness -Name "Updates" -Lines $UpdatesLines
    )

    $ReadinessOverviewHtml = New-ArcForgeReadinessOverviewHtml -ReadinessCards $ReadinessCards

    $OkCount = $CheckCounts.OK
    $WarnCount = $CheckCounts.WARN
    $FailCount = $CheckCounts.FAIL
    $TotalChecks = $OkCount + $WarnCount + $FailCount

    if ($FailCount -gt 0) {
        $OverallStatus = "Action Required"
        $StatusClass = "status-fail"
    }
    elseif ($WarnCount -gt 0) {
        $OverallStatus = "Attention Recommended"
        $StatusClass = "status-warn"
    }
    else {
        $OverallStatus = "Healthy"
        $StatusClass = "status-ok"
    }

    # Build the v0.20 Endpoint Summary badge.
    #
    # This reuses the metadata and counts already calculated for the report. It is
    # presentation-only and only affects the static HTML report's top summary area.
    $EndpointSummaryHtml = New-ArcForgeEndpointSummaryHtml `
        -ReportId $ReportId `
        -ComputerName $ComputerName `
        -CurrentUser $CurrentUser `
        -BattlestationProfile $BattlestationProfile `
        -GeneratedAt $GeneratedAt `
        -OverallStatus $OverallStatus `
        -StatusClass $StatusClass `
        -OkCount $OkCount `
        -WarnCount $WarnCount `
        -FailCount $FailCount `
        -TotalChecks $TotalChecks

    # Build the v0.17 Recommended Actions queue.
    #
    # Step 1: Convert raw WARN/FAIL report lines into simple action objects.
    # Step 2: Convert those action objects into grouped static HTML.
    #
    # This happens after the overall counts/status are calculated because the
    # action queue depends on the completed report output.
    $RecommendedActionItems = Get-ArcForgeActionItems -ReportLines $ReportLines -BattlestationProfile $BattlestationProfile
    $RecommendedActionsHtml = New-ArcForgeRecommendedActionsHtml -ActionItems $RecommendedActionItems

    # Build the v0.19 static report navigation.
    #
    # This creates the sidebar HTML once, then the main template inserts it beside
    # the report content. The same readiness objects used by the Readiness Overview
    # are passed in here so the sidebar can show compact presentation-only status
    # segments without recalculating any checks.
    $ReportNavigationHtml = New-ArcForgeReportNavigationHtml -ReadinessCards $ReadinessCards

    $RawFindings = ConvertTo-HtmlSafeText ($ReportLines -join "`r`n")

    $Html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>ArcForge First Response Report - $(ConvertTo-HtmlSafeText $ReportId)</title>
    <style>
        :root {
            --bg: #f4f6f8;
            --panel: #ffffff;
            --border: #d9dee5;
            --text: #1f2933;
            --muted: #65758b;
            --ok: #1f8f4d;
            --warn: #b7791f;
            --fail: #c53030;
            --header: #111827;
            --chip: #eef2f7;
        }

        body {
            margin: 0;
            padding: 32px;
            background: var(--bg);
            color: var(--text);
            font-family: "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, sans-serif;
            line-height: 1.5;
        }

        /* v0.18 report layout shell.
           The report now has two presentation-only columns on desktop:
           a sidebar navigation panel on the left and the existing report content
           on the right. This is still a static local HTML report, not a GUI. */
        .report-shell {
            max-width: 1380px;
            margin: 0 auto;
            display: grid;
            grid-template-columns: 260px minmax(0, 1fr);
            gap: 24px;
            align-items: start;
        }

        /* Main report column.
           min-width: 0 prevents long code/raw findings text from forcing the
           grid wider than the browser window. */
        .report-main {
            min-width: 0;
        }

        /* v0.18 sidebar navigation card.
           position: sticky keeps the quick links visible while scrolling on
           desktop. It is safe because it is pure CSS and does not require JS. */
        .report-sidebar {
            position: sticky;
            top: 24px;
            background: var(--panel);
            border: 1px solid var(--border);
            border-radius: 14px;
            padding: 16px;
            box-shadow: 0 2px 8px rgba(15, 23, 42, 0.05);
        }

        .sidebar-title {
            font-weight: 700;
            margin-bottom: 4px;
        }

        .sidebar-subtitle {
            color: var(--muted);
            font-size: 13px;
            margin-bottom: 14px;
        }

        .sidebar-nav {
            display: grid;
            gap: 8px;
        }

        .sidebar-link {
            color: var(--text);
            text-decoration: none;
            border: 1px solid transparent;
            border-radius: 10px;
            padding: 9px 10px;
            font-size: 14px;
            font-weight: 600;
        }

        .sidebar-link:hover,
        .sidebar-link:focus {
            background: var(--chip);
            border-color: var(--border);
        }

        /* v0.19 sidebar readiness segments.
           Why this exists:
           - These tiny segments make the sidebar act more like a static triage
             dashboard while keeping the report local, self-contained, and simple.
           - They are presentation-only and reuse the same readiness data shown in
             the Readiness Overview cards.

           Important:
           - No JavaScript is involved.
           - Empty segments are muted placeholders.
           - Filled segments map to the existing readiness status:
             Critical = 1 filled segment, Attention = 2, OK = 3, No Data = 0. */
        .sidebar-link-with-status {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 10px;
        }

        .sidebar-link-label {
            min-width: 0;
        }

        .sidebar-status-segments {
            display: inline-flex;
            align-items: center;
            gap: 3px;
            flex-shrink: 0;
        }

        .sidebar-segment {
            width: 8px;
            height: 14px;
            border-radius: 3px;
            display: inline-block;
        }

        .sidebar-segment-empty {
            background: #dbe2ea;
        }

        .sidebar-segment-ok {
            background: var(--ok);
        }

        .sidebar-segment-attention {
            background: var(--warn);
        }

        .sidebar-segment-critical {
            background: var(--fail);
        }

        /* ============================================================
           v0.18 HTML ANCHOR TARGET SPACING
           ============================================================
           These rules help the sidebar links land cleanly.

           How the sidebar jump works:
           - A sidebar link such as href="#report-summary" looks for a matching
             HTML element with id="report-summary".
           - The browser handles that jump automatically.
           - scroll-margin-top gives the jump target a little breathing room so
             the section does not land too tightly against the top of the window.

           Troubleshooting rule:
           - If a sidebar link changes the browser URL but does not visibly jump,
             confirm the matching id exists in the HTML template.
           ============================================================ */
        .section,
        .report-summary {
            scroll-margin-top: 24px;
        }

        .ticket-header {
            background: var(--header);
            color: white;
            border-radius: 14px;
            padding: 24px 28px;
            margin-bottom: 20px;
            box-shadow: 0 8px 24px rgba(15, 23, 42, 0.16);
        }

        .eyebrow {
            color: #aeb8c7;
            font-size: 13px;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            margin-bottom: 6px;
        }

        h1 {
            margin: 0;
            font-size: 28px;
            font-weight: 700;
        }

        .subtitle {
            margin-top: 8px;
            color: #d6dce8;
        }

        .status-row {
            margin-top: 18px;
        }

        .status-badge {
            display: inline-block;
            border-radius: 999px;
            padding: 7px 13px;
            font-weight: 700;
            font-size: 13px;
        }

        .status-ok {
            background: rgba(31, 143, 77, 0.16);
            color: #b7f7d1;
            border: 1px solid rgba(183, 247, 209, 0.35);
        }

        .status-warn {
            background: rgba(183, 121, 31, 0.18);
            color: #ffe0a3;
            border: 1px solid rgba(255, 224, 163, 0.35);
        }

        .status-fail {
            background: rgba(197, 48, 48, 0.18);
            color: #ffc0c0;
            border: 1px solid rgba(255, 192, 192, 0.35);
        }

        .grid {
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 14px;
            margin-bottom: 20px;
        }

        .card {
            background: var(--panel);
            border: 1px solid var(--border);
            border-radius: 14px;
            padding: 18px;
            box-shadow: 0 2px 8px rgba(15, 23, 42, 0.05);
        }

        .card h2 {
            margin: 0 0 12px 0;
            font-size: 17px;
        }

        .meta-label {
            color: var(--muted);
            font-size: 12px;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            margin-bottom: 4px;
        }

        .meta-value {
            font-weight: 650;
            word-break: break-word;
        }

        /* v0.20 Endpoint Summary / ID badge.
           Why this exists:
           - This compact top panel makes the first screen feel more like a
             professional endpoint triage dashboard.
           - It replaces the older spread-out metadata card grid with one
             scannable asset-style summary.

           Important:
           - These styles affect the HTML report only.
           - They do not affect health-check logic, console output, or TXT output.
           - Keep this scoped to endpoint-* classes so later report sections are
             not accidentally redesigned during this release. */
        .endpoint-summary {
            display: grid;
            gap: 14px;
            margin-bottom: 20px;
            border-radius: 10px;
            border: 1px solid rgba(148, 163, 184, 0.22);
            box-shadow: 0 2px 8px rgba(15, 23, 42, 0.06);
        }

        .endpoint-identity {
            min-width: 0;
        }

        .endpoint-main {
            display: grid;
            grid-template-columns: minmax(0, 1.7fr) minmax(320px, 0.9fr);
            gap: 18px;
            align-items: stretch;
        }

        .endpoint-kicker {
            color: var(--muted);
            font-size: 12px;
            font-weight: 700;
            letter-spacing: 0.08em;
            margin-bottom: 6px;
            text-transform: uppercase;
        }

        .endpoint-title-row {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 12px;
            flex-wrap: wrap;
            margin-bottom: 6px;
        }

        .endpoint-title-row h2 {
            margin: 0;
            font-size: 24px;
            line-height: 1.2;
            word-break: break-word;
        }

        .endpoint-bottom-row {
            display: grid;
            grid-template-columns: minmax(0, 1.7fr) minmax(320px, 0.9fr);
            gap: 18px;
        }

        .endpoint-description {
            margin: 0 0 16px 0;
            color: var(--muted);
            font-size: 14px;
        }

        .endpoint-status-pill {
            border-radius: 999px;
            display: inline-block;
            font-size: 12px;
            font-weight: 800;
            letter-spacing: 0.04em;
            padding: 7px 12px;
            text-transform: uppercase;
        }

        .endpoint-status-pill.status-ok {
            background: rgba(31, 143, 77, 0.12);
            border: 1px solid rgba(31, 143, 77, 0.35);
            color: var(--ok);
        }

        .endpoint-status-pill.status-warn {
            background: rgba(183, 121, 31, 0.12);
            border: 1px solid rgba(183, 121, 31, 0.35);
            color: var(--warn);
        }

        .endpoint-status-pill.status-fail {
            background: rgba(197, 48, 48, 0.12);
            border: 1px solid rgba(197, 48, 48, 0.35);
            color: var(--fail);
        }

        .endpoint-meta-stack {
            display: grid;
            gap: 10px;
        }

        .endpoint-meta-item {
            background: #f8fafc;
            border: 1px solid rgba(148, 163, 184, 0.22);
            border-radius: 10px;
            padding: 10px 12px;
            min-width: 0;
        }

        .endpoint-meta-label {
            color: var(--muted);
            display: block;
            font-size: 11px;
            font-weight: 700;
            letter-spacing: 0.06em;
            margin-bottom: 4px;
            text-transform: uppercase;
        }

        .endpoint-meta-item strong {
            display: block;
            font-size: 14px;
            overflow-wrap: break-word;
        }

        .endpoint-status-panel {
            background: #f8fafc;
            border: 1px solid rgba(148, 163, 184, 0.22);
            border-radius: 10px;
            padding: 14px;
            display: grid;
            align-content: center;
            gap: 14px;
        }

        .endpoint-status-heading {
            color: var(--muted);
            font-size: 12px;
            font-weight: 700;
            letter-spacing: 0.08em;
            text-transform: uppercase;
        }

        .endpoint-count-strip {
            display: grid;
            grid-template-columns: repeat(4, minmax(0, 1fr));
            gap: 10px;
        }

        .endpoint-count {
            background: var(--panel);
            border: 1px solid var(--border);
            border-radius: 10px;
            padding: 12px 8px;
            text-align: center;
        }

        .endpoint-count-number {
            display: block;
            font-size: 24px;
            font-weight: 800;
            line-height: 1;
            margin-bottom: 5px;
        }

        .endpoint-count-label {
            color: var(--muted);
            display: block;
            font-size: 11px;
            font-weight: 800;
            letter-spacing: 0.08em;
        }

        .endpoint-count-ok .endpoint-count-number {
            color: var(--ok);
        }

        .endpoint-count-warn .endpoint-count-number {
            color: var(--warn);
        }

        .endpoint-count-fail .endpoint-count-number {
            color: var(--fail);
        }

        .endpoint-count-total .endpoint-count-number {
            color: var(--text);
        }

        .summary-counts {
            display: flex;
            gap: 10px;
            flex-wrap: wrap;
        }

        .count-pill {
            background: var(--chip);
            border: 1px solid var(--border);
            border-radius: 999px;
            padding: 7px 12px;
            font-weight: 650;
            font-size: 13px;
        }

        .count-ok {
            color: var(--ok);
        }

        .count-warn {
            color: var(--warn);
        }

        .count-fail {
            color: var(--fail);
        }

        .section {
            margin-bottom: 20px;
        }

        .section-title {
            margin-bottom: 16px;
        }

        .section-title h2 {
            margin-bottom: 4px;
        }

        .section-title p {
            margin: 0;
            color: var(--muted);
            font-size: 0.95rem;
        }

        .readiness-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
            gap: 14px;
        }

        .readiness-card {
            background: #f8fafc;
            border: 1px solid var(--border);
            border-radius: 14px;
            padding: 16px;
        }

        .readiness-card-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 12px;
            margin-bottom: 14px;
        }

        .readiness-card h3 {
            margin: 0;
            font-size: 16px;
        }

        .readiness-status {
            border-radius: 999px;
            padding: 4px 10px;
            font-size: 11px;
            font-weight: 700;
            letter-spacing: 0.04em;
            text-transform: uppercase;
        }

        .readiness-counts {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 8px;
            margin-bottom: 12px;
        }

        .readiness-counts span {
            background: var(--panel);
            border: 1px solid var(--border);
            border-radius: 10px;
            color: var(--muted);
            font-size: 12px;
            padding: 8px;
            text-align: center;
        }

        .readiness-counts strong {
            display: block;
            color: var(--text);
            font-size: 17px;
        }

        .readiness-card p {
            margin: 0;
            color: var(--muted);
            font-size: 13px;
        }

        .readiness-ok {
            border-color: rgba(31, 143, 77, 0.45);
        }

        .readiness-ok .readiness-status {
            background: rgba(31, 143, 77, 0.12);
            color: var(--ok);
        }

        .readiness-attention {
            border-color: rgba(183, 121, 31, 0.45);
        }

        .readiness-attention .readiness-status {
            background: rgba(183, 121, 31, 0.12);
            color: var(--warn);
        }

        .readiness-critical {
            border-color: rgba(197, 48, 48, 0.45);
        }

        .readiness-critical .readiness-status {
            background: rgba(197, 48, 48, 0.12);
            color: var(--fail);
        }

        .readiness-neutral {
            border-color: rgba(101, 117, 139, 0.35);
        }

        .readiness-neutral .readiness-status {
            background: rgba(101, 117, 139, 0.12);
            color: var(--muted);
        }

        /* v0.17 Recommended Actions queue styles.
           These classes only affect the HTML report. They do not affect console
           output, TXT reports, or the health-check logic. */

        /* Overall container for all action groups. Grid gives us even spacing
           between groups without needing JavaScript or external CSS. */
        .action-queue {
            display: grid;
            gap: 16px;
        }

        /* One group box, such as System Stability or Network Connectivity. */
        .action-group {
            border: 1px solid var(--border);
            border-radius: 14px;
            background: #f8fafc;
            padding: 16px;
        }

        /* Category heading inside each action group. */
        .action-group h3 {
            margin: 0 0 12px 0;
            font-size: 15px;
        }

        /* One individual ticket-style action item. The left border is neutral by
           default and becomes yellow/red when action-warn or action-fail is added. */
        .action-item {
            background: var(--panel);
            border: 1px solid var(--border);
            border-left: 5px solid var(--muted);
            border-radius: 12px;
            padding: 13px 14px;
            margin-top: 10px;
        }

        /* WARN action items get the warning color on the left border. */
        .action-warn {
            border-left-color: var(--warn);
        }

        /* FAIL action items get the failure color on the left border. */
        .action-fail {
            border-left-color: var(--fail);
        }

        /* Header row inside an action item. Flex keeps the severity badge and
           title aligned while still allowing wrapping on smaller screens. */
        .action-header {
            display: flex;
            align-items: center;
            gap: 10px;
            flex-wrap: wrap;
        }

        /* Small WARN/FAIL pill shown beside each action title. */
        .action-severity {
            background: var(--chip);
            border: 1px solid var(--border);
            border-radius: 999px;
            font-size: 11px;
            font-weight: 700;
            letter-spacing: 0.04em;
            padding: 4px 9px;
        }

        /* Make the WARN badge text use the same warning color used elsewhere. */
        .action-warn .action-severity {
            color: var(--warn);
        }

        /* Make the FAIL badge text use the same failure color used elsewhere. */
        .action-fail .action-severity {
            color: var(--fail);
        }

        /* Optional detail text and suggested-action text under each action item. */
        .action-detail,
        .action-suggestion {
            margin: 9px 0 0 0;
            color: var(--muted);
            font-size: 13px;
        }

        ul {
            margin: 0;
            padding-left: 22px;
        }

        li {
            margin: 8px 0;
        }

        code {
            background: #f1f5f9;
            border: 1px solid #e2e8f0;
            border-radius: 6px;
            padding: 2px 5px;
            font-family: Consolas, "Courier New", monospace;
            font-size: 13px;
        }

        pre {
            white-space: pre-wrap;
            word-break: break-word;
            background: #0f172a;
            color: #e5e7eb;
            border-radius: 12px;
            padding: 18px;
            overflow-x: auto;
            font-family: Consolas, "Courier New", monospace;
            font-size: 13px;
        }

        /* v0.21 collapsible technical depth
        Why this exists:
        - Raw Findings are useful for troubleshooting, but visually heavy.
        - Native <details>/<summary> gives us a clean collapse/expand control.
        - This uses no JavaScript and does not affect check logic, console output, or TXT output.

        Troubleshooting rule:
        - Treat <summary> like the clickable button.
        - Treat summary::after like the chevron icon.
        */
        .technical-depth {
            margin: 0;
        }

        .technical-depth summary {
            min-height: 48px;
            padding: 0 24px;
            cursor: pointer;
            list-style: none;
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 16px;
            font-size: 1.05rem;
            font-weight: 700;
            color: #0f172a;
            user-select: none;
        }

        .technical-depth summary::-webkit-details-marker {
            display: none;
        }

        .technical-depth summary::marker {
            content: "";
        }

        /* Custom chevron.
        This avoids the "door off the hinges" look because the icon is a small CSS shape
        rotating around its own center instead of rotating a text glyph.
        */
        .technical-depth summary::after {
            content: "";
            width: 9px;
            height: 9px;
            border-right: 2px solid #64748b;
            border-bottom: 2px solid #64748b;
            transform: rotate(-45deg);
            transform-origin: center center;
            transition: transform 0.2s ease;
            flex: 0 0 auto;
        }

        /* When the native details block is open, rotate the chevron downward. */
        .technical-depth[open] summary::after {
            transform: rotate(45deg);
        }

        .technical-depth[open] summary {
            border-bottom: 1px solid #e2e8f0;
        }

        .technical-depth pre {
            margin: 20px;
        }

        .muted {
            color: var(--muted);
        }

        @media (max-width: 900px) {
            .report-shell {
                grid-template-columns: 1fr;
            }

            .report-sidebar {
                position: static;
            }

            .endpoint-main,
            .endpoint-bottom-row {
                grid-template-columns: 1fr;
            }

            .grid {
                grid-template-columns: repeat(2, minmax(0, 1fr));
            }
        }

        @media (max-width: 600px) {
            body {
                padding: 16px;
            }

            .grid,
            .endpoint-count-strip {
                grid-template-columns: 1fr;
            }
        }
    </style>
</head>
<body>
    <div class="report-shell">
$ReportNavigationHtml

        <main class="report-main">
        <!-- ============================================================
             v0.18 REPORT SUMMARY MODULE
             ============================================================
             This named section groups the top summary area of the HTML report.

             Sidebar link:
             - href="#report-summary"

             Matching anchor:
             - id="report-summary"

             Why this matters:
             - The sidebar can now jump back to the top summary area without using
               a generic "Back to Top" label.
             - This keeps the HTML report modular: Report Summary, Incident Summary,
               Readiness Overview, detailed sections, Recommended Actions, and Raw Findings.

             Contains:
             - report title/status header
             - v0.20 Endpoint Summary badge
             - report ID
             - generated time
             - computer name
             - current user
             - Battlestation Profile
             - overall status
             - OK/WARN/FAIL/total check counts
             ============================================================ -->
        <section id="report-summary" class="report-summary">
        <header class="ticket-header">
            <div class="eyebrow">ArcForge First Response</div>
            <h1>First Response Report</h1>
            <div class="subtitle">Static triage record generated from local workstation readiness checks.</div>
        </header>

$EndpointSummaryHtml
        </section>
        <!-- v0.18 REPORT SUMMARY MODULE END -->

        <section id="incident-summary" class="card section">
            <h2>Incident Summary</h2>
            <p class="muted">
                ArcForge First Response completed a local workstation readiness check using the
                <strong>$(ConvertTo-HtmlSafeText $BattlestationProfile)</strong> Battlestation Profile.
            </p>
            <div class="summary-counts">
                <span class="count-pill count-ok">OK: $OkCount</span>
                <span class="count-pill count-warn">WARN: $WarnCount</span>
                <span class="count-pill count-fail">FAIL: $FailCount</span>
            </div>
        </section>

$ReadinessOverviewHtml

        <section id="system" class="card section">
            <h2>System</h2>
            <ul>
                $SystemFindingsHtml
            </ul>
        </section>

        <section id="network" class="card section">
            <h2>Network</h2>
            <ul>
                $NetworkFindingsHtml
            </ul>
        </section>

        <section id="software-readiness" class="card section">
            <h2>Software Readiness</h2>
            <ul>
                $SoftwareFindingsHtml
            </ul>
        </section>

        <section id="security" class="card section">
            <h2>Security</h2>
            <ul>
                $SecurityFindingsHtml
            </ul>
        </section>

        <section id="updates" class="card section">
            <h2>Updates</h2>
            <ul>
                $UpdatesFindingsHtml
            </ul>
        </section>

$RecommendedActionsHtml

        <!--
            v0.21 collapsible technical depth
            Why this exists:
            - Raw Findings are still part of the HTML report, but they are no
              longer forced into the main scan path.
            - <details>/<summary> is native HTML, so this adds collapsible
              behavior without JavaScript or external dependencies.

            Important:
            - This is presentation-only.
            - This does not change check logic, console output, or TXT output.
            - The sidebar link still points to this same raw-findings section.
        -->
        <section id="raw-findings" class="card section">
            <details class="technical-depth">
                <summary>Raw Findings</summary>
                <pre>$RawFindings</pre>
            </details>
        </section>
        </main>
    </div>
</body>
</html>
"@

    $Html | Out-File -FilePath $OutputPath -Encoding UTF8
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
                        -Commands $DetectionConfig.Commands `
                        -DisplayNamePatterns $DetectionConfig.DisplayNamePatterns `
                        -CommonPaths $DetectionConfig.CommonPaths `
                        -Services $DetectionConfig.Services

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
Write-Host "TXT report saved to: $ReportFile" -ForegroundColor Gray
Write-Host "HTML report saved to: $HtmlReportFile" -ForegroundColor Gray

Add-ReportLine
Add-ReportLine -Line "Health check complete."
Add-ReportLine -Line "TXT report saved to: $ReportFile"
Add-ReportLine -Line "HTML report saved to: $HtmlReportFile"

$ReportLines | Out-File -FilePath $ReportFile -Encoding UTF8

New-ArcForgeHtmlReport `
    -OutputPath $HtmlReportFile `
    -ReportId $ReportId `
    -ComputerName $ComputerName `
    -CurrentUser $CurrentUser `
    -BattlestationProfile $BattlestationProfile `
    -GeneratedAt $ReportDate `
    -CheckCounts $CheckCounts `
    -ReportLines $ReportLines