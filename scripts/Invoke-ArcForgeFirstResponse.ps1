# ArcForge First Response
# ArcForge First Response Report v0.16

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

    function ConvertTo-HtmlSafeText {
        param (
            [string]$Text
        )

        return [System.Net.WebUtility]::HtmlEncode($Text)
    }

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
        <section class="card section">
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

    $ReadinessOverviewHtml = New-ArcForgeReadinessOverviewHtml -ReadinessCards @(
        Get-ArcForgeSectionReadiness -Name "System" -Lines $SystemLines
        Get-ArcForgeSectionReadiness -Name "Network" -Lines $NetworkLines
        Get-ArcForgeSectionReadiness -Name "Software" -Lines $SoftwareLines
        Get-ArcForgeSectionReadiness -Name "Security" -Lines $SecurityLines
        Get-ArcForgeSectionReadiness -Name "Updates" -Lines $UpdatesLines
    )

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

    $RecommendedActions = @(
        $ReportLines |
            Where-Object {
                $_ -match "^\[(WARN|FAIL)\]" -and
                $_ -notmatch "^\[(WARN|FAIL)\]\s+Warnings:" -and
                $_ -notmatch "^\[(WARN|FAIL)\]\s+Failures:" -and
                $_ -notmatch "^\[(WARN|FAIL)\]\s+Overall Status:"
            } |
            ForEach-Object { ConvertTo-HtmlSafeText $_ }
    )

    if (-not $RecommendedActions -or $RecommendedActions.Count -eq 0) {
        $RecommendedActionsHtml = "<li>No immediate recommended actions. System appears healthy based on current checks.</li>"
    }
    else {
        $RecommendedActionsHtml = ($RecommendedActions | ForEach-Object {
            "<li><code>$_</code></li>"
        }) -join "`n"
    }

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
            font-family: "Segoe UI", Arial, sans-serif;
            line-height: 1.5;
        }

        .report-shell {
            max-width: 1100px;
            margin: 0 auto;
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

        .muted {
            color: var(--muted);
        }

        @media (max-width: 900px) {
            .grid {
                grid-template-columns: repeat(2, minmax(0, 1fr));
            }
        }

        @media (max-width: 600px) {
            body {
                padding: 16px;
            }

            .grid {
                grid-template-columns: 1fr;
            }
        }
    </style>
</head>
<body>
    <main class="report-shell">
        <header class="ticket-header">
            <div class="eyebrow">ArcForge First Response</div>
            <h1>First Response Report</h1>
            <div class="subtitle">Static triage record generated from local workstation readiness checks.</div>
            <div class="status-row">
                <span class="status-badge $StatusClass">$OverallStatus</span>
            </div>
        </header>

        <section class="grid">
            <div class="card">
                <div class="meta-label">Report ID</div>
                <div class="meta-value">$(ConvertTo-HtmlSafeText $ReportId)</div>
            </div>
            <div class="card">
                <div class="meta-label">Generated At</div>
                <div class="meta-value">$(ConvertTo-HtmlSafeText $GeneratedAt)</div>
            </div>
            <div class="card">
                <div class="meta-label">Computer</div>
                <div class="meta-value">$(ConvertTo-HtmlSafeText $ComputerName)</div>
            </div>
            <div class="card">
                <div class="meta-label">Current User</div>
                <div class="meta-value">$(ConvertTo-HtmlSafeText $CurrentUser)</div>
            </div>
            <div class="card">
                <div class="meta-label">Battlestation Profile</div>
                <div class="meta-value">$(ConvertTo-HtmlSafeText $BattlestationProfile)</div>
            </div>
            <div class="card">
                <div class="meta-label">Overall Status</div>
                <div class="meta-value">$OverallStatus</div>
            </div>
            <div class="card">
                <div class="meta-label">Total Checks</div>
                <div class="meta-value">$TotalChecks</div>
            </div>
            <div class="card">
                <div class="meta-label">Report Type</div>
                <div class="meta-value">First Response Triage</div>
            </div>
        </section>

        <section class="card section">
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

        <section class="card section">
            <h2>System</h2>
            <ul>
                $SystemFindingsHtml
            </ul>
        </section>

        <section class="card section">
            <h2>Network</h2>
            <ul>
                $NetworkFindingsHtml
            </ul>
        </section>

        <section class="card section">
            <h2>Software Readiness</h2>
            <ul>
                $SoftwareFindingsHtml
            </ul>
        </section>

        <section class="card section">
            <h2>Security</h2>
            <ul>
                $SecurityFindingsHtml
            </ul>
        </section>

        <section class="card section">
            <h2>Updates</h2>
            <ul>
                $UpdatesFindingsHtml
            </ul>
        </section>

        <section class="card section">
            <h2>Recommended Actions</h2>
            <ul>
                $RecommendedActionsHtml
            </ul>
        </section>

        <section class="card section">
            <h2>Raw Findings</h2>
            <pre>$RawFindings</pre>
        </section>
    </main>
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