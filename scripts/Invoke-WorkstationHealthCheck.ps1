# ArcForge Studio IT Toolkit
# Workstation Health Check v0.1

$ReportDate = Get-Date
$ComputerName = $env:COMPUTERNAME
$CurrentUser = $env:USERNAME

Write-Host "========================================"
Write-Host " ArcForge Studio Workstation Health Check"
Write-Host "========================================"
Write-Host ""
Write-Host "Computer Name: $ComputerName"
Write-Host "Current User:  $CurrentUser"
Write-Host "Report Date:   $ReportDate"
Write-Host ""

# OS Information
$OS = Get-CimInstance Win32_OperatingSystem

Write-Host "[SYSTEM]"
Write-Host "OS Name:        $($OS.Caption)"
Write-Host "OS Version:     $($OS.Version)"
Write-Host "Architecture:   $($OS.OSArchitecture)"
Write-Host ""

# Uptime
$LastBoot = $OS.LastBootUpTime
$Uptime = (Get-Date) - $LastBoot

Write-Host "[UPTIME]"
Write-Host "Last Boot:      $LastBoot"
Write-Host "Uptime Days:    $([math]::Round($Uptime.TotalDays, 2))"
Write-Host ""

# Disk Check
$Disk = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
$FreeGB = [math]::Round($Disk.FreeSpace / 1GB, 2)
$TotalGB = [math]::Round($Disk.Size / 1GB, 2)
$FreePercent = [math]::Round(($Disk.FreeSpace / $Disk.Size) * 100, 2)

Write-Host "[STORAGE]"
Write-Host "Drive:          C:"
Write-Host "Free Space:     $FreeGB GB"
Write-Host "Total Size:     $TotalGB GB"
Write-Host "Free Percent:   $FreePercent%"
Write-Host ""

# Network Checks
Write-Host "[NETWORK]"

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

        Write-Host "IPv4 Address:    $IPAddress"
        Write-Host "Default Gateway: $Gateway"
        Write-Host "DNS Servers:     $DNSServers"

        if (Test-Connection -ComputerName $Gateway -Count 2 -Quiet) {
            Write-Host "Gateway Ping:    OK"
        }
        else {
            Write-Host "Gateway Ping:    FAIL"
        }
    }
    else {
        Write-Host "IPv4 Address:    Not found"
        Write-Host "Default Gateway: Not found"
        Write-Host "DNS Servers:     Not found"
        Write-Host "Gateway Ping:    SKIPPED"
    }
}
catch {
    Write-Host "Network Config:  ERROR - $($_.Exception.Message)"
    Write-Host "Gateway Ping:    SKIPPED"
}

if (Test-Connection -ComputerName "1.1.1.1" -Count 2 -Quiet) {
    Write-Host "Internet Ping:   OK"
}
else {
    Write-Host "Internet Ping:   FAIL"
}

try {
    Resolve-DnsName "github.com" -ErrorAction Stop | Out-Null
    Write-Host "DNS Resolution:  OK"
}
catch {
    Write-Host "DNS Resolution:  FAIL"
}

Write-Host ""
Write-Host "Health check complete."