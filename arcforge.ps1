<#
.SYNOPSIS
Launches ArcForge First Response.

.DESCRIPTION
Root launcher for ArcForge First Response. This allows users to run the toolkit
from the repository root without calling the full script path.
#>

$ScriptPath = Join-Path $PSScriptRoot 'scripts\Invoke-ArcForgeFirstResponse.ps1'

if (-not (Test-Path $ScriptPath)) {
    Write-Error "Unable to find Invoke-ArcForgeFirstResponse.ps1 in the scripts folder."
    exit 1
}

& $ScriptPath @args