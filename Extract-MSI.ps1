<#
.SYNOPSIS
Extracts an MSI package to a target folder using msiexec administrative install.

.PARAMETER MsiPath
Path to the .msi file to extract.

.PARAMETER Destination
Target folder where extracted files are written.

.EXAMPLE
.\Extract-MSI.ps1 -MsiPath "C:\Temp\package.msi" -Destination "C:\Temp\package-extracted"
#>

param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$MsiPath,

    [Parameter(Mandatory = $true, Position = 1)]
    [ValidateNotNullOrEmpty()]
    [string]$Destination
)

function Test-IsAdministrator {
    $currentIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($currentIdentity)
    return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-IsAdministrator)) {
    Write-Warning "This script should be run as Administrator for MSI extraction to work correctly."
}

$msiFile = Resolve-Path -Path $MsiPath -ErrorAction Stop
if (-not (Test-Path -Path $msiFile)) {
    Throw "MSI file not found: $MsiPath"
}

$destinationPath = Resolve-Path -Path $Destination -ErrorAction SilentlyContinue
if (-not $destinationPath) {
    New-Item -ItemType Directory -Path $Destination -Force | Out-Null
    $destinationPath = Resolve-Path -Path $Destination
}

$targetDir = $destinationPath.ProviderPath
$msiFilePath = $msiFile.ProviderPath
$arguments = "/a `"$msiFilePath`" /qn TARGETDIR=`"$targetDir`""

Write-Host "Extracting MSI to: $targetDir" -ForegroundColor Cyan
Write-Host "msiexec.exe $arguments" -ForegroundColor DarkGray

$process = Start-Process -FilePath "msiexec.exe" -ArgumentList $arguments -Wait -NoNewWindow -PassThru

if ($process.ExitCode -eq 0) {
    Write-Host "Extraction complete: $targetDir" -ForegroundColor Green
} else {
    Throw "MSI extraction failed with exit code $($process.ExitCode)."
}
