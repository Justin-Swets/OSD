[CmdletBinding()]
param(
    [string]$SourceDrive = 'E:',
    [string]$DestinationPath = 'E:\KB5101684',
    [string]$TargetRoot = 'C:\',
    [string]$CatalogUrl = 'https://catalog.update.microsoft.com/Search.aspx?q=KB5101684'
)

$ErrorActionPreference = 'Stop'

function Ensure-Directory {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Find-KBPackage {
    param([string]$RootPath)

    if (-not (Test-Path -LiteralPath $RootPath)) {
        return $null
    }

    $candidates = Get-ChildItem -Path $RootPath -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match 'KB5101684|5101684' -and $_.Extension -in '.msu','.cab','.exe' }

    if ($candidates) {
        return $candidates | Sort-Object Length -Descending | Select-Object -First 1 -ExpandProperty FullName
    }

    return $null
}

function Download-KB5101684 {
    param(
        [string]$DestinationDir,
        [string]$CatalogUrl
    )

    Ensure-Directory -Path $DestinationDir

    $packagePath = Join-Path $DestinationDir 'windows11.0-kb5101684-x64.msu'
    if (Test-Path -LiteralPath $packagePath) {
        return $packagePath
    }

    $catalogHint = @"
Microsoft Update Catalog search:
$CatalogUrl

Please download the Windows 11 25H2 x64 package manually from the catalog and save it as:
$packagePath
"@

    throw $catalogHint
}

function Install-KBUpdate {
    param(
        [string]$PackagePath,
        [string]$TargetRoot
    )

    if (-not (Test-Path -LiteralPath $PackagePath)) {
        throw "Package not found: $PackagePath"
    }

    if (-not (Test-Path -LiteralPath (Join-Path $TargetRoot 'Windows'))) {
        throw "Target Windows root not found: $TargetRoot"
    }

    Ensure-Directory -Path (Join-Path $TargetRoot 'Temp\KB5101684')

    $logPath = Join-Path $TargetRoot 'Windows\Logs\KB5101684-install.log'
    Write-Host "Installing update package from $PackagePath"
    dism /Image:$TargetRoot /Add-Package /PackagePath:$PackagePath /IgnoreCheck /LogPath:$logPath
}

$packageOnDrive = Find-KBPackage -RootPath $SourceDrive
if ($packageOnDrive) {
    Write-Host "Found package on source drive: $packageOnDrive"
    $packagePath = $packageOnDrive
}
else {
    $packagePath = Download-KB5101684 -DestinationDir $DestinationPath -CatalogUrl $CatalogUrl
    Write-Host "Downloaded package to $packagePath"
}

Install-KBUpdate -PackagePath $packagePath -TargetRoot $TargetRoot

Write-Host 'KB5101684 package processing completed.'
