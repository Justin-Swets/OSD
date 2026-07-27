[CmdletBinding()]
param(
    [string]$SourceDrive = 'E:',
    [string]$DestinationPath = 'E:\KB5121767',
    [string]$TargetRoot = 'C:\',
    [string]$CatalogUrl = 'https://catalog.update.microsoft.com/Search.aspx?q=KB5121767'
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
        Where-Object { $_.Name -match 'KB5121767|5121767' -and $_.Extension -in '.msu','.cab','.exe' }

    if ($candidates) {
        return $candidates | Sort-Object Length -Descending | Select-Object -First 1 -ExpandProperty FullName
    }

    return $null
}

function Download-KB5121767 {
    param(
        [string]$DestinationDir,
        [string]$CatalogUrl
    )

    Ensure-Directory -Path $DestinationDir

    $packagePath = Join-Path $DestinationDir 'windows11.0-kb5121767-x64.msu'
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

    Write-Host "Installing update package from $PackagePath"
    $logPath = Join-Path $TargetRoot 'Windows\Logs\KB5121767-install.log'
    $wusaArgs = @('/quiet', '/norestart', '/extract:' + $TargetRoot + 'Temp\KB5121767', $PackagePath)
    Start-Process -FilePath 'wusa.exe' -ArgumentList $wusaArgs -Wait -PassThru -NoNewWindow | Out-Null
}

$packageOnDrive = Find-KBPackage -RootPath $SourceDrive
if ($packageOnDrive) {
    Write-Host "Found package on source drive: $packageOnDrive"
    $packagePath = $packageOnDrive
}
else {
    $packagePath = Download-KB5121767 -DestinationDir $DestinationPath -CatalogUrl $CatalogUrl
    Write-Host "Downloaded package to $packagePath"
}

Install-KBUpdate -PackagePath $packagePath -TargetRoot $TargetRoot

Write-Host 'KB5121767 package processing completed.'
