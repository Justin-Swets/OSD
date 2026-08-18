[CmdletBinding()]
param(
    [string]$SourceDrive = 'E:',
    [string]$DestinationPath = 'E:\Latest-CU-25H2',
    [string]$TargetRoot = 'C:\',
    [string]$SearchQuery = 'Cumulative Update for Windows 11, version 25H2 x64'
)

$ErrorActionPreference = 'Stop'

function Ensure-Directory {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Find-PackagesOnDrive {
    param([string]$RootPath)

    if (-not (Test-Path -LiteralPath $RootPath)) { return $null }

    $candidates = Get-ChildItem -Path $RootPath -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object {
            ($_.Extension -in '.msu','.cab') -and (
                $_.Name -match 'KB\d{6,7}' -or
                $_.Name -match 'Cumulative' -or
                $_.Name -match 'Windows11|Win11|Windows_11'
            )
        }

    if ($candidates) {
        return $candidates | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName
    }

    return $null
}

function Get-UpdateIdFromSearchPage {
    param([string]$Query)

    $searchUrl = "https://www.catalog.update.microsoft.com/Search.aspx?q=" + [uri]::EscapeDataString($Query)
    Write-Host "Querying Microsoft Update Catalog: $searchUrl"

    try {
        $resp = Invoke-WebRequest -Uri $searchUrl -UseBasicParsing -Headers @{ 'User-Agent' = 'Mozilla/5.0' }
    }
    catch {
        throw "Failed to query Microsoft Update Catalog: $($_.Exception.Message)"
    }

    $matches = [regex]::Matches($resp.Content, 'updateid=([0-9A-Fa-f\-]{36})') | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique
    if (-not $matches -or $matches.Count -eq 0) { return $null }

    return $matches[0]
}

function Get-DownloadUrlFromUpdateId {
    param([string]$UpdateId)

    $viewUrl = "https://www.catalog.update.microsoft.com/ScopedViewInline.aspx?updateid=$UpdateId"
    Write-Host "Fetching update details: $viewUrl"

    try {
        $resp = Invoke-WebRequest -Uri $viewUrl -UseBasicParsing -Headers @{ 'User-Agent' = 'Mozilla/5.0' }
    }
    catch {
        throw "Failed to fetch update detail page: $($_.Exception.Message)"
    }

    # Look for direct download links hosted on download.windowsupdate.com ending in .msu or .cab
    $msuMatch = [regex]::Match($resp.Content, '(https?://download\.windowsupdate\.com/[^"'']+?\.(msu|cab))', 'IgnoreCase')
    if ($msuMatch.Success) { return $msuMatch.Groups[1].Value }

    # As a fallback, try to find any href that looks like a download dialog and follow it
    $dlgMatch = [regex]::Match($resp.Content, 'href\s*=\s*"([^"]*DownloadDialog\.aspx[^\"]*)"', 'IgnoreCase')
    if ($dlgMatch.Success) {
        $dlg = $dlgMatch.Groups[1].Value
        if ($dlg -notmatch '^http') { $dlg = 'https://www.catalog.update.microsoft.com' + $dlg }
        try {
            $dlgResp = Invoke-WebRequest -Uri $dlg -UseBasicParsing -Headers @{ 'User-Agent' = 'Mozilla/5.0' }
            $msuMatch2 = [regex]::Match($dlgResp.Content, '(https?://download\.windowsupdate\.com/[^"'']+?\.(msu|cab))', 'IgnoreCase')
            if ($msuMatch2.Success) { return $msuMatch2.Groups[1].Value }
        }
        catch { }
    }

    return $null
}

function Download-UpdatePackage {
    param(
        [string]$Url,
        [string]$DestinationDir
    )

    Ensure-Directory -Path $DestinationDir

    $fileName = Split-Path -Path $Url -Leaf
    $destPath = Join-Path $DestinationDir $fileName
    if (Test-Path -LiteralPath $destPath) { return $destPath }

    Write-Host "Downloading $Url -> $destPath"
    try {
        Invoke-WebRequest -Uri $Url -OutFile $destPath -UseBasicParsing -Headers @{ 'User-Agent' = 'Mozilla/5.0' }
        return $destPath
    }
    catch {
        throw "Failed to download update package: $($_.Exception.Message)"
    }
}

function Install-CumulativeUpdate {
    param(
        [string]$PackagePath,
        [string]$TargetRoot
    )

    if (-not (Test-Path -LiteralPath $PackagePath)) { throw "Package not found: $PackagePath" }
    if (-not (Test-Path -LiteralPath (Join-Path $TargetRoot 'Windows'))) { throw "Target Windows root not found: $TargetRoot" }

    $logPath = Join-Path $TargetRoot 'Windows\Logs\LatestCU-install.log'
    Write-Host "Installing update package from $PackagePath"
    dism /Image:$TargetRoot /Add-Package /PackagePath:$PackagePath /IgnoreCheck /LogPath:$logPath
}

# Main flow
$packageOnDrive = Find-PackagesOnDrive -RootPath $SourceDrive
if ($packageOnDrive) {
    Write-Host "Found package on source drive: $packageOnDrive"
    $packagePath = $packageOnDrive
}
else {
    Write-Host "No package on $SourceDrive; attempting to locate latest CU via Microsoft Update Catalog using query: $SearchQuery"
    $updateId = Get-UpdateIdFromSearchPage -Query $SearchQuery
    if (-not $updateId) {
        $hint = @"
Could not discover an update id from the Microsoft Update Catalog for query:
$SearchQuery

Automatic discovery can fail due to catalog page changes or network restrictions.
If automatic lookup fails, manually download the latest Windows 11 25H2 x64 cumulative update from:
  https://www.catalog.update.microsoft.com/
Save the .msu file to $DestinationPath and re-run this script.
"@
        throw $hint
    }

    $fileUrl = Get-DownloadUrlFromUpdateId -UpdateId $updateId
    if (-not $fileUrl) {
        throw "Could not extract a direct download URL for updateid $updateId. Manual download required."
    }

    $packagePath = Download-UpdatePackage -Url $fileUrl -DestinationDir $DestinationPath
    Write-Host "Downloaded package to $packagePath"
}

Install-CumulativeUpdate -PackagePath $packagePath -TargetRoot $TargetRoot

Write-Host 'Latest cumulative update processing completed.'
