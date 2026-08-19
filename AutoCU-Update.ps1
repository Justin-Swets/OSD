<#
.SYNOPSIS
    Finds, downloads, and applies the latest Windows 11 25H2 cumulative update.

.DESCRIPTION
    Part 1  Scrapes the Windows 11 25H2 update-history page for the newest CU,
            identified by OS Build 26200.x. Monthly and Out-of-band releases are
            eligible; Preview (C-release) entries are excluded unless -IncludePreview.
    Part 2  Resolves that KB to a Microsoft Update Catalog row for the requested
            architecture and downloads the .msu/.cab.
    Part 3  Applies the package to an OFFLINE Windows image with DISM.

.NOTES
    25H2 and 24H2 ship under the SAME KB (e.g. KB5121003 = builds 26200.9168 and
    26100.9168) but as SEPARATE catalog packages. Selection is therefore filtered on
    the catalog title's build number, not on the KB alone.

    The catalog also returns a checkpoint prerequisite (KB5043080, the 2024-09 24H2
    baseline) alongside the CU. That checkpoint is only required for images that predate
    it. Applying it to an already-serviced image fails with 0x80070228 / error 552
    ("An error occurred applying the Unattend.xml file from the .msu package"), so it is
    downloaded but NOT applied unless -IncludePrerequisites is passed - and even then a
    prerequisite failure is logged rather than aborting the run.

    Install is offline-only (dism /Image). Run from WinPE. DISM rejects /Image against
    the running OS, so this script fails fast in that case. The target image's build is
    read up front, before anything is downloaded.

.EXAMPLE
    .\AutoCU-Update.ps1 -Mode Find

.EXAMPLE
    .\AutoCU-Update.ps1 -Mode Download -Architecture x64,arm64 -DestinationPath E:\Latest-CU

.EXAMPLE
    .\AutoCU-Update.ps1 -Mode All -Architecture arm64 -TargetRoot C:\
#>
[CmdletBinding()]
param(
    [string]$SourceDrive = 'E:',
    [string]$DestinationPath = 'E:\Latest-CU',
    [string]$TargetRoot = 'C:\',
    [ValidateSet('All','Find','Download','Install')][string]$Mode = 'All',
    [ValidateSet('x64','arm64')][string[]]$Architecture = @('x64'),
    [string]$KB,
    [string]$PackagePath,
    [int]$DaysBack = 45,
    [string]$ScratchDir,
    [switch]$IncludePreview,
    [switch]$IncludePrerequisites,
    [switch]$SkipLocalSearch,
    [switch]$WhatIf,
    [switch]$OpenCatalog
)

$ErrorActionPreference = 'Stop'
# CU packages are ~5 GB; the progress bar cripples Invoke-WebRequest throughput on PS 5.1.
$ProgressPreference = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$Script:BuildFamily25H2 = '26200'
$Script:HistoryUrl      = 'https://support.microsoft.com/en-us/servicing/os/windows-11/2025/07/windows-11-version-25h2-update-history'
$Script:CatalogBase     = 'https://www.catalog.update.microsoft.com'
$Script:UserAgent       = 'Mozilla/5.0'

function Ensure-Directory {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Invoke-Page {
    param([string]$Uri, [string]$Method = 'Get', $Body = $null, [string]$Referer)

    # Accept-Language reduces (but does not eliminate) the catalog serving localized titles.
    $headers = @{ 'User-Agent' = $Script:UserAgent; 'Accept-Language' = 'en-US,en;q=0.9' }
    if ($Referer) { $headers['Referer'] = $Referer }

    $splat = @{ Uri = $Uri; UseBasicParsing = $true; Headers = $headers; TimeoutSec = 120; Method = $Method }
    if ($null -ne $Body) { $splat['Body'] = $Body }

    return Invoke-WebRequest @splat
}

function ConvertFrom-HtmlCell {
    param([string]$Html)
    return (($Html -replace '<[^>]+>', ' ') -replace '&nbsp;', ' ' -replace '\s+', ' ').Trim()
}

#region Part 1 - discover the current CU from the update-history page

function Get-CUHistoryEntry {
    <#
        The 25H2 history page embeds EVERY Windows 11 servicing branch (22000, 22621,
        22631, 26100, 26200, 28000), and the 26H1 (28000) block appears FIRST in the
        markup. Document order is therefore meaningless - entries must be filtered on
        the 26200 build family. Each entry is also emitted twice, so results are deduped.

        Heading format on the page:
            August 11, 2026-KB5121003 (OS Builds 26200.9168 and 26100.9168)
            July 28, 2026-KB5101684 (OS Builds 26200.8973 and 26100.8973) Preview
            July 18, 2026-KB5121767 (OS Builds 26200.8894 and 26100.8894) Out-of-band
    #>
    param(
        [string]$Url = $Script:HistoryUrl,
        [switch]$IncludePreview
    )

    Write-Host "Reading Windows 11 25H2 update history: $Url"
    try {
        $resp = Invoke-Page -Uri $Url
    }
    catch {
        Write-Host "Failed to fetch update-history page: $($_.Exception.Message)" -ForegroundColor Yellow
        return @()
    }

    $pattern = '(?<date>[A-Z][a-z]+ \d{1,2}, \d{4})[^<]{0,20}?KB(?<kb>\d{6,7})\s*\(OS Builds?\s*' +
               $Script:BuildFamily25H2 + '\.(?<rev>\d+)[^)]*\)(?<tag>[^<]{0,24})'

    $seen    = @{}
    $entries = New-Object System.Collections.Generic.List[object]

    foreach ($m in [regex]::Matches($resp.Content, $pattern)) {
        $kbId = 'KB' + $m.Groups['kb'].Value
        if ($seen.ContainsKey($kbId)) { continue }
        $seen[$kbId] = $true

        $tag  = $m.Groups['tag'].Value
        $type = if     ($tag -match 'Preview')     { 'Preview' }
                elseif ($tag -match 'Out-of-band') { 'OutOfBand' }
                else                               { 'Monthly' }

        $parsedDate = $null
        try {
            $parsedDate = [datetime]::Parse($m.Groups['date'].Value, [Globalization.CultureInfo]::GetCultureInfo('en-US'))
        }
        catch { }

        $entries.Add([pscustomobject]@{
            KB          = $kbId
            Date        = $parsedDate
            Build       = "$($Script:BuildFamily25H2).$($m.Groups['rev'].Value)"
            Revision    = [int]$m.Groups['rev'].Value
            ReleaseType = $type
        })
    }

    if ($entries.Count -eq 0) {
        Write-Host "No $($Script:BuildFamily25H2).x entries parsed from the update-history page. Layout may have changed." -ForegroundColor Yellow
        return @()
    }

    $eligible = if ($IncludePreview) { $entries } else { $entries | Where-Object { $_.ReleaseType -ne 'Preview' } }

    # Sort on build revision: monotonic, and independent of date parsing.
    return @($eligible | Sort-Object Revision -Descending)
}

function Get-LatestCU {
    param([int]$DaysBack = 45, [switch]$IncludePreview)

    $entries = Get-CUHistoryEntry -IncludePreview:$IncludePreview
    if ($entries.Count -eq 0) { return $null }

    Write-Host "25H2 candidates (newest first):" -ForegroundColor Cyan
    $entries | Select-Object -First 5 | ForEach-Object {
        $d = if ($_.Date) { $_.Date.ToString('yyyy-MM-dd') } else { 'unknown' }
        Write-Host ("  {0}  {1}  build {2}  [{3}]" -f $_.KB, $d, $_.Build, $_.ReleaseType)
    }

    $latest = $entries[0]

    # Staleness is a warning, not a hard filter. Returning $null on a stale page is what
    # made the previous version fall through to an unrelated package on the flash drive.
    if ($latest.Date -and $latest.Date -lt (Get-Date).AddDays(-$DaysBack)) {
        Write-Host ("WARNING: newest 25H2 CU ({0}, {1}) is older than {2} days. Proceeding anyway." -f
                    $latest.KB, $latest.Date.ToString('yyyy-MM-dd'), $DaysBack) -ForegroundColor Yellow
    }

    return $latest
}

#endregion

#region Part 2 - resolve the KB in the Update Catalog and download

function Get-CatalogUpdateRow {
    <#
        Catalog search results are an ASP.NET table. A row's updateId is the GUID in the
        cell id attribute: id="<guid>_C1_R<n>" is the title cell and id="<guid>_C6_R<n>"
        carries the size in bytes. The catalog does NOT emit "updateid=<guid>" anywhere
        in the search markup - that pattern matches zero times.
    #>
    param(
        [Parameter(Mandatory=$true)][string]$KB,
        [string]$Version = '25H2',
        [Parameter(Mandatory=$true)][ValidateSet('x64','arm64')][string]$Arch
    )

    $url      = "$($Script:CatalogBase)/Search.aspx?q=" + [uri]::EscapeDataString($KB)
    $attempts = 3

    for ($try = 1; $try -le $attempts; $try++) {
        Write-Host "Querying Update Catalog: $url"

        try { $resp = Invoke-Page -Uri $url }
        catch {
            Write-Host "Catalog query failed (attempt $try/$attempts): $($_.Exception.Message)" -ForegroundColor Yellow
            if ($try -lt $attempts) { Start-Sleep -Seconds 3; continue }
            return $null
        }

        $content = $resp.Content
        $rows    = [regex]::Matches($content, '(?s)id="(?<guid>[0-9A-Fa-f\-]{36})_C1_R\d+"[^>]*>(?<title>.*?)</td>')

        if ($rows.Count -eq 0) {
            Write-Host "No catalog result rows found for $KB (attempt $try/$attempts)." -ForegroundColor Yellow
            if ($try -lt $attempts) { Start-Sleep -Seconds 3; continue }
            return $null
        }

        $found = New-Object System.Collections.Generic.List[object]

        foreach ($row in $rows) {
            $guid  = $row.Groups['guid'].Value
            $title = ConvertFrom-HtmlCell -Html $row.Groups['title'].Value

            # These filters are deliberately locale-invariant. The catalog intermittently
            # serves localized titles - e.g. the Russian form renders "x64-based Systems"
            # as "процессоров x64" - so matching English prose is unreliable. The KB, the
            # build number and the bare architecture token survive translation, and the
            # build number is what actually separates 25H2 (26200.x) from 24H2 (26100.x),
            # which share the same KB.
            $buildInTitle = [regex]::Match($title, '\((?<b>' + $Script:BuildFamily25H2 + '\.\d+)\)')

            if ($title -notmatch [regex]::Escape($KB))                            { continue }
            if (-not $buildInTitle.Success)                                       { continue }
            if ($title -notmatch "(?<![a-z0-9])$Arch(?![a-z0-9])")                { continue }
            if ($title -match 'Dynamic Update|\.NET Framework|server operating')  { continue }

            $sizeCell = [regex]::Match($content, 'id="' + [regex]::Escape($guid) + '_C6_R\d+"[^>]*>(?<s>.*?)</td>', 'Singleline')
            $bytes    = $null
            if ($sizeCell.Success) {
                $b = [regex]::Match((ConvertFrom-HtmlCell -Html $sizeCell.Groups['s'].Value), '(\d{6,})')
                if ($b.Success) { $bytes = [int64]$b.Groups[1].Value }
            }

            $found.Add([pscustomobject]@{
                UpdateId = $guid
                Title    = $title
                Bytes    = $bytes
                Build    = $buildInTitle.Groups['b'].Value
            })
        }

        if ($found.Count -gt 0) {
            if ($found.Count -gt 1) {
                Write-Host "Multiple matching rows; using the first:" -ForegroundColor Yellow
                $found | ForEach-Object { Write-Host "  $($_.Title)" }
            }
            return $found[0]
        }

        Write-Host "$KB is in the catalog, but no row matched build $($Script:BuildFamily25H2).x + '$Arch' (attempt $try/$attempts)." -ForegroundColor Yellow
        if ($try -lt $attempts) { Start-Sleep -Seconds 3 }
    }

    return $null
}

function Get-CatalogDownloadFile {
    <#
        The download URLs are only obtainable by POSTing an updateIDs payload to
        DownloadDialog.aspx. ScopedViewInline.aspx contains no .msu/.cab link and no
        DownloadDialog reference at all. Files are served from *.delivery.mp.microsoft.com,
        NOT from download.windowsupdate.com.

        IMPORTANT: the dialog returns EVERY file the update requires, not just the CU.
        For Windows 11 24H2/25H2 that includes the checkpoint baseline (KB5043080)
        alongside the target CU, e.g. for KB5121003 x64:

            windows11.0-kb5121003-x64_....msu   4867 MB  (target CU)
            windows11.0-kb5043080-x64_....msu    509 MB  (checkpoint prerequisite)

        The order of those entries VARIES between requests, so files must be matched by
        KB number - never by position. Returns them in DISM apply order: prerequisites
        first, target CU last.
    #>
    param(
        [Parameter(Mandatory=$true)][string]$UpdateId,
        [Parameter(Mandatory=$true)][string]$KB
    )

    $payload = @{ size = 0; languages = ''; uidInfo = $UpdateId; updateID = $UpdateId }
    $body    = @{ updateIDs = '[' + ($payload | ConvertTo-Json -Compress) + ']' }

    try {
        $resp = Invoke-Page -Uri "$($Script:CatalogBase)/DownloadDialog.aspx" -Method Post -Body $body -Referer "$($Script:CatalogBase)/Search.aspx"
    }
    catch {
        Write-Host "DownloadDialog POST failed: $($_.Exception.Message)" -ForegroundColor Yellow
        return @()
    }

    $urls = [regex]::Matches($resp.Content, "downloadInformation\[\d+\]\.files\[\d+\]\.url\s*=\s*'([^']+)'") |
                ForEach-Object { $_.Groups[1].Value } |
                Where-Object   { $_ -match '\.(msu|cab)$' } |
                Select-Object -Unique

    if (-not $urls -or @($urls).Count -eq 0) {
        Write-Host "No .msu/.cab URL in the DownloadDialog response for $UpdateId." -ForegroundColor Yellow
        return @()
    }

    $kbDigits = $KB -replace '^KB', ''

    $files = foreach ($u in @($urls)) {
        $name  = Split-Path -Path ($u -split '\?')[0] -Leaf
        $fileKB = if ($name -match 'kb(\d{6,7})') { 'KB' + $Matches[1] } else { $null }
        [pscustomobject]@{
            Url      = $u
            FileName = $name
            KB       = $fileKB
            IsTarget = ($fileKB -eq "KB$kbDigits")
        }
    }

    if (-not (@($files) | Where-Object { $_.IsTarget })) {
        Write-Host "DownloadDialog returned files, but none matched ${KB}:" -ForegroundColor Yellow
        @($files) | ForEach-Object { Write-Host "  $($_.FileName)" }
        return @()
    }

    # Checkpoint/prerequisite packages must be applied before the target CU.
    $prereqs = @(@($files) | Where-Object { -not $_.IsTarget } | Sort-Object KB)
    $target  = @(@($files) | Where-Object { $_.IsTarget })

    return @($prereqs + $target)
}

function Get-RemoteFileSize {
    <#
        Per-file Content-Length from the CDN. The catalog row's size column covers only
        the primary CU, so it cannot be used to verify prerequisite files.
    #>
    param([Parameter(Mandatory=$true)][string]$Url)

    try {
        $h   = Invoke-WebRequest -Uri $Url -Method Head -UseBasicParsing -Headers @{ 'User-Agent' = $Script:UserAgent } -TimeoutSec 120
        $len = @($h.Headers['Content-Length'])[0]
        if ($len) { return [int64]$len }
    }
    catch { }
    return $null
}

function Format-Bytes {
    param([int64]$Bytes)
    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N1} MB" -f ($Bytes / 1MB) }
    return "{0:N0} KB" -f ($Bytes / 1KB)
}

function Invoke-StreamDownload {
    <#
        Streams the response to disk in chunks so download progress can be reported.

        Invoke-WebRequest -OutFile is not used here: it gives no usable progress, and on
        Windows PowerShell 5.1 (likely in WinPE) its own progress handling is what makes
        multi-GB transfers crawl. HttpWebRequest works on both 5.1 and 7.
    #>
    param(
        [Parameter(Mandatory=$true)][string]$Url,
        [Parameter(Mandatory=$true)][string]$OutFile,
        [int64]$ExpectedBytes
    )

    $request = [System.Net.HttpWebRequest]::Create($Url)
    $request.UserAgent        = $Script:UserAgent
    $request.Timeout          = 120000      # connect
    $request.ReadWriteTimeout = 300000      # stall between chunks
    $request.AllowAutoRedirect = $true

    $response = $null; $input = $null; $output = $null
    try {
        $response = $request.GetResponse()

        # Prefer the server's ContentLength for the progress denominator - it describes
        # this transfer. $ExpectedBytes is what the caller verifies against afterwards.
        $total = if ([int64]$response.ContentLength -gt 0) { [int64]$response.ContentLength } else { $ExpectedBytes }
        $input  = $response.GetResponseStream()
        $output = [System.IO.File]::Open($OutFile, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write)

        $buffer    = New-Object byte[] (1MB)
        $totalRead = [int64]0
        $clock     = [System.Diagnostics.Stopwatch]::StartNew()
        $lastTick  = [int64]0
        $lineWidth = 0

        # In-place `r repainting only works on a live console. When output is redirected
        # (transcript, log file, CI) it collapses into one unreadable line, so fall back
        # to one discrete line per 10%.
        $redirected = [Console]::IsOutputRedirected
        $nextLogPct = 0

        while (($read = $input.Read($buffer, 0, $buffer.Length)) -gt 0) {
            $output.Write($buffer, 0, $read)
            $totalRead += $read

            $elapsedMs = $clock.ElapsedMilliseconds
            $done      = ($totalRead -eq $total)

            # Repaint at most ~2x/second so the console (and any transcript) stays readable.
            if ($elapsedMs - $lastTick -ge 500 -or $done) {
                $lastTick = $elapsedMs
                $seconds  = [math]::Max($clock.Elapsed.TotalSeconds, 0.001)
                $rate     = $totalRead / $seconds

                if ($total -gt 0) {
                    $pct  = [math]::Min(100, [math]::Round(($totalRead / $total) * 100, 1))
                    $eta  = if ($rate -gt 0) { [TimeSpan]::FromSeconds(($total - $totalRead) / $rate) } else { [TimeSpan]::Zero }
                    $line = "    {0,5:N1}%  {1} / {2}  at {3}/s  eta {4:mm\:ss}" -f `
                            $pct, (Format-Bytes $totalRead), (Format-Bytes $total), (Format-Bytes ([int64]$rate)), $eta
                }
                else {
                    $line = "    {0} at {1}/s" -f (Format-Bytes $totalRead), (Format-Bytes ([int64]$rate))
                }

                if ($redirected) {
                    $reached = if ($total -gt 0) { [int][math]::Floor(($totalRead / $total) * 100) } else { 0 }
                    if ($reached -ge $nextLogPct -or $done) {
                        Write-Host $line.TrimEnd()
                        $nextLogPct = [math]::Max($nextLogPct + 10, $reached + 10)
                    }
                }
                else {
                    # Pad to erase the tail of any longer previous line, then return to column 0.
                    if ($line.Length -lt $lineWidth) { $line = $line.PadRight($lineWidth) }
                    $lineWidth = $line.Length
                    Write-Host "`r$line" -NoNewline
                }
            }
        }

        $output.Flush()
        if (-not $redirected) { Write-Host "" }   # close off the in-place progress line
        return $totalRead
    }
    finally {
        if ($output)   { $output.Dispose() }
        if ($input)    { $input.Dispose() }
        if ($response) { $response.Dispose() }
    }
}

function Save-UpdatePackage {
    param(
        [Parameter(Mandatory=$true)][string]$Url,
        [Parameter(Mandatory=$true)][string]$DestinationDir,
        [int64]$ExpectedBytes
    )

    Ensure-Directory -Path $DestinationDir

    $fileName = Split-Path -Path ($Url -split '\?')[0] -Leaf
    $destPath = [IO.Path]::Combine($DestinationDir, $fileName)

    if (Test-Path -LiteralPath $destPath) {
        $existing = (Get-Item -LiteralPath $destPath).Length
        if (-not $ExpectedBytes -or $existing -eq $ExpectedBytes) {
            Write-Host "Already present ($([math]::Round($existing/1MB,1)) MB): $destPath"
            return $destPath
        }
        Write-Host ("Existing file is {0} bytes but the catalog reports {1}; re-downloading." -f $existing, $ExpectedBytes) -ForegroundColor Yellow
        Remove-Item -LiteralPath $destPath -Force
    }

    $sizeText = if ($ExpectedBytes) { " ($(Format-Bytes $ExpectedBytes))" } else { '' }
    Write-Host "Downloading$sizeText -> $destPath"

    # Download to .partial so an interrupted transfer is never mistaken for a good file.
    $tmp   = "$destPath.partial"
    $start = Get-Date
    try {
        $null = Invoke-StreamDownload -Url $Url -OutFile $tmp -ExpectedBytes $ExpectedBytes
    }
    catch {
        if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
        throw "Failed to download update package: $($_.Exception.Message)"
    }

    $got = (Get-Item -LiteralPath $tmp).Length
    if ($ExpectedBytes -and $got -ne $ExpectedBytes) {
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
        throw "Download size mismatch: got $got bytes, expected $ExpectedBytes."
    }

    Move-Item -LiteralPath $tmp -Destination $destPath -Force

    $took = (Get-Date) - $start
    Write-Host ("Downloaded {0} in {1:mm\:ss}: {2}" -f (Format-Bytes $got), $took, $destPath) -ForegroundColor Green
    return $destPath
}

function Get-PackageForKB {
    <#
        Returns the full ordered set of package files for this KB/architecture
        (prerequisites first, target CU last), downloading any not already available
        locally. Returns an empty array on failure.
    #>
    param(
        [Parameter(Mandatory=$true)][string]$KB,
        [Parameter(Mandatory=$true)][ValidateSet('x64','arm64')][string]$Arch,
        [string]$DestinationDir
    )

    $row = Get-CatalogUpdateRow -KB $KB -Arch $Arch
    if (-not $row) {
        if ($OpenCatalog) {
            $browse = "$($Script:CatalogBase)/Search.aspx?q=" + [uri]::EscapeDataString($KB)
            Write-Host "Opening catalog for manual download: $browse"
            try { Start-Process -FilePath $browse } catch { Write-Host "Unable to open a browser here." -ForegroundColor Yellow }
        }
        return @()
    }

    Write-Host "Selected: $($row.Title)" -ForegroundColor Green

    $files = @(Get-CatalogDownloadFile -UpdateId $row.UpdateId -KB $KB)
    if ($files.Count -eq 0) { return @() }

    Write-Host "Package set ($($files.Count) file(s), listed in DISM apply order):"
    foreach ($f in $files) {
        $tag = if ($f.IsTarget) { 'target CU   ' } else { 'prerequisite' }
        Write-Host "  [$tag] $($f.FileName)"
    }

    $resolved = New-Object System.Collections.Generic.List[object]

    foreach ($f in $files) {
        $path = $null

        # The source drive acts as a cache so a package already staged is not re-fetched.
        if (-not $SkipLocalSearch) {
            $local = Find-FileOnDrive -RootPath $SourceDrive -FileName $f.FileName
            if ($local) {
                Write-Host "  reusing from ${SourceDrive}: $local" -ForegroundColor Green
                $path = $local
            }
        }

        if (-not $path -and $WhatIf) {
            # [IO.Path]::Combine, not Join-Path: the destination drive (e.g. E:) may not
            # be mounted during a dry run, and Join-Path fails on an unknown PSDrive.
            $path = [IO.Path]::Combine($DestinationDir, $f.FileName)
            Write-Host "  WhatIf: would download $($f.Url) -> $path" -ForegroundColor Cyan
        }

        if (-not $path) {
            $bytes = Get-RemoteFileSize -Url $f.Url
            $path  = Save-UpdatePackage -Url $f.Url -DestinationDir $DestinationDir -ExpectedBytes $bytes
        }

        $resolved.Add([pscustomobject]@{
            Path     = $path
            FileName = $f.FileName
            KB       = $f.KB
            IsTarget = $f.IsTarget
        })
    }

    # .ToArray(), not @($resolved): on PowerShell 7.6 the array subexpression operator
    # throws "Argument types do not match" for a List[object] holding PSCustomObjects.
    return $resolved.ToArray()
}

#endregion

#region Local flash-drive fallback

function Find-FileOnDrive {
    param([string]$RootPath, [string]$FileName)

    if (-not (Test-Path -LiteralPath $RootPath)) { return $null }

    return Get-ChildItem -Path $RootPath -Recurse -File -Filter $FileName -ErrorAction SilentlyContinue |
           Select-Object -First 1 -ExpandProperty FullName
}

function Find-PackageByKBOnDrive {
    param([string]$RootPath, [string]$KB, [string]$Arch)

    if (-not (Test-Path -LiteralPath $RootPath)) { return $null }

    $candidates = Get-ChildItem -Path $RootPath -Recurse -File -Include '*.msu','*.cab' -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match [regex]::Escape($KB) }

    if ($Arch) {
        $candidates = $candidates | Where-Object { $_.Name -match "(?<![a-z0-9])$Arch(?![a-z0-9])" }
    }

    if ($candidates) {
        return $candidates | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName
    }
    return $null
}

function Find-PackagesOnDrive {
    param([string]$RootPath)

    if (-not (Test-Path -LiteralPath $RootPath)) { return $null }

    $candidates = Get-ChildItem -Path $RootPath -Recurse -File -Include '*.msu','*.cab' -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Name -match 'KB\d{6,7}' -or
            $_.Name -match 'Cumulative' -or
            $_.Name -match 'Windows11|Win11|Windows_11'
        }

    if ($candidates) {
        return $candidates | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName
    }
    return $null
}

#endregion

#region Part 3 - apply offline with DISM

function Assert-OfflineTarget {
    param([string]$TargetRoot)

    $targetWindows = Join-Path $TargetRoot 'Windows'
    if (-not (Test-Path -LiteralPath $targetWindows)) {
        throw "Target Windows directory not found: $targetWindows"
    }

    $targetResolved  = (Resolve-Path -LiteralPath $targetWindows).Path.TrimEnd('\')
    $runningResolved = (Resolve-Path -LiteralPath $env:SystemRoot).Path.TrimEnd('\')

    if ($targetResolved -ieq $runningResolved) {
        throw ("Refusing to service the RUNNING operating system ($targetResolved) with dism /Image. " +
               "Boot into WinPE and target the offline volume, or run dism /Online manually.")
    }
}

function Get-OfflineImageBuild {
    <#
        Reads CurrentBuild/UBR from the offline image's SOFTWARE hive, e.g. 26200.8457.
        Knowing this up front lets a non-applicable update be caught before a ~5 GB
        download rather than after DISM rejects it. Returns $null if it cannot be read.
    #>
    param([string]$TargetRoot)

    $hive = Join-Path $TargetRoot 'Windows\System32\config\SOFTWARE'
    if (-not (Test-Path -LiteralPath $hive)) { return $null }

    $mountKey = 'AutoCU_Offline'
    $loaded   = $false

    try {
        $null = reg load "HKLM\$mountKey" $hive 2>&1
        if ($LASTEXITCODE -ne 0) { return $null }
        $loaded = $true

        $cv = Get-ItemProperty -Path "HKLM:\$mountKey\Microsoft\Windows NT\CurrentVersion" -ErrorAction Stop
        if (-not $cv.CurrentBuild) { return $null }

        $ubr = if ($null -ne $cv.UBR) { [int]$cv.UBR } else { 0 }
        return [pscustomobject]@{
            Build   = [int]$cv.CurrentBuild
            UBR     = $ubr
            Display = "$([int]$cv.CurrentBuild).$ubr"
        }
    }
    catch { return $null }
    finally {
        if ($loaded) {
            # Release the handles Get-ItemProperty took, or the hive will not unload.
            [gc]::Collect()
            [gc]::WaitForPendingFinalizers()
            $null = reg unload "HKLM\$mountKey" 2>&1
        }
    }
}

function Initialize-ScratchDir {
    <#
        WinPE's default scratch space is the X: RAM disk, which DISM warns is too small
        ("recommended size is at least 1024 MB"). Expanding a ~5 GB MSU there fails, so
        servicing always runs with an explicit /ScratchDir on a real volume.
    #>
    param([string]$TargetRoot)

    $path = if ($ScratchDir) { $ScratchDir } else { Join-Path $TargetRoot 'OSDScratch' }
    Ensure-Directory -Path $path
    return (Resolve-Path -LiteralPath $path).Path
}

function Install-CumulativeUpdate {
    param(
        [string]$PackageFile,
        [string]$TargetRoot,
        [string]$Scratch,
        [switch]$NonFatal
    )

    # Under -WhatIf nothing was downloaded, so the file legitimately does not exist yet.
    if (-not $WhatIf -and -not (Test-Path -LiteralPath $PackageFile)) {
        throw "Package not found: $PackageFile"
    }
    Assert-OfflineTarget -TargetRoot $TargetRoot

    if (-not $Scratch) { $Scratch = Initialize-ScratchDir -TargetRoot $TargetRoot }

    $logPath = Join-Path $TargetRoot 'Windows\Logs\AutoCU-install.log'
    Write-Host "Applying $(Split-Path $PackageFile -Leaf) to offline image $TargetRoot" -ForegroundColor Cyan

    if ($WhatIf) {
        Write-Host "WhatIf: dism /Image:$TargetRoot /Add-Package /PackagePath:$PackageFile /IgnoreCheck /ScratchDir:$Scratch /LogPath:$logPath" -ForegroundColor Cyan
        return $true
    }

    dism /Image:$TargetRoot /Add-Package /PackagePath:$PackageFile /IgnoreCheck /ScratchDir:$Scratch /LogPath:$logPath

    if ($LASTEXITCODE -ne 0) {
        $msg = "DISM failed with exit code $LASTEXITCODE for $(Split-Path $PackageFile -Leaf). See $logPath"
        if ($NonFatal) {
            Write-Host "WARNING: $msg" -ForegroundColor Yellow
            return $false
        }
        throw $msg
    }

    Write-Host "DISM completed successfully." -ForegroundColor Green
    return $true
}

#endregion

#==============================  Main  ==============================

# Install mode is explicit: use the package the caller supplied and skip all discovery.
if ($Mode -eq 'Install') {
    if (-not $PackagePath) { throw "-Mode Install requires -PackagePath." }
    Install-CumulativeUpdate -PackageFile $PackagePath -TargetRoot $TargetRoot | Out-Null
    Write-Host 'Install completed.'
    exit 0
}

# Validate the servicing target BEFORE downloading ~5 GB. Previously the offline-target
# check ran only at install time, after the download had already completed.
$imageBuild = $null
if ($Mode -eq 'All') {
    Assert-OfflineTarget -TargetRoot $TargetRoot
    $imageBuild = Get-OfflineImageBuild -TargetRoot $TargetRoot
    if ($imageBuild) { Write-Host "Offline image build: $($imageBuild.Display)" -ForegroundColor Cyan }
    else             { Write-Host "Could not read the offline image build from $TargetRoot." -ForegroundColor Yellow }
}

$targetKB    = $null
$targetBuild = $null
if ($KB) {
    $targetKB = if ($KB -match '^\d+$') { "KB$KB" } else { $KB.ToUpper() }
    Write-Host "Using caller-supplied KB: $targetKB"
}
else {
    $latest = Get-LatestCU -DaysBack $DaysBack -IncludePreview:$IncludePreview
    if (-not $latest) { throw "Could not determine the current 25H2 cumulative update from the update-history page." }
    $targetKB    = $latest.KB
    $targetBuild = $latest.Revision
    Write-Host ("Latest 25H2 CU: {0}  build {1}  [{2}]" -f $latest.KB, $latest.Build, $latest.ReleaseType) -ForegroundColor Green
}

if ($Mode -eq 'Find') {
    Write-Host "Find complete: $targetKB"
    exit 0
}

# Nothing to do if the image already carries this CU or newer.
if ($imageBuild -and $targetBuild -and
    $imageBuild.Build -eq [int]$Script:BuildFamily25H2 -and $imageBuild.UBR -ge $targetBuild) {
    Write-Host ("Image is already at $($imageBuild.Display), which is at or above $targetKB " +
                "($($Script:BuildFamily25H2).$targetBuild). Nothing to apply.") -ForegroundColor Green
    exit 0
}

# Deliberately NOT named $packagePath: PowerShell variables are case-insensitive, so
# that would silently overwrite the $PackagePath parameter.
$resolvedPackages = [ordered]@{}

foreach ($arch in $Architecture) {
    Write-Host ""
    Write-Host "--- $targetKB / $arch ---" -ForegroundColor Cyan

    $pkgSet = @(Get-PackageForKB -KB $targetKB -Arch $arch -DestinationDir $DestinationPath)

    if ($pkgSet.Count -eq 0 -and -not $SkipLocalSearch) {
        Write-Host "Catalog path failed for $arch; falling back to newest package on $SourceDrive." -ForegroundColor Yellow
        $fallback = Find-PackagesOnDrive -RootPath $SourceDrive
        if ($fallback) {
            Write-Host "WARNING: fallback package is unverified and may omit checkpoint prerequisites." -ForegroundColor Yellow
            $pkgSet = @($fallback)
        }
    }

    if ($pkgSet.Count -gt 0) { $resolvedPackages[$arch] = $pkgSet }
    else                     { Write-Host "No package obtained for $targetKB / $arch." -ForegroundColor Yellow }
}

if ($resolvedPackages.Count -eq 0) {
    if ($WhatIf) { Write-Host "WhatIf: no package located." -ForegroundColor Yellow; exit 0 }
    throw "Failed to locate a package for $targetKB. Manual download required."
}

if ($Mode -eq 'Download') {
    Write-Host ""
    Write-Host "Download complete:" -ForegroundColor Green
    foreach ($entry in $resolvedPackages.GetEnumerator()) {
        Write-Host "  $($entry.Key):"
        foreach ($file in @($entry.Value)) { Write-Host "    $($file.Path)" }
    }
    exit 0
}

# Mode = All -> install. Only one architecture can apply to a given offline image.
if ($resolvedPackages.Count -gt 1) {
    throw ("Multiple architectures were resolved ({0}). Re-run with a single -Architecture, " -f ($resolvedPackages.Keys -join ', ')) +
          "or use -Mode Install -PackagePath <file> to choose explicitly."
}

$selected = @(@($resolvedPackages.Values)[0])
$scratch  = Initialize-ScratchDir -TargetRoot $TargetRoot
Write-Host "DISM scratch directory: $scratch"

try {
    # Checkpoint prerequisites are only needed by images that predate them. Applying
    # KB5043080 to an already-serviced image fails with 0x80070228 / 552, so they are
    # opt-in and never abort the run - the target CU is what matters.
    foreach ($file in ($selected | Where-Object { -not $_.IsTarget })) {
        if (-not $IncludePrerequisites) {
            Write-Host "Skipping prerequisite $($file.KB) (pass -IncludePrerequisites for a base image that needs it)." -ForegroundColor Yellow
            continue
        }
        Install-CumulativeUpdate -PackageFile $file.Path -TargetRoot $TargetRoot -Scratch $scratch -NonFatal | Out-Null
    }

    foreach ($file in ($selected | Where-Object { $_.IsTarget })) {
        Install-CumulativeUpdate -PackageFile $file.Path -TargetRoot $TargetRoot -Scratch $scratch | Out-Null
    }
}
finally {
    # Only clean up a scratch folder this script created.
    if (-not $ScratchDir -and (Test-Path -LiteralPath $scratch)) {
        Remove-Item -LiteralPath $scratch -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Write-Host 'AutoCU update processing completed.'
exit 0
