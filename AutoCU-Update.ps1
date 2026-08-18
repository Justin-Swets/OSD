[CmdletBinding()]
param(
	[string]$SourceDrive = 'E:',
	[string]$DestinationPath = 'E:\Latest-CU',
	[string]$TargetRoot = 'C:\',
	[string]$SearchQuery = 'Cumulative Update for Windows 11, version 25H2 x64',
	[switch]$WhatIf
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

function Get-LatestKBFromUpdateHistory {
	param(
		[string]$UpdateHistoryUrl = 'https://support.microsoft.com/en-us/servicing/os/windows-11/2025/07/windows-11-version-25h2-update-history',
		[int]$DaysBack = 30
	)

	Write-Host "Checking Windows 11 update-history page: $UpdateHistoryUrl"
	try {
		[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
		$resp = Invoke-WebRequest -Uri $UpdateHistoryUrl -UseBasicParsing -Headers @{ 'User-Agent' = 'Mozilla/5.0' }
	}
	catch {
		Write-Host "Failed to fetch update-history page: $($_.Exception.Message)" -ForegroundColor Yellow
		return $null
	}

	$content = $resp.Content
	$results = @()

	# Find every KB occurrence and try to locate a nearby date (within 500 chars before the KB)
	$kbMatches = [regex]::Matches($content, 'KB\d{6,7}') | ForEach-Object { $_.Value } | Select-Object -Unique
	foreach ($kb in $kbMatches) {
		$pos = $content.IndexOf($kb)
		if ($pos -lt 0) { continue }
		$start = [Math]::Max(0, $pos - 500)
		$length = [Math]::Min(500 + $kb.Length, $content.Length - $start)
		$segment = $content.Substring($start, $length)

		# Try common date formats near the KB
		$dateMatch = [regex]::Match($segment, '([A-Za-z]+\s+\d{1,2},\s+\d{4})|((?:20)\d{2}-\d{2}-\d{2})')
		if ($dateMatch.Success) {
			$dateText = $dateMatch.Value
			try {
				$d = [datetime]::Parse($dateText)
				$is25 = $segment -match '25H2|version\s*25H2|25-H2'
				$results += [pscustomobject]@{ KB = $kb; Date = $d; Segment = $segment; Is25H2 = $is25 }
			}
			catch { }
		}
	}

	if ($results.Count -eq 0) {
		Write-Host "No KB entries with nearby dates found on update-history page." -ForegroundColor Yellow
		return $null
	}

	$threshold = (Get-Date).AddDays(-$DaysBack)
	$recent = $results | Sort-Object Date -Descending | Where-Object { $_.Date -ge $threshold }
	if ($recent.Count -eq 0) {
		Write-Host "Found KBs but none within the last $DaysBack days." -ForegroundColor Yellow
		return $null
	}

	Write-Host "Update-history candidates:" -ForegroundColor Cyan
	$recent | ForEach-Object { Write-Host "  $($_.KB) - $($_.Date.ToShortDateString()) (25H2? $($_.Is25H2))" }

	# Prefer the first candidate (newest) that explicitly references 25H2 in its segment
	$first25 = $recent | Where-Object { $_.Is25H2 } | Select-Object -First 1
	if ($first25) { return $first25.KB }

	Write-Host "No update-history candidates explicitly mentioning 25H2; returning null to allow catalog fallback." -ForegroundColor Yellow
	return $null
}

function Find-PackageByKBOnDrive {
	param(
		[string]$RootPath,
		[string]$KB
	)
	if (-not (Test-Path -LiteralPath $RootPath)) { return $null }
	$candidates = Get-ChildItem -Path $RootPath -Recurse -File -ErrorAction SilentlyContinue |
		Where-Object { $_.Name -match [regex]::Escape($KB) -and ($_.Extension -in '.msu','.cab') }
	if ($candidates) { return $candidates | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName }
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

	$logPath = Join-Path $TargetRoot 'Windows\Logs\AutoCU-install.log'
	Write-Host "Installing update package from $PackagePath"
	if ($WhatIf) {
		Write-Host "WhatIf: would run: dism /Image:$TargetRoot /Add-Package /PackagePath:$PackagePath /IgnoreCheck /LogPath:$logPath"
	}
	else {
		dism /Image:$TargetRoot /Add-Package /PackagePath:$PackagePath /IgnoreCheck /LogPath:$logPath
	}
}


# Main flow: prefer update-history page
$kbFromHistory = Get-LatestKBFromUpdateHistory
if ($kbFromHistory) {
	Write-Host "Found recent KB from update-history: $kbFromHistory"
	$packageOnDrive = Find-PackageByKBOnDrive -RootPath $SourceDrive -KB $kbFromHistory
	if ($packageOnDrive) {
		Write-Host "Found $kbFromHistory package on source drive: $packageOnDrive"
		$packagePath = $packageOnDrive
	}
	else {
		Write-Host "No $kbFromHistory file on flash; searching Microsoft Update Catalog for $kbFromHistory"
		$updateId = Get-UpdateIdFromSearchPage -Query $kbFromHistory
		if (-not $updateId) {
			Write-Host "Could not find updateid for $kbFromHistory in catalog; falling back to any package on drive" -ForegroundColor Yellow
			$packageOnDrive = Find-PackagesOnDrive -RootPath $SourceDrive
			if ($packageOnDrive) { $packagePath = $packageOnDrive }
			else {
				if ($WhatIf) { Write-Host "WhatIf: no package on drive for $kbFromHistory; would require manual download." -ForegroundColor Yellow; exit 0 }
				else { throw "Failed to locate any package to install. Manual download required." }
			}
		}
		else {
			$fileUrl = Get-DownloadUrlFromUpdateId -UpdateId $updateId
			if (-not $fileUrl) { throw "Could not extract a direct download URL for updateid $updateId. Manual download required." }
			$packagePath = Download-UpdatePackage -Url $fileUrl -DestinationDir $DestinationPath
			Write-Host "Downloaded package to $packagePath"
		}
	}
}
else {
	Write-Host "Update-history parsing failed or no recent CU found; attempting catalog search as fallback."
	# Try Update Catalog fallback using broad SearchQuery
	try {
		$updateId = Get-UpdateIdFromSearchPage -Query $SearchQuery
	}
	catch {
		$updateId = $null
	}

	if ($updateId) {
		Write-Host "Found updateid from catalog search: $updateId"
		$fileUrl = Get-DownloadUrlFromUpdateId -UpdateId $updateId
		if ($fileUrl) {
			$packagePath = Download-UpdatePackage -Url $fileUrl -DestinationDir $DestinationPath
			Write-Host "Downloaded package to $packagePath"
		}
		else {
			Write-Host "Catalog search returned updateid but no direct download URL. Falling back to flash-drive packages." -ForegroundColor Yellow
			$packageOnDrive = Find-PackagesOnDrive -RootPath $SourceDrive
			if ($packageOnDrive) { $packagePath = $packageOnDrive }
			else { throw "No package found on $SourceDrive and catalog fallback failed. Manual download required." }
		}
	}
	else {
		Write-Host "Catalog search did not return an updateid; using most-recent package on flash drive." -ForegroundColor Yellow
		$packageOnDrive = Find-PackagesOnDrive -RootPath $SourceDrive
		if ($packageOnDrive) {
			Write-Host "Found package on source drive: $packageOnDrive"
			$packagePath = $packageOnDrive
		}
		else {
			throw "No package found on $SourceDrive and update-history/catalog lookup provided no candidate. Manual download required."
		}
	}
}

if (-not $packagePath) {
	if ($WhatIf) {
		Write-Host "WhatIf: no package was located. Script would have required a manual download or further catalog investigation." -ForegroundColor Yellow
		exit 0
	}
	else {
		throw "No package located to install. Manual download required."
	}
}

Install-CumulativeUpdate -PackagePath $packagePath -TargetRoot $TargetRoot

Write-Host 'AutoCU update processing completed.'

