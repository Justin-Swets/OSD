[CmdletBinding()]
param(
	[string]$SourceDrive = 'E:',
	[string]$DestinationPath = 'E:\Latest-CU',
	[string]$TargetRoot = 'C:\',
	[string]$SearchQuery = 'Cumulative Update for Windows 11, version 25H2 x64',
	[ValidateSet('All','Find','Download','Install')][string]$Mode = 'All',
	[string]$KB,
	[string]$PackagePath,
	[switch]$WhatIf,
	[switch]$OpenCatalog
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

	$searchUrls = @(
		"https://www.catalog.update.microsoft.com/Search.aspx?q=" + [uri]::EscapeDataString($Query),
		"https://www.catalog.update.microsoft.com/Search.aspx?q=" + [uri]::EscapeDataString("$Query x64"),
		"https://www.catalog.update.microsoft.com/Search.aspx?q=" + [uri]::EscapeDataString("$Query Windows 11")
	)

	foreach ($searchUrl in $searchUrls) {
		Write-Host "Querying Microsoft Update Catalog: $searchUrl"
		try {
			$resp = Invoke-WebRequest -Uri $searchUrl -UseBasicParsing -Headers @{ 'User-Agent' = 'Mozilla/5.0' }
		}
		catch {
			Write-Host ("Catalog query failed for {0}: {1}" -f $searchUrl, $_.Exception.Message) -ForegroundColor Yellow
			continue
		}

		$content = $resp.Content

		# Try multiple extraction patterns for updateid
		$patterns = @(
			'updateid=([0-9A-Fa-f\-]{36})',
			'updateid%3d([0-9A-Fa-f\-]{36})',
			'"updateid"\s*:\s*"([0-9A-Fa-f\-]{36})"',
			'data-updateid="([0-9A-Fa-f\-]{36})"',
			'ScopedViewInline.aspx\?updateid=([0-9A-Fa-f\-]{36})'
		)

		foreach ($pat in $patterns) {
			$m = [regex]::Matches($content, $pat, 'IgnoreCase') | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique
			if ($m -and $m.Count -gt 0) { return $m[0] }
		}

		# As a last resort, look for GUID-like strings near the KB number
		$kbPattern = [regex]::Escape($Query)
		$pos = $content.IndexOf($Query)
		if ($pos -ge 0) {
			$start = [Math]::Max(0, $pos - 400)
			$len = [Math]::Min(800, $content.Length - $start)
			$seg = $content.Substring($start, $len)
			$guidMatch = [regex]::Match($seg, '([0-9A-Fa-f]{8}\-[0-9A-Fa-f]{4}\-[0-9A-Fa-f]{4}\-[0-9A-Fa-f]{4}\-[0-9A-Fa-f]{12})')
			if ($guidMatch.Success) { return $guidMatch.Groups[1].Value }
		}
	}

	return $null
}

function Get-DownloadUrlFromUpdateId {
	param([string]$UpdateId)

	$viewUrl = "https://www.catalog.update.microsoft.com/ScopedViewInline.aspx?updateid=$UpdateId"
	Write-Host "Fetching update details: $viewUrl"

	try {
		$resp = Invoke-WebRequest -Uri $viewUrl -UseBasicParsing -Headers @{ 'User-Agent' = 'Mozilla/5.0' }
	}
	catch {
		throw ("Failed to fetch update detail page: {0}" -f $_.Exception.Message)
	}

	$content = $resp.Content

	# Primary: look for direct download links hosted on download.windowsupdate.com ending in .msu or .cab
	$patterns = @(
		'(https?://download\.windowsupdate\.com/[^"'']+?\.(msu|cab))',
		'data-downloadlink="(https?://download\.windowsupdate\.com/[^"]+?)"',
		'"downloadUrl"\s*:\s*"(https?://download\.windowsupdate\.com/[^"]+?)"'
	)

	foreach ($pat in $patterns) {
		$m = [regex]::Match($content, $pat, 'IgnoreCase')
		if ($m.Success) { return $m.Groups[1].Value }
	}

	# If not found, attempt to follow any DownloadDialog link from the scoped view
	$dlgMatch = [regex]::Match($content, 'href\s*=\s*"([^"]*DownloadDialog\.aspx[^\"]*)"', 'IgnoreCase')
	if ($dlgMatch.Success) {
		$dlg = $dlgMatch.Groups[1].Value
		if ($dlg -notmatch '^http') { $dlg = 'https://www.catalog.update.microsoft.com' + $dlg }
		try {
			$dlgResp = Invoke-WebRequest -Uri $dlg -UseBasicParsing -Headers @{ 'User-Agent' = 'Mozilla/5.0'; 'Referer' = $viewUrl }
			$dlgContent = $dlgResp.Content
			foreach ($pat in $patterns) {
				$m2 = [regex]::Match($dlgContent, $pat, 'IgnoreCase')
				if ($m2.Success) { return $m2.Groups[1].Value }
			}
		}
		catch { }
	}

	# Try DownloadDialog.aspx directly with common query patterns and a Referer/cookie container
	$tryUrls = @(
		"https://www.catalog.update.microsoft.com/DownloadDialog.aspx?updateid=$UpdateId",
		"https://www.catalog.update.microsoft.com/DownloadDialog.aspx?updateid=$UpdateId&ln=en-US"
	)

	$wc = New-Object System.Net.CookieContainer
	foreach ($u in $tryUrls) {
		try {
			$dlgResp2 = Invoke-WebRequest -Uri $u -UseBasicParsing -Headers @{ 'User-Agent' = 'Mozilla/5.0'; 'Referer' = $viewUrl } -WebSession @{ CookieContainer = $wc }  -ErrorAction Stop
		}
		catch {
			# Some environments block this, continue to next attempt
			continue
		}
		$dlgContent2 = $dlgResp2.Content
		foreach ($pat in $patterns) {
			$m3 = [regex]::Match($dlgContent2, $pat, 'IgnoreCase')
			if ($m3.Success) { return $m3.Groups[1].Value }
		}
	}

	# Last resort: look for any GUID-prefixed resources in the scoped view and attempt to construct a download dialog
	$guidMatch = [regex]::Match($content, '([0-9A-Fa-f]{8}\-[0-9A-Fa-f]{4}\-[0-9A-Fa-f]{4}\-[0-9A-Fa-f]{4}\-[0-9A-Fa-f]{12})')
	if ($guidMatch.Success) {
		$tryGuid = $guidMatch.Groups[1].Value
		$tryUrl = "https://www.catalog.update.microsoft.com/DownloadDialog.aspx?updateid=$tryGuid"
		try {
			$tryResp = Invoke-WebRequest -Uri $tryUrl -UseBasicParsing -Headers @{ 'User-Agent' = 'Mozilla/5.0'; 'Referer' = $viewUrl }
			$tryContent = $tryResp.Content
			foreach ($pat in $patterns) {
				$m4 = [regex]::Match($tryContent, $pat, 'IgnoreCase')
				if ($m4.Success) { return $m4.Groups[1].Value }
			}
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

function Get-RecentKBCandidates {
	param(
		[string]$UpdateHistoryUrl = 'https://support.microsoft.com/en-us/servicing/os/windows-11/2025/07/windows-11-version-25h2-update-history',
		[int]$DaysBack = 30
	)

	try {
		[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
		$resp = Invoke-WebRequest -Uri $UpdateHistoryUrl -UseBasicParsing -Headers @{ 'User-Agent' = 'Mozilla/5.0' }
	}
	catch {
		Write-Host "Failed to fetch update-history page: $($_.Exception.Message)" -ForegroundColor Yellow
		return @()
	}

	$content = $resp.Content
	$results = @()

	$kbMatches = [regex]::Matches($content, 'KB\d{6,7}') | ForEach-Object { $_.Value } | Select-Object -Unique
	foreach ($kb in $kbMatches) {
		$pos = $content.IndexOf($kb)
		if ($pos -lt 0) { continue }
		$start = [Math]::Max(0, $pos - 500)
		$length = [Math]::Min(500 + $kb.Length, $content.Length - $start)
		$segment = $content.Substring($start, $length)
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

	$threshold = (Get-Date).AddDays(-$DaysBack)
	$recent = $results | Where-Object { $_.Date -ge $threshold } | Sort-Object Date -Descending
	return $recent
}

function Get-DownloadUrlForKBWithRetries {
	param(
		[string]$KB,
		[int]$Attempts = 3,
		[int]$DelaySeconds = 2
	)

	$queries = @("$KB Windows 11 25H2", "$KB Windows 11", "$KB")
	foreach ($q in $queries) {
		for ($i = 1; $i -le $Attempts; $i++) {
			try {
				$uid = Get-UpdateIdFromSearchPage -Query $q
			}
			catch { $uid = $null }

			if ($uid) {
				for ($j = 1; $j -le $Attempts; $j++) {
					try {
						$fileUrl = Get-DownloadUrlFromUpdateId -UpdateId $uid
						if ($fileUrl) { return $fileUrl }
					}
					catch { }
					Start-Sleep -Seconds $DelaySeconds
				}
			}

			Start-Sleep -Seconds $DelaySeconds
		}
	}

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

# High-level part 1: find the current CU KB ID
function Find-CurrentCUKB {
	param([int]$DaysBack = 30)
	return Get-LatestKBFromUpdateHistory -DaysBack $DaysBack
}

# High-level part 2: download the .msu for a KB (uses existing helpers)
function Download-MSUForKB {
	param(
		[Parameter(Mandatory=$true)][string]$KBToGet,
		[string]$DestinationDir = $DestinationPath
	)

	# Check flash for matching package first
	$existing = Find-PackageByKBOnDrive -RootPath $SourceDrive -KB $KBToGet
	if ($existing) { Write-Host "Found $KBToGet on drive: $existing"; return $existing }

	# Try to get direct download URL via catalog
	$fileUrl = Get-DownloadUrlForKBWithRetries -KB $KBToGet -Attempts 3 -DelaySeconds 2
	if (-not $fileUrl) {
		$browseUrl = "https://www.catalog.update.microsoft.com/Search.aspx?q=" + [uri]::EscapeDataString($KBToGet)
		Write-Host "Could not automatically locate .msu for $KBToGet. Catalog: $browseUrl" -ForegroundColor Yellow
		if ($OpenCatalog) { try { Start-Process -FilePath $browseUrl } catch { Write-Host "Unable to open browser." -ForegroundColor Yellow } }
		return $null
	}

	$fileName = Split-Path -Path $fileUrl -Leaf
	$destPath = Join-Path $DestinationDir $fileName
	if ($WhatIf) {
		Write-Host "WhatIf: would download $fileUrl -> $destPath"
		return $destPath
	}

	return Download-UpdatePackage -Url $fileUrl -DestinationDir $DestinationDir
}

# High-level part 3: apply the msu with DISM
function Apply-MSUWithDISM {
	param(
		[Parameter(Mandatory=$true)][string]$MSUPath,
		[string]$Target = $TargetRoot
	)
	Install-CumulativeUpdate -PackagePath $MSUPath -TargetRoot $Target
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
			else {
				if ($OpenCatalog) {
					$browseUrl = "https://www.catalog.update.microsoft.com/Search.aspx?q=" + [uri]::EscapeDataString($kbFromHistory)
					Write-Host "Opening catalog search page for manual download: $browseUrl"
					try { Start-Process -FilePath $browseUrl } catch { Write-Host "Unable to open browser in this environment." -ForegroundColor Yellow }
					if ($WhatIf) { exit 0 }
					throw "Catalog fallback failed; opened catalog page for manual download."
				}
				else { throw "No package found on $SourceDrive and catalog fallback failed. Manual download required." }
			}
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

