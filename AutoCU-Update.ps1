[CmdletBinding()]
param(
	[string]$SourceDrive = 'E:',
	[string]$DestinationPath = 'E:\Latest-CU',
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

function Get-LatestKBFromUpdateHistory {
	param(
		[string]$UpdateHistoryUrl = 'https://support.microsoft.com/en-us/servicing/os/windows-11/2025/07/windows-11-version-25h2-update-history',
		[int]$DaysBack = 30
	)

	Write-Host "Checking Windows 11 update-history page: $UpdateHistoryUrl"
	try {
		$resp = Invoke-WebRequest -Uri $UpdateHistoryUrl -UseBasicParsing -Headers @{ 'User-Agent' = 'Mozilla/5.0' }
	}
	catch {
		Write-Host "Failed to fetch update-history page: $($_.Exception.Message)" -ForegroundColor Yellow
		return $null
	}

	$content = $resp.Content

	# Look for KB IDs and nearby dates. We'll try two patterns: 'Month Day, Year' near KB and ISO dates
	$results = @()

	$pattern1 = '([A-Za-z]+\s+\d{1,2},\s+\d{4}).{0,200}?KB(\d{6,7})'
	foreach ($m in [regex]::Matches($content, $pattern1)) {
		$dateText = $m.Groups[1].Value
		$kb = 'KB' + $m.Groups[2].Value
		if ([datetime]::TryParse($dateText, [ref]$d)) {
			$results += [pscustomobject]@{ KB = $kb; Date = $d }
		}
	}

	$pattern2 = 'KB(\d{6,7}).{0,200}?([A-Za-z]+\s+\d{1,2},\s+\d{4})'
	foreach ($m in [regex]::Matches($content, $pattern2)) {
		$kb = 'KB' + $m.Groups[1].Value
		$dateText = $m.Groups[2].Value
		if ([datetime]::TryParse($dateText, [ref]$d)) {
			$results += [pscustomobject]@{ KB = $kb; Date = $d }
		}
	}

	if ($results.Count -eq 0) { return $null }

	$threshold = (Get-Date).AddDays(-$DaysBack)
	$recent = $results | Where-Object { $_.Date -ge $threshold } | Sort-Object Date -Descending
	if ($recent.Count -eq 0) { return $null }

	return $recent[0].KB
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
	dism /Image:$TargetRoot /Add-Package /PackagePath:$PackagePath /IgnoreCheck /LogPath:$logPath
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
			else { throw "Failed to locate any package to install. Manual download required." }
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
	Write-Host "Update-history parsing failed or no recent CU found; using most-recent package on flash drive."
	$packageOnDrive = Find-PackagesOnDrive -RootPath $SourceDrive
	if ($packageOnDrive) {
		Write-Host "Found package on source drive: $packageOnDrive"
		$packagePath = $packageOnDrive
	}
	else {
		throw "No package found on $SourceDrive and update-history lookup provided no candidate. Manual download required."
	}
}

Install-CumulativeUpdate -PackagePath $packagePath -TargetRoot $TargetRoot

Write-Host 'AutoCU update processing completed.'

