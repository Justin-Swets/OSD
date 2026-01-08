Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- CONFIGURATION ---
$WildcardXmlPath = "X:\Program Files\WindowsPowerShell\Modules\OSD\*\cache\driverpack-catalogs\build-driverpacks.xml"
$SurfaceDocsUrl = "https://support.microsoft.com/en-us/surface/download-drivers-and-firmware-for-surface-09bb2e09-2a4b-cb69-0951-078a7739e120"
$ExternalFolderName = "SurfaceDriverXML"
$ExternalFileName = "build-driverpacks.xml"

# --- FUNCTIONS ---

function Get-ExternalXmlPath {
    # Preference: E:\ if OSDCloud exists, else D:\
    $Drive = if (Test-Path "E:\OSDCloud") { "E:" } elseif (Test-Path "D:\") { "D:" } else { $null }
    if ($Drive) {
        $FullPath = Join-Path -Path $Drive -ChildPath "$ExternalFolderName\$ExternalFileName"
        return $FullPath
    }
    return $null
}

function Get-OfflineMatch {
    param([string]$CurrentSku)
    $OfflinePath = Get-ExternalXmlPath
    if ($OfflinePath -and (Test-Path $OfflinePath)) {
        Write-Host "[OFFLINE] Checking external cache: $OfflinePath" -ForegroundColor Gray
        try {
            [xml]$xml = Get-Content $OfflinePath
            # Look for an entry where the 'Name' property matches our SystemSku
            $Match = $xml.Objs.Obj.MS.S | Where-Object { $_.N -eq "Name" -and $_.'#text' -eq $CurrentSku }
            if ($Match) {
                $ParentMS = $Match.ParentNode
                return [PSCustomObject]@{
                    Model    = ($ParentMS.S | Where-Object { $_.N -eq "Model" }).'#text'
                    URL      = ($ParentMS.S | Where-Object { $_.N -eq "Url" }).'#text'
                    FileName = ($ParentMS.S | Where-Object { $_.N -eq "FileName" }).'#text'
                    Source   = "OfflineCache"
                }
            }
        } catch { return $null }
    }
    return $null
}

function Write-NewXmlEntry {
    param ([array]$Paths, [string]$ModelName, [string]$NewUrl, [string]$FileName, [hashtable]$SysInfo)
    
    $Guid = [guid]::NewGuid().ToString()
    $XmlContent = @"
<Objs Version="1.1.0.1" xmlns="http://schemas.microsoft.com/powershell/2004/04">
  <Obj RefId="0">
    <TN RefId="0">
      <T>Selected.System.Management.Automation.PSCustomObject</T>
      <T>System.Management.Automation.PSCustomObject</T>
      <T>System.Object</T>
    </TN>
    <MS>
      <S N="Manufacturer">Microsoft</S>
      <S N="Product">$($SysInfo.Product)</S>
      <S N="Model">$ModelName</S>
      <S N="Name">$($SysInfo.Name)</S>
      <S N="Url">$NewUrl</S>
      <S N="FileName">$FileName</S>
      <S N="OS">Windows 11 x64</S>
      <S N="OSReleaseID">$($SysInfo.OSReleaseID)</S>
      <S N="OSArchitecture">$($SysInfo.OSArchitecture)</S>
      <S N="Guid">$Guid</S>
    </MS>
  </Obj>
</Objs>
"@
    # 1. Write to X: drive locations (Wildcard paths)
    foreach ($CurrentPath in $Paths) {
        Write-Host "[XML] Writing to OSD Cache: $CurrentPath" -ForegroundColor Yellow
        if (Test-Path $CurrentPath) { Remove-Item $CurrentPath -Force }
        Set-Content -Path $CurrentPath -Value $XmlContent -Encoding UTF8
    }

    # 2. Write to External Media (D: or E:)
    $ExternalPath = Get-ExternalXmlPath
    if ($ExternalPath) {
        $Dir = Split-Path $ExternalPath
        if (-not (Test-Path $Dir)) { New-Item -ItemType Directory -Path $Dir -Force | Out-Null }
        Write-Host "[XML] Syncing to External Media: $ExternalPath" -ForegroundColor Cyan
        Set-Content -Path $ExternalPath -Value $XmlContent -Encoding UTF8
    }
}

function Show-SelectionMenu {
    param([array]$Options)
    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "Surface Driver Selection"
    $Form.Size = New-Object System.Drawing.Size(1000,500)
    $Form.StartPosition = "CenterScreen"
    $Form.FormBorderStyle = "FixedDialog"

    $ListBox = New-Object System.Windows.Forms.ListBox
    $ListBox.Location = New-Object System.Drawing.Point(10,40); $ListBox.Size = New-Object System.Drawing.Size(960,320)
    $ListBox.Font = New-Object System.Drawing.Font("Consolas", 9); $ListBox.HorizontalScrollbar = $true

    $MaxWidth = 0
    $Graphics = $ListBox.CreateGraphics()
    foreach ($Opt in $Options) { 
        $ItemString = "Score: $($Opt.MatchScore.ToString().PadRight(3)) | Date: $($Opt.Date.PadRight(12)) | Model: $($Opt.Model.PadRight(35)) | File: $($Opt.FileName)"
        [void]$ListBox.Items.Add($ItemString)
        $TextSize = $Graphics.MeasureString($ItemString, $ListBox.Font)
        if ($TextSize.Width -gt $MaxWidth) { $MaxWidth = $TextSize.Width }
    }
    $ListBox.HorizontalExtent = [int]$MaxWidth + 50
    $ListBox.SelectedIndex = 0; $Form.Controls.Add($ListBox)

    $Button = New-Object System.Windows.Forms.Button
    $Button.Text = "Confirm"; $Button.Location = New-Object System.Drawing.Point(430,380); $Button.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $Form.Controls.Add($Button); $Form.AcceptButton = $Button

    if ($Form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $Options[$ListBox.SelectedIndex] }
    return $null
}

# --- MAIN EXECUTION ---

# 0. Resolve X: Paths
try {
    $TargetPaths = Resolve-Path -Path $WildcardXmlPath -ErrorAction Stop | Select-Object -ExpandProperty Path
} catch {
    Write-Error "OSD Wildcard path not found." ; return
}

# 1. Gather System Info
$ComputerSystem = Get-CimInstance Win32_ComputerSystem
$Baseboard = Get-CimInstance Win32_Baseboard
$CPU = (Get-CimInstance Win32_Processor).Name
$IsSnapdragon = ($ComputerSystem.Model -match "Snapdragon") -or ($CPU -match "Snapdragon|SQ1|SQ2|SQ3")
$ArchitectureString = if ($IsSnapdragon) { "arm64" } else { "amd64" }
$ReleaseId = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name "DisplayVersion" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty DisplayVersion

$SystemInfo = @{
    Product = $ComputerSystem.SystemSkuNumber; Name = $Baseboard.Product
    OSReleaseID = $ReleaseId; OSArchitecture = $ArchitectureString
}

Write-Host "[1/4] Checking for Offline Cache match..." -ForegroundColor Cyan
$OfflineResult = Get-OfflineMatch -CurrentSku $SystemInfo.Name

if ($OfflineResult) {
    Write-Host "[!] MATCH FOUND IN OFFLINE CACHE: $($OfflineResult.Model)" -ForegroundColor Green
    Write-NewXmlEntry -Paths $TargetPaths -ModelName $OfflineResult.Model -NewUrl $OfflineResult.URL -FileName $OfflineResult.FileName -SysInfo $SystemInfo
    return
}

# 2. Scrape if Offline match fails
Write-Host "[2/4] No offline match. Scraping Microsoft Support..." -ForegroundColor Cyan
try {
    $MainPage = Invoke-WebRequest -Uri $SurfaceDocsUrl -UseBasicParsing
    $Rows = [regex]::Matches($MainPage.Content, '(?s)<tr>.*?</tr>')
    $DriverMap = @()
    foreach ($Row in $Rows) {
        if ($Row.Value -match 'href="(https://www\.microsoft\.com/[^"]+?details\.aspx\?id=\d+)"') {
            $LandingUrl = $Matches[1]
            $RowDate = if ($Row.Value -match '(\d{1,2}/\d{1,2}/\d{4})|([A-Z][a-z]+ \d{1,2}, \d{4})') { $Matches[0] } else { "Unknown" }
            try {
                $DetailPage = Invoke-WebRequest -Uri $LandingUrl -UseBasicParsing -TimeoutSec 10
                $Name = if ($DetailPage.Content -match "<title>Download (.*?) Drivers") { ($Matches[1] -replace " Drivers.*", "").Trim() }
                $DirectDownloadUrl = if ($DetailPage.Content -match '(https://download\.microsoft\.com/[^"]+?\.msi)') { $Matches[1] } else { "" }
                if ($Name) {
                    $DriverMap += [PSCustomObject]@{
                        Model = $Name; URL = if ($DirectDownloadUrl) { $DirectDownloadUrl } else { $LandingUrl }
                        FileName = if ($DirectDownloadUrl) { Split-Path -Leaf $DirectDownloadUrl } else { "" }; Date = $RowDate
                    }
                }
            } catch { continue }
        }
    }

    # 3. Filtering & Selection
    $LocalModelRaw = $ComputerSystem.Model
    $LocalModelNoBus = $LocalModelRaw -replace " for Business", ""
    $IsAMD = ($LocalModelRaw -match "AMD") -or ($CPU -match "Ryzen")
    
    $FilteredMap = $DriverMap | Where-Object {
        $TargetText = ($_.Model + " " + $_.FileName).ToLower()
        if ($IsSnapdragon) { return $TargetText -notmatch "intel|amd|x64" }
        if ($LocalModelRaw -match "Intel") { return $TargetText -notmatch "snapdragon|arm64|amd" }
        if ($IsAMD) { return $TargetText -notmatch "snapdragon|arm64|intel" }
        return $true
    }

    $FinalSelection = $FilteredMap | Where-Object { $_.Model -ieq $LocalModelRaw -or $_.Model -ieq $LocalModelNoBus } | Select-Object -First 1

    if (-not $FinalSelection) {
        Write-Host "[3/4] Performing Tiered Scoring..." -ForegroundColor Gray
        $Keywords = $LocalModelNoBus.ToLower().Split(" ", [System.StringSplitOptions]::RemoveEmptyEntries)
        $Candidates = foreach ($Item in $FilteredMap) {
            $Score = 0
            $SearchText = ($Item.Model + " " + $Item.FileName).ToLower()
            if ($IsSnapdragon -and ($SearchText -match "arm64|snapdragon")) { $Score += 50 }
            elseif ($SearchText -match "intel") { $Score += 20 }
            foreach ($k in $Keywords) { if ($SearchText.Contains($k)) { $Score += 10 } }
            $Item | Add-Member -MemberType NoteProperty -Name "MatchScore" -Value $Score -Force
            if ($Score -ge 30) { $Item }
        }
        $MatchesFound = $Candidates | Sort-Object MatchScore -Descending
        if ($MatchesFound.Count -gt 0) { $FinalSelection = Show-SelectionMenu -Options $MatchesFound }
    }

    # 4. Final Write (X: and External)
    if ($FinalSelection) {
        Write-Host "`n[4/4] Finalizing: $($FinalSelection.Model)" -ForegroundColor Green
        Write-NewXmlEntry -Paths $TargetPaths -ModelName $FinalSelection.Model -NewUrl $FinalSelection.URL -FileName $FinalSelection.FileName -SysInfo $SystemInfo
    }
} catch { Write-Error "Failure: $($_.Exception.Message)" }