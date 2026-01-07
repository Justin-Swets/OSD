Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- CONFIGURATION ---
$WildcardXmlPath = "X:\Program Files\WindowsPowerShell\Modules\OSD\*\cache\driverpack-catalogs\build-driverpacks.xml"
$SurfaceDocsUrl = "https://support.microsoft.com/en-us/surface/download-drivers-and-firmware-for-surface-09bb2e09-2a4b-cb69-0951-078a7739e120"

# --- PATH RESOLUTION ---
Write-Host "[0/4] Identifying all target XML locations..." -ForegroundColor Gray
try {
    $TargetPaths = Resolve-Path -Path $WildcardXmlPath -ErrorAction Stop | Select-Object -ExpandProperty Path
    Write-Host "      Found $($TargetPaths.Count) target locations." -ForegroundColor DarkGray
} catch {
    Write-Error "Could not resolve any paths matching $WildcardXmlPath"
    return
}

# --- FUNCTIONS ---

function Show-SelectionMenu {
    param([array]$Options)
    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "Surface Driver Selection"
    $Form.Size = New-Object System.Drawing.Size(700,450)
    $Form.StartPosition = "CenterScreen"
    $Form.FormBorderStyle = "FixedDialog"

    $Label = New-Object System.Windows.Forms.Label
    $Label.Text = "Multiple potential matches found. Select the correct MSI package:"
    $Label.Location = New-Object System.Drawing.Point(10,10); $Label.AutoSize = $true
    $Form.Controls.Add($Label)

    $ListBox = New-Object System.Windows.Forms.ListBox
    $ListBox.Location = New-Object System.Drawing.Point(10,40); $ListBox.Size = New-Object System.Drawing.Size(660,300)
    foreach ($Opt in $Options) { [void]$ListBox.Items.Add("Score: $($Opt.MatchScore) | Model: $($Opt.Model) | File: $($Opt.FileName)") }
    $ListBox.SelectedIndex = 0; $Form.Controls.Add($ListBox)

    $Button = New-Object System.Windows.Forms.Button
    $Button.Text = "Confirm Selection"; $Button.Location = New-Object System.Drawing.Point(280,355)
    $Button.Size = New-Object System.Drawing.Size(120,30); $Button.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $Form.Controls.Add($Button); $Form.AcceptButton = $Button

    if ($Form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $Options[$ListBox.SelectedIndex] }
    return $null
}

function Write-NewXmlEntry {
    param ([array]$Paths, [string]$ModelName, [string]$NewUrl, [string]$FileName)
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
      <S N="Model">$ModelName</S>
      <S N="Url">$NewUrl</S>
      <S N="FileName">$FileName</S>
      <S N="OS">Windows 10/11 x64</S>
      <S N="Guid">$Guid</S>
    </MS>
  </Obj>
</Objs>
"@
    foreach ($CurrentPath in $Paths) {
        Write-Host "[XML] Writing fresh driver entry to: $CurrentPath" -ForegroundColor Yellow
        if (Test-Path $CurrentPath) { Remove-Item $CurrentPath -Force }
        Set-Content -Path $CurrentPath -Value $XmlContent -Encoding UTF8
    }
}

# --- MAIN EXECUTION ---

Write-Host "[1/4] Scraping Microsoft Support for Driver Links..." -ForegroundColor Cyan
try {
    $Response = Invoke-WebRequest -Uri $SurfaceDocsUrl -UseBasicParsing
    $Links = $Response.Links | Where-Object { $_.href -like "*details.aspx?id=*" } | Select-Object -ExpandProperty href -Unique

    $DriverMap = @()
    foreach ($Url in $Links) {
        try {
            $Page = Invoke-WebRequest -Uri $Url -UseBasicParsing -TimeoutSec 5
            $Name = if ($Page.Content -match "<title>Download (.*?) Drivers") { ($Matches[1] -replace " Drivers.*", "").Trim() }
            $File = if ($Page.Content -match "([\w\d\-_]+\.msi)") { $Matches[1] }
            if ($Name) { $DriverMap += [PSCustomObject]@{ Model = $Name; URL = $Url; FileName = $File } }
        } catch { continue }
    }

    # Identify Local System & Architecture
    $LocalModelRaw = (Get-CimInstance Win32_ComputerSystem).Model
    $LocalModelNoBus = $LocalModelRaw -replace " for Business", ""
    $CPU = (Get-CimInstance Win32_Processor).Name

    # Logic to determine strict exclusion platform
    $IsSnapdragon = $LocalModelRaw -match "Snapdragon" -or $CPU -match "Snapdragon|SQ1|SQ2|SQ3"
    $IsIntel = $LocalModelRaw -match "Intel" -or $CPU -match "Intel"
    $IsAMD = $LocalModelRaw -match "AMD" -or $CPU -match "Ryzen"

    $LogArch = if($IsSnapdragon){"Snapdragon"} elseif($IsAMD){"AMD"} else {"Intel"}
    Write-Host "[2/4] Detected: $LocalModelRaw ($LogArch)" -ForegroundColor Cyan

    # --- TIERED MATCHING LOGIC ---
    Write-Host "[3/4] Performing Tiered Analysis with Architecture Exclusion..." -ForegroundColor Gray
    
    # Pre-Filter DriverMap based on Strict Architecture Exclusion
    $FilteredMap = $DriverMap | Where-Object {
        $TargetText = ($_.Model + " " + $_.FileName).ToLower()
        if ($IsSnapdragon) { return $TargetText -notmatch "intel|amd|x64" }
        if ($IsIntel) { return $TargetText -notmatch "snapdragon|arm64|amd" }
        if ($IsAMD) { return $TargetText -notmatch "snapdragon|arm64|intel" }
        return $true
    }

    $FinalSelection = $null

    # TIER 1: Literal Exact Match
    $FinalSelection = $FilteredMap | Where-Object { $_.Model -ieq $LocalModelRaw } | Select-Object -First 1

    # TIER 2: "For Business" Stripped Match
    if (-not $FinalSelection -and $LocalModelRaw -like "*for Business*") {
        $FinalSelection = $FilteredMap | Where-Object { $_.Model -ieq $LocalModelNoBus } | Select-Object -First 1
    }

    # TIER 3: Complex Scoring (Size, Generation, Edition)
    if (-not $FinalSelection) {
        $Keywords = $LocalModelNoBus.ToLower().Split(" ", [System.StringSplitOptions]::RemoveEmptyEntries)
        $Candidates = foreach ($Item in $FilteredMap) {
            $Score = 0
            $SearchText = ($Item.Model + " " + $Item.FileName).ToLower()

            # Architecture affinity boost
            if ($IsSnapdragon -and $SearchText -match "arm64|snapdragon") { $Score += 50 }
            elseif ($IsIntel -and $SearchText -match "intel") { $Score += 20 }
            elseif ($IsAMD -and $SearchText -match "amd") { $Score += 20 }

            foreach ($k in $Keywords) { if ($SearchText.Contains($k)) { $Score += 10 } }
            $Item | Add-Member -MemberType NoteProperty -Name "MatchScore" -Value $Score -Force
            if ($Score -ge 30) { $Item }
        }

        $MatchesFound = $Candidates | Sort-Object MatchScore -Descending
        if ($MatchesFound.Count -eq 1) { $FinalSelection = $MatchesFound[0] }
        elseif ($MatchesFound.Count -gt 1) { $FinalSelection = Show-SelectionMenu -Options $MatchesFound }
    }

    # --- FINAL WRITE ---
    if ($FinalSelection) {
        Write-Host "`n[4/4] Selection Confirmed: $($FinalSelection.Model)" -ForegroundColor Green
        Write-NewXmlEntry -Paths $TargetPaths -ModelName $FinalSelection.Model -NewUrl $FinalSelection.URL -FileName $FinalSelection.FileName
    } else {
        Write-Error "No architecture-compliant match could be determined."
    }

} catch {
    Write-Error "Critical Script Failure: $($_.Exception.Message)"
}