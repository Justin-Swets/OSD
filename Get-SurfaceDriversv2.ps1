Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- CONFIGURATION ---
$WildcardXmlPath = "X:\Program Files\WindowsPowerShell\Modules\OSD\*\cache\driverpack-catalogs\build-driverpacks.xml"
$SurfaceDocsUrl = "https://support.microsoft.com/en-us/surface/download-drivers-and-firmware-for-surface-09bb2e09-2a4b-cb69-0951-078a7739e120"

# --- PATH RESOLUTION ---
try {
    $XmlPath = Resolve-Path -Path $WildcardXmlPath -ErrorAction Stop | Select-Object -ExpandProperty Path -First 1
    Write-Host "[0/4] Resolved XML Path: $XmlPath" -ForegroundColor Gray
} catch {
    Write-Error "XML Path could not be resolved. Ensure the OSD module is installed."
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
    $Form.MaximizeBox = $false

    $Label = New-Object System.Windows.Forms.Label
    $Label.Text = "Multiple potential matches found. Select the correct MSI package:"
    $Label.Location = New-Object System.Drawing.Point(10,10)
    $Label.AutoSize = $true
    $Form.Controls.Add($Label)

    $ListBox = New-Object System.Windows.Forms.ListBox
    $ListBox.Location = New-Object System.Drawing.Point(10,40)
    $ListBox.Size = New-Object System.Drawing.Size(660,300)
    foreach ($Opt in $Options) { 
        [void]$ListBox.Items.Add("Score: $($Opt.MatchScore) | Model: $($Opt.Model) | File: $($Opt.FileName)") 
    }
    $ListBox.SelectedIndex = 0
    $Form.Controls.Add($ListBox)

    $Button = New-Object System.Windows.Forms.Button
    $Button.Text = "Confirm Selection"
    $Button.Location = New-Object System.Drawing.Point(280,355)
    $Button.Size = New-Object System.Drawing.Size(120,30)
    $Button.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $Form.Controls.Add($Button)

    $Form.AcceptButton = $Button

    if ($Form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $Options[$ListBox.SelectedIndex]
    }
    return $null
}

function Write-NewXmlEntry {
    param ([string]$Path, [string]$ModelName, [string]$NewUrl, [string]$FileName)

    Write-Host "[XML] Wiping $Path and writing fresh driver entry..." -ForegroundColor Yellow
    
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
    # Destructive Write
    if (Test-Path $Path) { Remove-Item $Path -Force }
    Set-Content -Path $Path -Value $XmlContent -Encoding UTF8
    Write-Host "[XML] Successfully written new catalog entry." -ForegroundColor Green
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

    # Identify Local System
    $LocalModelRaw = (Get-CimInstance Win32_ComputerSystem).Model
    $LocalModelNoBus = $LocalModelRaw -replace " for Business", ""
    
    # Architecture Check
    $CPU = (Get-CimInstance Win32_Processor).Name
    $IsSnapdragon = $CPU -match "Snapdragon|SQ1|SQ2|SQ3"
    
    # LINE 116 FIX (Ternary replaced with if/else for PS 5.1 compatibility)
    $LogArch = if ($IsSnapdragon) { "ARM64" } else { "x64" }
    Write-Host "[2/4] Detected: $LocalModelRaw ($LogArch)" -ForegroundColor Cyan

    # --- TIERED MATCHING LOGIC ---
    Write-Host "[3/4] Performing Tiered Analysis..." -ForegroundColor Gray
    $FinalSelection = $null

    # TIER 1: Literal Exact Match
    $FinalSelection = $DriverMap | Where-Object { $_.Model -ieq $LocalModelRaw } | Select-Object -First 1

    # TIER 2: "For Business" Stripped Match
    if (-not $FinalSelection -and $LocalModelRaw -like "*for Business*") {
        Write-Host "      Tier 1 failed. Trying 'For Business' normalized match..." -ForegroundColor DarkGray
        $FinalSelection = $DriverMap | Where-Object { $_.Model -ieq $LocalModelNoBus } | Select-Object -First 1
    }

    # TIER 3: Complex Multi-Factor Matching (CPU, Size, Edition, Filename)
    if (-not $FinalSelection) {
        Write-Host "      Performing complex tiered scoring..." -ForegroundColor DarkGray
        $Keywords = $LocalModelNoBus.ToLower().Split(" ", [System.StringSplitOptions]::RemoveEmptyEntries)

        $Candidates = foreach ($Item in $DriverMap) {
            $Score = 0
            $SearchText = ($Item.Model + " " + $Item.FileName).ToLower()

            # Architecture Weighting
            if ($IsSnapdragon -and $SearchText -match "arm64|snapdragon") { $Score += 50 }
            elseif (-not $IsSnapdragon -and $SearchText -match "intel|x64") { $Score += 20 }

            # Keyword Weighting (Size, Generation, Edition)
            foreach ($k in $Keywords) {
                if ($SearchText.Contains($k)) { $Score += 10 }
            }

            # Update the object with its score
            $Item | Add-Member -MemberType NoteProperty -Name "MatchScore" -Value $Score -Force
            if ($Score -ge 30) { $Item }
        }

        $MatchesFound = $Candidates | Sort-Object MatchScore -Descending

        if ($MatchesFound.Count -eq 1) {
            $FinalSelection = $MatchesFound[0]
        } elseif ($MatchesFound.Count -gt 1) {
            Write-Host "      Multiple candidates identified. Opening selection menu..." -ForegroundColor Yellow
            $FinalSelection = Show-SelectionMenu -Options $MatchesFound
        }
    }

    # --- FINAL WRITE ---
    if ($FinalSelection) {
        Write-Host "`n[4/4] Selection Confirmed: $($FinalSelection.Model)" -ForegroundColor Green
        Write-NewXmlEntry -Path $XmlPath -ModelName $FinalSelection.Model -NewUrl $FinalSelection.URL -FileName $FinalSelection.FileName
        return $FinalSelection.URL
    } else {
        Write-Error "Could not determine or select a matching driver package."
    }

} catch {
    Write-Error "Critical Script Failure: $($_.Exception.Message)"
}