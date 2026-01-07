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
    $Links = $Response.Links | Where-Object { $_.href -like "*details.aspx?id=