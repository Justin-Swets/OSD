<#
.SYNOPSIS
    Collects OSDCloud, driver, setup and event logs for provisioning troubleshooting and
    saves them (zipped) to the flash drive.

.DESCRIPTION
    Works in both environments:

    WinPE        OSDCloud logs from X:, plus the deployed (offline) Windows on C: - its
                 Panther/DISM/CBS/SetupAPI logs and its .evtx files copied directly.
    Full Windows Live event logs exported with wevtutil (needs elevation for most channels),
                 Intune Management Extension logs, and an MdmDiagnosticsTool Autopilot cab.

    Collected:
      OSDCloud   X:\Windows\Temp\osdcloud-logs, X:\OSDCloud\Logs, <Windows>\Temp\osdcloud-logs,
                 C:\OSDCloud\Logs, AutoCU-install.log
      Drivers    SetupAPI dev/app/offline logs, DISM log, driver inventory, devices with problems
      Setup      Panther (setupact/setuperr/unattend), CBS log
      Events     System, Application, Setup, and the Autopilot / MDM enrollment / AAD /
                 device registration / PnP / Shell-Core (ESP) channels
      System     computer, BIOS, OS, disks, volumes, ipconfig

    Output: <flash drive>\DiagLogs\<Serial>_<yyyyMMdd-HHmmss>.zip (the first removable or
    USB volume other than X:). Falls back to C:\DiagLogs when no flash drive is found.
    Anything that could not be collected is listed in _Summary.txt instead of failing the run.

.EXAMPLE
    iex (irm 'https://raw.githubusercontent.com/Justin-Swets/OSD/refs/heads/main/Get-Diagnosticlogs.ps1')

.EXAMPLE
    .\Get-Diagnosticlogs.ps1 -Destination E:\Logs -NoZip
#>
[CmdletBinding()]
param(
    [string]$Destination,
    [string]$WindowsRoot,
    [switch]$NoZip
)

$ErrorActionPreference = 'Continue'
$Script:Summary = New-Object System.Collections.Generic.List[string]

function Add-Summary {
    param([string]$Status, [string]$Text)
    $line = '{0,-7} {1}' -f $Status, $Text
    $Script:Summary.Add($line)
    $color = switch ($Status) { 'OK' { 'Gray' } 'SKIP' { 'DarkGray' } default { 'Yellow' } }
    Write-Host "  $line" -ForegroundColor $color
}

function Test-WinPE {
    return (Test-Path -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\MiniNT')
}

function Get-FlashDriveRoot {
    <# First removable volume, else first USB-bus volume. Never X: (the WinPE RAM disk). #>
    $removable = Get-CimInstance -ClassName Win32_LogicalDisk -Filter 'DriveType=2' -ErrorAction SilentlyContinue |
        Where-Object { $_.DeviceID -ne 'X:' -and $_.Size -gt 0 } |
        Select-Object -First 1 -ExpandProperty DeviceID
    if ($removable) { return "$removable\" }

    $usb = Get-Disk -ErrorAction SilentlyContinue | Where-Object { $_.BusType -eq 'USB' } |
        Get-Partition -ErrorAction SilentlyContinue |
        Where-Object { $_.DriveLetter -and $_.DriveLetter -ne 'X' } |
        Select-Object -First 1 -ExpandProperty DriveLetter
    if ($usb) { return "${usb}:\" }

    return $null
}

function Find-WindowsRoot {
    <# The Windows folder to collect from: the running OS, or in WinPE the deployed offline one. #>
    if ($WindowsRoot) { return $WindowsRoot.TrimEnd('\') }
    if (-not (Test-WinPE)) { return $env:SystemRoot }

    if (Test-Path -LiteralPath 'C:\Windows\System32\config\SYSTEM') { return 'C:\Windows' }
    foreach ($d in (Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Name -ne 'X' })) {
        $candidate = "$($d.Name):\Windows"
        if (Test-Path -LiteralPath "$candidate\System32\config\SYSTEM") { return $candidate }
    }
    return $null
}

function Copy-LogItem {
    <# Copies a file or folder (wildcards allowed) into the collection; missing sources are skipped. #>
    param([string]$Source, [string]$DestDir)

    $items = @(Get-Item -Path $Source -Force -ErrorAction SilentlyContinue)
    if ($items.Count -eq 0) { Add-Summary 'SKIP' "$Source (not found)"; return }

    New-Item -ItemType Directory -Path $DestDir -Force | Out-Null
    foreach ($item in $items) {
        if (-not $item.PSIsContainer) {
            try {
                Copy-Item -LiteralPath $item.FullName -Destination $DestDir -Force -ErrorAction Stop
                Add-Summary 'OK' $item.FullName
            }
            catch { Add-Summary 'FAILED' "$($item.FullName): $($_.Exception.Message)" }
            continue
        }

        # Folders are copied file by file: one locked or protected file (e.g. Panther\diagerr.xml
        # on a live OS) must not abort the rest of the folder.
        $base   = Join-Path $DestDir $item.Name
        $copied = 0
        $failed = New-Object System.Collections.Generic.List[string]
        foreach ($file in (Get-ChildItem -LiteralPath $item.FullName -Recurse -File -Force -ErrorAction SilentlyContinue)) {
            $target = Join-Path $base $file.FullName.Substring($item.FullName.Length).TrimStart('\')
            try {
                New-Item -ItemType Directory -Path (Split-Path -Path $target -Parent) -Force | Out-Null
                Copy-Item -LiteralPath $file.FullName -Destination $target -Force -ErrorAction Stop
                $copied++
            }
            catch { $failed.Add($file.Name) }
        }
        if ($failed.Count -eq 0) { Add-Summary 'OK' "$($item.FullName) ($copied files)" }
        else { Add-Summary 'PARTIAL' "$($item.FullName) ($copied copied; skipped: $($failed -join ', '))" }
    }
}

function Save-CommandOutput {
    <# Runs a scriptblock and saves its text output; a failure is recorded, not thrown. #>
    param([string]$Name, [string]$DestDir, [scriptblock]$Command)

    New-Item -ItemType Directory -Path $DestDir -Force | Out-Null
    $file = Join-Path $DestDir $Name
    try {
        & $Command 2>&1 | Out-File -FilePath $file -Encoding utf8 -Width 400
        Add-Summary 'OK' $Name
    }
    catch {
        Add-Summary 'FAILED' "${Name}: $($_.Exception.Message)"
    }
}

# Event channels relevant to provisioning. Offline .evtx names replace '/' with '%4'.
$Script:EventChannels = @(
    'System'
    'Application'
    'Setup'
    'Microsoft-Windows-ModernDeployment-Diagnostics-Provider/Autopilot'
    'Microsoft-Windows-ModernDeployment-Diagnostics-Provider/ManagementService'
    'Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Admin'
    'Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Operational'
    'Microsoft-Windows-AAD/Operational'
    'Microsoft-Windows-User Device Registration/Admin'
    'Microsoft-Windows-Shell-Core/Operational'
    'Microsoft-Windows-Kernel-PnP/Configuration'
    'Microsoft-Windows-DriverFrameworks-UserMode/Operational'
    'Microsoft-Windows-Provisioning-Diagnostics-Provider/Admin'
)

#==============================  Main  ==============================

$inWinPE = Test-WinPE
$winDir  = Find-WindowsRoot
$serial  = (Get-CimInstance -ClassName Win32_BIOS -ErrorAction SilentlyContinue).SerialNumber
$serial  = if ($serial) { ($serial.Trim() -replace '[\\/:*?"<>|\s]', '_') } else { 'UnknownSerial' }
$stamp   = Get-Date -Format 'yyyyMMdd-HHmmss'

if (-not $Destination) {
    $flash = Get-FlashDriveRoot
    $Destination = if ($flash) { Join-Path $flash 'DiagLogs' } else { 'C:\DiagLogs' }
    if (-not $flash) { Write-Host "No flash drive found; saving to $Destination" -ForegroundColor Yellow }
}

$root = Join-Path $Destination "${serial}_$stamp"
New-Item -ItemType Directory -Path $root -Force | Out-Null

Write-Host "Environment : $(if ($inWinPE) { 'WinPE' } else { 'Full Windows' })" -ForegroundColor Cyan
Write-Host "Windows root: $(if ($winDir) { $winDir } else { 'not found' })" -ForegroundColor Cyan
Write-Host "Output      : $root" -ForegroundColor Cyan

#--- OSDCloud
Write-Host "`nOSDCloud logs" -ForegroundColor Cyan
$osd = Join-Path $root 'OSDCloud'
Copy-LogItem 'X:\Windows\Temp\osdcloud-logs' $osd
Copy-LogItem 'X:\OSDCloud\Logs'              $osd
Copy-LogItem 'C:\OSDCloud\Logs'              $osd
if ($winDir) {
    Copy-LogItem "$winDir\Temp\osdcloud-logs"       $osd
    Copy-LogItem "$winDir\Logs\AutoCU-install.log"  $osd
    Copy-LogItem "$winDir\Logs\AutoCU-Update_*.log" $osd
    Copy-LogItem "$winDir\Logs\OSDCloud*"           $osd
}
# AutoCU-Update falls back to %TEMP% when neither <Windows>\Logs nor the WinPE log folder exists.
Copy-LogItem "$env:TEMP\AutoCU-Update_*.log" $osd

#--- Drivers
Write-Host "`nDriver logs" -ForegroundColor Cyan
$drv = Join-Path $root 'Drivers'
if ($winDir) {
    Copy-LogItem "$winDir\INF\setupapi.dev.log"     $drv
    Copy-LogItem "$winDir\INF\setupapi.app.log"     $drv
    Copy-LogItem "$winDir\INF\setupapi.offline.log" $drv
    Copy-LogItem "$winDir\Logs\DISM\dism.log"       $drv
}
if ($inWinPE) {
    Copy-LogItem 'X:\Windows\INF\setupapi.dev.log' (Join-Path $drv 'WinPE')
    Copy-LogItem 'X:\Windows\Logs\DISM\dism.log'   (Join-Path $drv 'WinPE')
}

# Devices with a problem code (e.g. 28 = no driver). Win32_PnPEntity works in WinPE too.
Save-CommandOutput 'ProblemDevices.txt' $drv {
    Get-CimInstance -ClassName Win32_PnPEntity |
        Where-Object { $_.ConfigManagerErrorCode -ne 0 } |
        Select-Object Name, PNPClass, ConfigManagerErrorCode, DeviceID |
        Format-Table -AutoSize | Out-String -Width 400
}

if ($inWinPE -and $winDir) {
    $image = Split-Path -Path $winDir -Qualifier
    Save-CommandOutput 'InstalledDrivers-Offline.txt' $drv { dism /English /Image:"$image\" /Get-Drivers /Format:Table }
}
elseif (-not $inWinPE) {
    Save-CommandOutput 'InstalledDrivers.txt' $drv { pnputil /enum-drivers }
}

#--- Setup / servicing
Write-Host "`nSetup logs" -ForegroundColor Cyan
$setup = Join-Path $root 'Setup'
if ($winDir) {
    Copy-LogItem "$winDir\Panther"           (Join-Path $setup 'Panther')
    Copy-LogItem "$winDir\Logs\CBS\CBS.log"  $setup
    Copy-LogItem "$winDir\Setup\Scripts"     (Join-Path $setup 'SetupScripts')
}

#--- Event logs
Write-Host "`nEvent logs" -ForegroundColor Cyan
$evt = Join-Path $root 'EventLogs'
New-Item -ItemType Directory -Path $evt -Force | Out-Null

if ($inWinPE) {
    # Offline image: the .evtx files are not in use, so they can be copied directly.
    if ($winDir) {
        foreach ($channel in $Script:EventChannels) {
            Copy-LogItem "$winDir\System32\winevt\Logs\$($channel -replace '/', '%4').evtx" $evt
        }
    }
}
else {
    # Live OS: the .evtx files are locked, so export each channel.
    foreach ($channel in $Script:EventChannels) {
        $file = Join-Path $evt "$($channel -replace '[/ ]', '_').evtx"
        $out  = wevtutil epl "$channel" "$file" /ow:true 2>&1
        if ($LASTEXITCODE -eq 0) { Add-Summary 'OK' $channel }
        else                     { Add-Summary 'FAILED' "${channel}: $(($out | Out-String).Trim())" }
    }

    Write-Host "`nIntune / Autopilot" -ForegroundColor Cyan
    $mdm = Join-Path $root 'Intune'
    Copy-LogItem "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs" $mdm

    # Bundles enrollment, provisioning and Autopilot diagnostics (registry, ETL, ESP state).
    $mdmTool = Join-Path $env:SystemRoot 'System32\MdmDiagnosticsTool.exe'
    if (Test-Path -LiteralPath $mdmTool) {
        New-Item -ItemType Directory -Path $mdm -Force | Out-Null
        $cab = Join-Path $mdm 'MDMDiagnostics.cab'
        & $mdmTool -area 'DeviceEnrollment;DeviceProvisioning;Autopilot' -cab $cab 2>&1 | Out-Null
        if (Test-Path -LiteralPath $cab) { Add-Summary 'OK' 'MdmDiagnosticsTool cab' }
        else                             { Add-Summary 'FAILED' 'MdmDiagnosticsTool produced no cab (needs elevation)' }
    }
}

#--- System info
Write-Host "`nSystem info" -ForegroundColor Cyan
$sys = Join-Path $root 'System'
Save-CommandOutput 'ComputerSystem.txt' $sys {
    Get-CimInstance Win32_ComputerSystem | Format-List Manufacturer, Model, SystemFamily, SystemSKUNumber, TotalPhysicalMemory
    Get-CimInstance Win32_BIOS           | Format-List Manufacturer, SMBIOSBIOSVersion, SerialNumber, ReleaseDate
    Get-CimInstance Win32_OperatingSystem | Format-List Caption, Version, BuildNumber, OSArchitecture
    "PROCESSOR_ARCHITECTURE = $env:PROCESSOR_ARCHITECTURE"
}
Save-CommandOutput 'Disks.txt' $sys {
    Get-Disk -ErrorAction SilentlyContinue | Format-Table -AutoSize Number, FriendlyName, BusType, PartitionStyle, Size, OperationalStatus
    Get-Volume -ErrorAction SilentlyContinue | Format-Table -AutoSize DriveLetter, FileSystemLabel, FileSystem, DriveType, Size, SizeRemaining
}
Save-CommandOutput 'ipconfig.txt' $sys { ipconfig /all }

if ($winDir -and $inWinPE) {
    # Build/UBR of the deployed image, read from its offline SOFTWARE hive.
    Save-CommandOutput 'OfflineImageVersion.txt' $sys {
        $hive = "$winDir\System32\config\SOFTWARE"
        reg load HKLM\DiagOffline $hive | Out-Null
        try {
            Get-ItemProperty 'HKLM:\DiagOffline\Microsoft\Windows NT\CurrentVersion' |
                Format-List ProductName, DisplayVersion, CurrentBuild, UBR, EditionID
        }
        finally {
            [gc]::Collect(); [gc]::WaitForPendingFinalizers()
            reg unload HKLM\DiagOffline | Out-Null
        }
    }
}

#--- Summary and zip
$Script:Summary | Out-File -FilePath (Join-Path $root '_Summary.txt') -Encoding utf8

$failed = @($Script:Summary | Where-Object { $_ -like 'FAILED*' }).Count
$result = $root
if (-not $NoZip) {
    $zip = "$root.zip"
    try {
        Compress-Archive -Path (Join-Path $root '*') -DestinationPath $zip -Force -ErrorAction Stop
        Remove-Item -LiteralPath $root -Recurse -Force
        $result = $zip
    }
    catch {
        Write-Host "Could not zip ($($_.Exception.Message)); logs left in the folder." -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "Diagnostic logs saved: $result" -ForegroundColor Green
if ($failed -gt 0) { Write-Host "$failed item(s) could not be collected; see _Summary.txt." -ForegroundColor Yellow }
