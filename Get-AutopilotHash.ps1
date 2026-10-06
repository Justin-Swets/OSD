<#
.SYNOPSIS
    Captures this device's Windows Autopilot hardware hash and saves it to the flash drive
    as <SerialNumber>.csv, ready for import into Intune.

.DESCRIPTION
    Downloads Get-WindowsAutopilotInfo from the PowerShell Gallery to a temp folder (it is
    not installed, so nothing is left behind on the device) and runs it with -OutputFile.

    Requires full Windows - including OOBE (Shift+F10) - in an elevated session, plus
    internet access to the PowerShell Gallery. It does NOT work in WinPE:
    Get-WindowsAutopilotInfo reads the MDM_DevDetail_Ext01 WMI class, which WinPE lacks.

    The flash drive is the first removable (or USB-attached) volume other than X:, unless
    -Destination is given. Files go to <drive>\Autopilot\<SerialNumber>.csv.

.EXAMPLE
    iex (irm 'https://raw.githubusercontent.com/Justin-Swets/OSD/refs/heads/main/Get-AutopilotHash.ps1')

.EXAMPLE
    .\Get-AutopilotHash.ps1 -Destination E:\Hashes -GroupTag Kiosk
#>
[CmdletBinding()]
param(
    [string]$Destination,
    [string]$GroupTag
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function Get-DeviceSerialNumber {
    $serial = (Get-CimInstance -ClassName Win32_BIOS).SerialNumber
    if ($serial) { $serial = $serial.Trim() }

    # Placeholder values some OEMs leave in the BIOS are useless as a file name or for Intune.
    if (-not $serial -or $serial -match '^(To be filled.*|Default string|System Serial Number|0+|None)$') {
        throw "BIOS serial number is missing or a placeholder ('$serial'); cannot name the hash file."
    }
    return $serial
}

function Get-FlashDriveRoot {
    <#
        Removable drives report DriveType 2. Some large USB sticks/SSDs report as fixed
        disks, so USB-bus disks are checked as a fallback. X: (the WinPE RAM disk) is never used.
    #>
    $removable = Get-CimInstance -ClassName Win32_LogicalDisk -Filter 'DriveType=2' |
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

function Save-AutopilotInfoScript {
    <# Downloads Get-WindowsAutopilotInfo.ps1 from the PowerShell Gallery and returns its path. #>
    param([Parameter(Mandatory=$true)][string]$Path)

    # A fresh Windows/OOBE image has no NuGet provider, and PowerShellGet prompts for it.
    if (-not (Get-PackageProvider -ListAvailable -Name NuGet -ErrorAction SilentlyContinue |
              Where-Object { $_.Version -ge [version]'2.8.5.201' })) {
        Write-Host 'Installing the NuGet package provider...'
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser | Out-Null
    }

    Write-Host 'Downloading Get-WindowsAutopilotInfo from the PowerShell Gallery...'
    Save-Script -Name Get-WindowsAutopilotInfo -Repository PSGallery -Path $Path -Force

    $script = Join-Path $Path 'Get-WindowsAutopilotInfo.ps1'
    if (-not (Test-Path -LiteralPath $script)) { throw "Get-WindowsAutopilotInfo.ps1 was not saved to $Path." }
    return $script
}

#==============================  Main  ==============================

if (Test-Path -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\MiniNT') {
    throw ('Running in WinPE. Get-WindowsAutopilotInfo needs the MDM WMI provider, which WinPE does not have. ' +
           'Run this from full Windows or OOBE (Shift+F10) instead.')
}

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).
    IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) { throw 'Run PowerShell as Administrator; the hardware hash cannot be read otherwise.' }

$serial = Get-DeviceSerialNumber
Write-Host "Serial number: $serial" -ForegroundColor Cyan

if (-not $Destination) {
    $flashRoot = Get-FlashDriveRoot
    if (-not $flashRoot) { throw 'No flash drive found. Insert one or pass -Destination.' }
    $Destination = Join-Path $flashRoot 'Autopilot'
}
New-Item -ItemType Directory -Path $Destination -Force | Out-Null

# Characters that are invalid in file names are replaced, so odd serials still save.
$outFile = Join-Path $Destination "$($serial -replace '[\\/:*?"<>|\s]', '_').csv"

$work = Join-Path $env:TEMP "AutopilotHash_$([guid]::NewGuid().ToString('N'))"
New-Item -ItemType Directory -Path $work -Force | Out-Null

try {
    $tool = Save-AutopilotInfoScript -Path $work

    # The default execution policy (Restricted on a fresh install) blocks running the saved script.
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

    $toolArgs = @{ OutputFile = $outFile }
    if ($GroupTag) { $toolArgs['GroupTag'] = $GroupTag }
    & $tool @toolArgs

    if (-not (Test-Path -LiteralPath $outFile)) { throw "Get-WindowsAutopilotInfo did not create $outFile." }
}
finally {
    Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "Autopilot hash saved: $outFile" -ForegroundColor Green
