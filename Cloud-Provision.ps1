###Update OSD and OSDCLoud Modules if on Wifi
Invoke-OSDCloudPEStartup UpdateModule -Value OSD
Invoke-OSDCloudPEStartup UpdateModule -Value OSDCloud

###Load the newest installed version of each module. UpdateModule installs the new version
###side by side with the one baked into the boot image, but the old one is already loaded in
###this session, so Deploy-OSDCloud would keep using it. OSD first: OSDCloud depends on it.
foreach ($name in 'OSD','OSDCloud') {
    $newest = Get-Module -Name $name -ListAvailable -ErrorAction SilentlyContinue |
        Sort-Object Version -Descending | Select-Object -First 1
    if (-not $newest) { Write-Host "Module $name not found." -ForegroundColor Yellow; continue }

    $loaded = Get-Module -Name $name
    if ($loaded -and $loaded.Version -eq $newest.Version -and $loaded.ModuleBase -eq $newest.ModuleBase) {
        Write-Host "$name $($newest.Version) already loaded." -ForegroundColor Cyan
        continue
    }
    if ($loaded) { Remove-Module -Name $name -Force -ErrorAction SilentlyContinue }
    try {
        # Import by explicit path so the session cannot fall back to an older copy on PSModulePath.
        Import-Module -Name $newest.Path -Force -Global -ErrorAction Stop
        Write-Host "Loaded $name $($newest.Version) from $($newest.ModuleBase)" -ForegroundColor Cyan
    }
    catch {
        Write-Host "Failed to load $name $($newest.Version): $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

###OSDValidation Function
function Test-OSDCloudProvisionValidation {
    Add-Type -AssemblyName System.Windows.Forms
    if (Test-Path "X:\Windows\temp\osdcloud-logs\transcript*.log") {
        $C = Get-Content "X:\Windows\temp\osdcloud-logs\transcript*.log" -Raw
        if ($C -notmatch "Workflow Task execution done"){
            $f = New-Object System.Windows.Forms.Form
            $f.Text = "Log Check Failed"; $f.size = "450,180"; $f.topmost = $true; $f.StartPosition = "CenterScreen"
            $l = New-Object System.Windows.Forms.Label; $l.text = "OSDCloud Did Not Complete Successfully"; $l.Location = "10,10"; $l.size = "400,40"; 
            $b1 = New-Object System.Windows.Forms.Button; 
            $b1.text = "Wipe and Reboot"; $b1.size = "160,35"; $b1.Location = "10,80"; 
            $b1.Add_Click({$f.Close(); Clear-Disk -Number 0 -RemoveData -Confirm:$false;Restart-Computer -Force})
            $b2 = New-Object System.Windows.Forms.Button;
            $b2.text = "Restart OSDCloud"; $b2.size = "160,35"; $b2.Location = "200,80";
            $b2.Add_Click({$f.Close(); Start-Process powershell.exe -ArgumentList "-File X:\Windows\Provision.ps1 -NoLogo -NoProfile -ExecutionPolicy Bypass -WindowStyle Normal"})
            $f.Controls.Add($l); $f.Controls.Add($b1); $f.Controls.Add($b2); $f.ShowDialog() | Out-Null
        }Else{
        $RestartPC = "True"
        Return $restartPC
        }
    }
}

###Write a file to every folder matching a wildcard pattern (e.g. multiple OSDCloud module versions)
function Copy-JsonToAllPaths {
    param(
        [Parameter(Mandatory=$true)][string]$Content,
        [Parameter(Mandatory=$true)][string]$Pattern   # full file path; wildcards allowed in the folders
    )

    $fileName = Split-Path -Path $Pattern -Leaf
    $dirs     = @(Get-ChildItem -Path (Split-Path -Path $Pattern -Parent) -Directory -ErrorAction SilentlyContinue)

    if ($dirs.Count -eq 0) {
        Write-Host "No folders match $(Split-Path -Path $Pattern -Parent); '$fileName' was not written." -ForegroundColor Yellow
        return
    }

    foreach ($dir in $dirs) {
        $dest = Join-Path $dir.FullName $fileName
        try {
            $Content | Out-File -FilePath $dest -Encoding utf8 -Force -ErrorAction Stop
            Write-Host "File '$fileName' has been created successfully: $dest" -ForegroundColor Cyan
        }
        catch {
            Write-Host "Failed to write ${dest}: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

##########Update the OS JSON for the GUI default settings
If ($env:processor_architecture -eq "ARM64") {
    ##Update OSDCLoudGUI Image Type

# 1. Define the content using a Here-String
$jsonContent = @"
{
  "OperatingSystem": {
    "default": "Windows 11 25H2",
    "values": ["Windows 11 25H2", "Windows 11 24H2", "Windows 11 23H2"]
  },
  "OSActivation": {
    "default": "Volume",
    "values": ["Retail", "Volume"]
  },
  "OSEdition": {
    "default": "Enterprise",
    "values": [
      {
        "Edition": "Home",
        "EditionId": "Core"
      },
      {
        "Edition": "Pro",
        "EditionId": "Professional"
      },
      {
        "Edition": "Enterprise",
        "EditionId": "Enterprise"
      }
    ]
  },
  "OSLanguageCode": {
    "default": "en-us",
    "values": [
      "ar-sa",
      "bg-bg",
      "cs-cz",
      "da-dk",
      "de-de",
      "el-gr",
      "en-gb",
      "en-us",
      "es-es",
      "es-mx",
      "et-ee",
      "fi-fi",
      "fr-ca",
      "fr-fr",
      "he-il",
      "hr-hr",
      "hu-hu",
      "it-it",
      "ja-jp",
      "ko-kr",
      "lt-lt",
      "lv-lv",
      "nb-no",
      "nl-nl",
      "pl-pl",
      "pt-br",
      "pt-pt",
      "ro-ro",
      "ru-ru",
      "sk-sk",
      "sl-si",
      "sr-latn-rs",
      "sv-se",
      "th-th",
      "tr-tr",
      "uk-ua",
      "zh-cn",
      "zh-tw"
    ]
  }
}
"@

# 2. Output the content to the file
# We use -Encoding utf8 to ensure standard JSON compatibility
Copy-JsonToAllPaths -Content $jsonContent -Pattern "X:\Program Files\WindowsPowerShell\Modules\OSDCloud\*\Workflow\Default\os-arm64.json"

##Check Drivers

#if ($null -eq $(Get-OSDCatalogDriverPack).name){iex(irm https://raw.githubusercontent.com/Justin-Swets/OSD/refs/heads/main/Get-SurfaceDriversv4.ps1)}
Deploy-OSDCloud
Test-OSDCloudProvisionValidation

# Last step: define Run-OSDGUI (run it manually when needed).
iex (irm 'https://raw.githubusercontent.com/Justin-Swets/OSD/refs/heads/main/OSD-Test.ps1')

If ($RestartPC -eq "True") {
    Restart-Computer -Force}

}elseif ($env:processor_architecture -eq "AMD64") {



    ##Update OSDCLoudGUI Image Type

# 1. Define the content using a Here-String
$jsonContent = @"

   {
  "OperatingSystem": {
    "default": "Windows 11 25H2",
    "values": ["Windows 11 25H2", "Windows 11 24H2", "Windows 11 23H2"]
  },
  "OSActivation": {
    "default": "Volume",
    "values": ["Retail", "Volume"]
  },
  "OSEdition": {
    "default": "Enterprise",
    "values": [
      {
        "Edition": "Home",
        "EditionId": "Core"
      },
      {
        "Edition": "Home N",
        "EditionId": "CoreN"
      },
      {
        "Edition": "Education",
        "EditionId": "Education"
      },
      {
        "Edition": "Education N",
        "EditionId": "EducationN"
      },
      {
        "Edition": "Pro",
        "EditionId": "Professional"
      },
      {
        "Edition": "Pro N",
        "EditionId": "ProfessionalN"
      },
      {
        "Edition": "Enterprise",
        "EditionId": "Enterprise"
      },
      {
        "Edition": "Enterprise N",
        "EditionId": "EnterpriseN"
      }
    ]
  },
  "OSLanguageCode": {
    "default": "en-us",
    "values": [
      "ar-sa",
      "bg-bg",
      "cs-cz",
      "da-dk",
      "de-de",
      "el-gr",
      "en-gb",
      "en-us",
      "es-es",
      "es-mx",
      "et-ee",
      "fi-fi",
      "fr-ca",
      "fr-fr",
      "he-il",
      "hr-hr",
      "hu-hu",
      "it-it",
      "ja-jp",
      "ko-kr",
      "lt-lt",
      "lv-lv",
      "nb-no",
      "nl-nl",
      "pl-pl",
      "pt-br",
      "pt-pt",
      "ro-ro",
      "ru-ru",
      "sk-sk",
      "sl-si",
      "sr-latn-rs",
      "sv-se",
      "th-th",
      "tr-tr",
      "uk-ua",
      "zh-cn",
      "zh-tw"
    ]
  }
}

"@

# 2. Output the content to the file
# We use -Encoding utf8 to ensure standard JSON compatibility
Copy-JsonToAllPaths -Content $jsonContent -Pattern "X:\Program Files\WindowsPowerShell\Modules\OSDCloud\*\Workflow\Default\os-amd64.json"

##Check Drivers

#if ($null -eq $(Get-OSDCatalogDriverPack).name){iex(irm https://raw.githubusercontent.com/Justin-Swets/OSD/refs/heads/main/Get-SurfaceDriversv4.ps1)}
Deploy-OSDCloud

If ((get-ciminstance -Class "Win32_ComputerSystem").Model -like "*Surface Laptop for Business*") {
   Start-Process powershell.exe -ArgumentList @(
    "-NoProfile", 
    "-ExecutionPolicy Bypass", 
    "-WindowStyle Normal", 
    "-Command `"iex (irm 'https://raw.githubusercontent.com/Justin-Swets/OSD/refs/heads/main/AutoCU-Update.ps1')`""
    ) -wait -windowstyle normal
# Do something
}

Test-OSDCloudProvisionValidation

# Last step: define Run-OSDGUI (run it manually when needed).
iex (irm 'https://raw.githubusercontent.com/Justin-Swets/OSD/refs/heads/main/OSD-Test.ps1')

If ($RestartPC -eq "True") {
    Restart-Computer -Force}

}

