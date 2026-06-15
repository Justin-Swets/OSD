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
$jsonContent | Out-File -FilePath "X:\Program Files\WindowsPowerShell\Modules\OSDCloud\*\Workflow\Default\os-arm64.json" -Encoding utf8 -Force

Write-Host "File 'os-arm64.json' has been created successfully." -ForegroundColor Cyan

##Check Drivers

#if ($null -eq $(Get-OSDCatalogDriverPack).name){iex(irm https://raw.githubusercontent.com/Justin-Swets/OSD/refs/heads/main/Get-SurfaceDriversv4.ps1)}
Deploy-OSDCloud

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
$jsonContent | Out-File -FilePath "X:\Program Files\WindowsPowerShell\Modules\OSDCloud\*\Workflow\Default\os-amd64.json" -Encoding utf8 -Force

Write-Host "File 'os-amd64.json' has been created successfully." -ForegroundColor Cyan

##Check Drivers

#if ($null -eq $(Get-OSDCatalogDriverPack).name){iex(irm https://raw.githubusercontent.com/Justin-Swets/OSD/refs/heads/main/Get-SurfaceDriversv4.ps1)}
Deploy-OSDCloud

}

####Validation Check
Add-Type -AssemblyName System.Windows.Forms
$C= Get-Content "X:\Windows\temp\osdcloud-logs\transcript*.log"
if ($c -notmatch "Stop Logs"){

    $f = New-Object System.Windows.Forms.Form
    $f.Text = "Log Check Failed"; $f.size = "350,150"; $f.topmost = $true; $f.StartPosition = "CenterScreen"
    $l = New-Object System.Windows.Forms.Label; $l.text = "OSDCloud Did Not Complete Successfully"; $l.Location = "10,10"; $l.size = "320,40"; 
    $b1 = New-Object System.Windows.Forms.Button; 
    $b1.text = "Wipe and Reboot"; $b1.size = "130,35"; $b1.Location = "125,70"; 
    $b1.Add_Click({$f.Close(); Clear-Disk -Number 0 -RemoveData -Confirm:$false;Restart-Computer -Force})
    $b2 = New-Object System.Windows.Forms.Button;
    $b2.text = "Restart OSDCloud"; $b2.size = "150,35"; $b2.Location = "170,70";
    $b2.Add_Click({$f.Close(); Start-Process powershell.exe -ArgumentList "-File X:\Windows\Provision.ps1" -NoLogo -NoProfile -ExecutionPolicy Bypass -WindowStyle Normal -NoExit})
    $f.Controls.Add($l); $f.Controls.Add($b1); $f.Controls.Add($b2); $f.ShowDialog() | Out-Null

}Else {Restart-Computer -Force}
