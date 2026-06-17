###OSDValidation Function
function Test-OSDCloudProvisionValidation {
    Add-Type -AssemblyName System.Windows.Forms

    $logContent = Get-Content "X:\Windows\temp\osdcloud-logs\transcript*.log"

    if ($logContent -notmatch "Workflow Task execution done") {
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Log Check Failed"
        $form.Size = "450,180"
        $form.TopMost = $true
        $form.StartPosition = "CenterScreen"

        $label = New-Object System.Windows.Forms.Label
        $label.Text = "OSDCloud Did Not Complete Successfully"
        $label.Location = "10,10"
        $label.Size = "400,40"

        $buttonReboot = New-Object System.Windows.Forms.Button
        $buttonReboot.Text = "Wipe and Reboot"
        $buttonReboot.Size = "160,35"
        $buttonReboot.Location = "10,80"
        $buttonReboot.Add_Click({
            $form.Close()
            Clear-Disk -Number 0 -RemoveData -Confirm:$false
            Restart-Computer -Force
        })

        $buttonRestart = New-Object System.Windows.Forms.Button
        $buttonRestart.Text = "Restart OSDCloud"
        $buttonRestart.Size = "160,35"
        $buttonRestart.Location = "200,80"
        $buttonRestart.Add_Click({
            $form.Close()
            Start-Process powershell.exe -ArgumentList "-File X:\Windows\Provision.ps1 -NoLogo -NoProfile -ExecutionPolicy Bypass -WindowStyle Normal"
        })

        $form.Controls.Add($label)
        $form.Controls.Add($buttonReboot)
        $form.Controls.Add($buttonRestart)
        $form.ShowDialog() | Out-Null
    }
    else {
        $RestartPC = "True"   }
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
$jsonContent | Out-File -FilePath "X:\Program Files\WindowsPowerShell\Modules\OSDCloud\*\Workflow\Default\os-arm64.json" -Encoding utf8 -Force

Write-Host "File 'os-arm64.json' has been created successfully." -ForegroundColor Cyan

##Check Drivers

#if ($null -eq $(Get-OSDCatalogDriverPack).name){iex(irm https://raw.githubusercontent.com/Justin-Swets/OSD/refs/heads/main/Get-SurfaceDriversv4.ps1)}
Deploy-OSDCloud
Test-OSDCloudProvisionValidation
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
$jsonContent | Out-File -FilePath "X:\Program Files\WindowsPowerShell\Modules\OSDCloud\*\Workflow\Default\os-amd64.json" -Encoding utf8 -Force

Write-Host "File 'os-amd64.json' has been created successfully." -ForegroundColor Cyan

##Check Drivers

#if ($null -eq $(Get-OSDCatalogDriverPack).name){iex(irm https://raw.githubusercontent.com/Justin-Swets/OSD/refs/heads/main/Get-SurfaceDriversv4.ps1)}
Deploy-OSDCloud
Test-OSDCloudProvisionValidation
If ($RestartPC -eq "True") {
    Restart-Computer -Force}

}

