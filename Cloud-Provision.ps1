##Update OSDCLoudGUI Image Type

# 1. Define the content using a Here-String
$jsonContent = @"
{
    "OSActivation.default": "Volume",
    "OSActivation.values": [
        "Retail",
        "Volume"
    ],
    "OSEdition.default": "Enterprise",
    "OSEditionId.default": "Enterprise",
    "OSEdition.values": [
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
    ],
    "OSLanguageCode.default": "en-us",
    "OSLanguageCode.values": [
        "ar-sa", "bg-bg", "cs-cz", "da-dk", "de-de", "el-gr", "en-gb", "en-us",
        "es-es", "es-mx", "et-ee", "fi-fi", "fr-ca", "fr-fr", "he-il", "hr-hr",
        "hu-hu", "it-it", "ja-jp", "ko-kr", "lt-lt", "lv-lv", "nb-no", "nl-nl",
        "pl-pl", "pt-br", "pt-pt", "ro-ro", "ru-ru", "sk-sk", "sl-si", "sr-latn-rs",
        "sv-se", "th-th", "tr-tr", "uk-ua", "zh-cn", "zh-tw"
    ],
    "OSName.default": "Win11-25H2-arm64",
    "OSName.values": [
        "Win11-25H2-arm64",
        "Win11-24H2-arm64",
        "Win11-23H2-arm64"
    ]
}
"@

# 2. Output the content to the file
# We use -Encoding utf8 to ensure standard JSON compatibility
$jsonContent | Out-File -FilePath "X:\Program Files\WindowsPowerShell\Modules\OSDCloud\*\Workflow\Default\os-arm64.json" -Encoding utf8 -Force

Write-Host "File 'os-arm64.json' has been created successfully." -ForegroundColor Cyan

##Check Drivers

if ($null -eq $(Get-OSDCatalogDriverPack).name){iex(irm https://raw.githubusercontent.com/Justin-Swets/OSD/refs/heads/main/Get-SurfaceDriversv4.ps1)}
Deploy-OSDCloud