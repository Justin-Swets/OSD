# --- CONFIGURATION ---
$XmlPath = "X:\Program Files\WindowsPowerShell\Modules\OSD\*\cache\driverpack-catalogs\build-driverpacks.xml"
$SurfaceDocsUrl = "https://learn.microsoft.com/en-us/surface/manage-surface-driver-and-firmware-updates"

# --- FUNCTIONS ---

function Sync-DriverToXml {
    param (
        [string]$Path,
        [string]$ModelName,
        [string]$NewUrl
    )

    if (-not (Test-Path $Path)) {
        Write-Error "XML File not found at $Path"
        return
    }

    [xml]$xmlData = Get-Content $Path
    
    # Target the specific 'Obj' containing the Model
    $ModelNode = $xmlData.Objs.Obj.MS.S | Where-Object { $_.N -eq "Model" -and $_.'#text' -eq $ModelName }

    if ($ModelNode) {
        # UPDATE LOGIC
        $UrlNode = $ModelNode.ParentNode.S | Where-Object { $_.N -eq "Url" }
        if ($UrlNode) {
            $UrlNode.'#text' = $NewUrl
            Write-Host "[XML] Updated existing entry for '$ModelName'." -ForegroundColor Green
        }
    } else {
        # ADD NEW ENTRY LOGIC
        Write-Host "[XML] No entry found for '$ModelName'. Creating new entry..." -ForegroundColor Yellow
        
        # 1. Create the new Obj element
        $newObj = $xmlData.CreateElement("Obj", "http://schemas.microsoft.com/powershell/2004/04")
        
        # 2. Determine next RefId (Simple increment based on count)
        $nextRefId = $xmlData.Objs.Obj.Count.ToString()
        $newObj.SetAttribute("RefId", $nextRefId)

        # 3. Create the inner MS (MemberSet) and properties
        # This mirrors the structure found in the source XML file
        $ms = $xmlData.CreateElement("MS", $xmlData.DocumentElement.NamespaceURI)
        
        $props = @(
            @{ N = "Manufacturer"; V = "Microsoft" }
            @{ N = "Model";        V = $ModelName }
            @{ N = "Url";          V = $NewUrl }
            @{ N = "OS";           V = "Windows 10/11 x64" }
            @{ N = "Guid";         V = [guid]::NewGuid().ToString() }
        )

        foreach ($p in $props) {
            $s = $xmlData.CreateElement("S", $xmlData.DocumentElement.NamespaceURI)
            $s.SetAttribute("N", $p.N)
            $s.InnerText = $p.V
            $null = $ms.AppendChild($s)
        }

        $null = $newObj.AppendChild($ms)
        $null = $xmlData.Objs.AppendChild($newObj)
        Write-Host "[XML] Successfully added new model entry to XML." -ForegroundColor Green
    }

    $xmlData.Save($Path)
}

# --- MAIN SCRAPER EXECUTION ---

Write-Host "[1/4] Starting Surface Driver Pack Scraper..." -ForegroundColor Cyan

try {
    # 1. Scrape URLs
    $Response = Invoke-WebRequest -Uri $SurfaceDocsUrl -UseBasicParsing
    $DriverPackLinks = $Response.Links | 
        Where-Object { $_.href -like "*microsoft.com*download*details.aspx?id=*" } | 
        Select-Object -ExpandProperty href -Unique

    $DriverMap = @()

    # 2. Map URLs to models
    Write-Host "[2/4] Mapping $(( $DriverPackLinks.Count )) URLs to models..." -ForegroundColor Gray
    foreach ($Url in $DriverPackLinks) {
        try {
            $DetailPage = Invoke-WebRequest -Uri $Url -UseBasicParsing -TimeoutSec 5
            if ($DetailPage.Content -match "<title>Download (.*?) Drivers") {
                $CleanModel = ($Matches[1] -replace " Drivers.*", "").Trim()
                $DriverMap += [PSCustomObject]@{ Model = $CleanModel; URL = $Url }
            }
        } catch { continue }
    }

    # 3. Detect Local Model
    $LocalModel = (Get-CimInstance Win32_ComputerSystem).Model
    Write-Host "[3/4] Local System Detected: $LocalModel" -ForegroundColor Cyan

    # 4. Advanced Scoring Match
    Write-Host "[4/4] Calculating best match score..." -ForegroundColor Gray
    $ScoredMatches = foreach ($Item in $DriverMap) {
        $MatchScore = 0
        $ScrapedWords = $Item.Model.ToLower().Split(" ", [System.StringSplitOptions]::RemoveEmptyEntries)
        foreach ($Word in $ScrapedWords) { if ($LocalModel.ToLower().Contains($Word)) { $MatchScore++ } }
        [PSCustomObject]@{ Model = $Item.Model; URL = $Item.URL; Score = $MatchScore }
    }

    $BestMatch = $ScoredMatches | Where-Object { $_.Score -gt 1 } | 
                 Sort-Object Score, {$_.Model.Length} -Descending | Select-Object -First 1

    if ($BestMatch) {
        # Sync the best match to the XML file
        Sync-DriverToXml -Path $XmlPath -ModelName $BestMatch.Model -NewUrl $BestMatch.URL
        return $BestMatch.URL
    } else {
        Write-Error "Could not find a high-confidence match for: $LocalModel"
    }

} catch {
    Write-Error "A critical error occurred: $($_.Exception.Message)"
}