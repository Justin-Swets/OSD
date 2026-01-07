# --- CONFIGURATION ---
$XmlPath = "X:\Program Files\WindowsPowerShell\Modules\OSD\*\cache\driverpack-catalogs\build-driverpacks.xml" 
$SurfaceDocsUrl = "https://support.microsoft.com/en-us/surface/download-drivers-and-firmware-for-surface-09bb2e09-2a4b-cb69-0951-078a7739e120"

# --- FUNCTIONS ---

function Sync-DriverToXml {
    param ([string]$Path, [string]$ModelName, [string]$NewUrl, [string]$FileName)

    if (-not (Test-Path $Path)) { Write-Error "XML File not found at $Path"; return }
    
    [xml]$xmlData = Get-Content $Path
    $root = $xmlData.DocumentElement
    
    $nsManager = New-Object System.Xml.XmlNamespaceManager($xmlData.NameTable)
    $nsManager.AddNamespace("ps", "http://schemas.microsoft.com/powershell/2004/04")
    
    $ModelNode = $xmlData.SelectSingleNode("//ps:S[@N='Model' and text()='$ModelName']", $nsManager)

    if ($ModelNode) {
        $parentMS = $ModelNode.ParentNode
        $UrlNode = $parentMS.SelectSingleNode("ps:S[@N='Url']", $nsManager)
        $FileNode = $parentMS.SelectSingleNode("ps:S[@N='FileName']", $nsManager)
        if ($UrlNode) { $UrlNode.'#text' = $NewUrl }
        if ($FileNode) { $FileNode.'#text' = $FileName }
        Write-Host "[XML] Updated existing entry for '$ModelName'." -ForegroundColor Green
    } else {
        Write-Host "[XML] Creating new entry for '$ModelName'..." -ForegroundColor Yellow
        $ns = $root.NamespaceURI
        $newObj = $xmlData.CreateElement("Obj", $ns)
        $newObj.SetAttribute("RefId", $xmlData.SelectNodes("//ps:Obj", $nsManager).Count.ToString())
        
        $ms = $xmlData.CreateElement("MS", $ns)
        $props = @(
            @{ N = "Manufacturer"; V = "Microsoft" }; @{ N = "Model"; V = $ModelName }
            @{ N = "Url"; V = $NewUrl }; @{ N = "FileName"; V = $FileName }
            @{ N = "OS"; V = "Windows 10/11 x64" }; @{ N = "Guid"; V = [guid]::NewGuid().ToString() }
        )

        foreach ($p in $props) {
            $s = $xmlData.CreateElement("S", $ns); $s.SetAttribute("N", $p.N); $s.InnerText = $p.V
            $null = $ms.AppendChild($s)
        }
        $null = $newObj.AppendChild($ms); $null = $root.AppendChild($newObj)
    }
    $xmlData.Save($Path)
}

# --- MAIN EXECUTION ---

Write-Host "[1/4] Scraping Microsoft Support Surface Page..." -ForegroundColor Cyan
try {
    $Response = Invoke-WebRequest -Uri $SurfaceDocsUrl -UseBasicParsing
    $Links = $Response.Links | Where-Object { $_.href -like "*details.aspx?id=*" } | Select-Object -ExpandProperty href -Unique
    
    $DriverMap = @()
    Write-Host "[2/4] Resolving model names and filenames..." -ForegroundColor Gray
    
    foreach ($Url in $Links) {
        try {
            $Page = Invoke-WebRequest -Uri $Url -UseBasicParsing -TimeoutSec 5
            $Name = if ($Page.Content -match "<title>Download (.*?) Drivers") { ($Matches[1] -replace " Drivers.*", "").Trim() }
            $File = if ($Page.Content -match "([\w\d\-_]+\.msi)") { $Matches[1] }
            if ($Name) { $DriverMap += [PSCustomObject]@{ Model = $Name; URL = $Url; FileName = $File } }
        } catch { continue }
    }

    $LocalModelRaw = (Get-CimInstance Win32_ComputerSystem).Model
    $LocalModelNoBus = $LocalModelRaw -replace " for Business", ""
    Write-Host "[3/4] Local System Detected: $LocalModelRaw" -ForegroundColor Cyan

    # --- TIERED MATCHING LOGIC ---
    Write-Host "[4/4] Executing Tiered Matching (Exact > Normalized > Filename > Multi-Factor Score)..." -ForegroundColor Gray
    $BestMatch = $null

    # TIER 1: Literal Exact
    $BestMatch = $DriverMap | Where-Object { $_.Model -ieq $LocalModelRaw } | Select-Object -First 1

    # TIER 2: "For Business" Removed Match
    if (-not $BestMatch -and $LocalModelRaw -like "*for Business*") {
        $BestMatch = $DriverMap | Where-Object { $_.Model -ieq $LocalModelNoBus } | Select-Object -First 1
    }

    # TIER 3: Filename Specific Keyword Match (Requires 3+ Keyword matches)
    if (-not $BestMatch) {
        $keywords = $LocalModelNoBus.ToLower().Split(" ", [System.StringSplitOptions]::RemoveEmptyEntries)
        $BestMatch = $DriverMap | Where-Object { 
            $fn = $_.FileName.ToLower()
            ($keywords | Where-Object { $fn.Contains($_) }).Count -ge 3
        } | Select-Object -First 1
    }

    # TIER 4: Combined Best Fit Score (Model Title + Filename)
    if (-not $BestMatch) {
        Write-Host "      Tiers 1-3 Failed. Calculating Multi-Factor Score..." -ForegroundColor DarkGray
        $BestMatch = $DriverMap | Select-Object *, @{Name='Score'; Expression={
            $score = 0
            $targetWords = $LocalModelNoBus.ToLower().Split(" ", [System.StringSplitOptions]::RemoveEmptyEntries)
            $scrapedTitle = $_.Model.ToLower()
            $scrapedFile  = $_.FileName.ToLower()

            foreach ($word in $targetWords) {
                # Weight Model Title matches higher
                if ($scrapedTitle.Contains($word)) { $score += 10 }
                # Weight Filename matches as a secondary helper
                if ($scrapedFile.Contains($word))  { $score += 5  }
            }
            $score
        }} | Sort-Object Score -Descending | Select-Object -First 1
    }

    # --- FINAL SYNC ---
    if ($BestMatch -and ($BestMatch.Score -gt 0 -or $null -ne $BestMatch.Model)) {
        Write-Host "`nMATCH FOUND: $($BestMatch.Model)" -ForegroundColor Green
        Sync-DriverToXml -Path $XmlPath -ModelName $BestMatch.Model -NewUrl $BestMatch.URL -FileName $BestMatch.FileName
        return $BestMatch.URL
    } else {
        Write-Error "No reliable match found for $LocalModelRaw"
    }

} catch {
    Write-Error "Critical error: $($_.Exception.Message)"
}