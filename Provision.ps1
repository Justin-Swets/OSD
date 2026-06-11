

function Get-BiosDate {
    $bios = Get-CimInstance -ClassName Win32_BIOS -ErrorAction Stop
    $releaseDate = $bios.ReleaseDate
    if (-not $releaseDate) {
        throw 'BIOS release date not available from Win32_BIOS.'
    }
    return [Management.ManagementDateTimeConverter]::ToDateTime($releaseDate)
}

function Get-OnlineUtcDate {
    try {
        $response = Invoke-RestMethod -Uri 'https://worldtimeapi.org/api/ip' -TimeoutSec 15
        if ($response.datetime) {
            return [DateTime]::Parse($response.datetime).ToUniversalTime()
        }
    } catch {
        Write-Verbose "Primary online time query failed: $_"
    }

    try {
        $response = Invoke-WebRequest -Uri 'https://www.microsoft.com' -TimeoutSec 15
        if ($response.Headers.Date) {
            return [DateTime]::Parse($response.Headers.Date).ToUniversalTime()
        }
    } catch {
        throw 'Unable to retrieve online date/time from known endpoints.'
    }
}

try {
    $biosDate = Get-BiosDate
    $onlineUtcDate = Get-OnlineUtcDate
    $biosUtcDate = $biosDate.ToUniversalTime()
    $difference = $onlineUtcDate - $biosUtcDate
    $MaxBiosTimeDifference = [TimeSpan]::FromMinutes(5)

    Write-Host "BIOS date (local): $biosDate"
    Write-Host "BIOS date (UTC):   $biosUtcDate"
    Write-Host "Online UTC date:   $onlineUtcDate"
    Write-Host "Difference:        $($difference.ToString())"

    if ($difference.Duration() -gt $MaxBiosTimeDifference) {
        $ShouldSetBiosTime = $true
        Write-Warning "BIOS date differs from online time by more than $($MaxBiosTimeDifference.TotalMinutes) minutes. Updating system clock."

        try {
            Set-Date -Date $onlineUtcDate.ToLocalTime() -ErrorAction Stop
            Write-Host 'System time has been updated to match online time.'
        } catch {
            Write-Error "Failed to set system time: $_"
            exit 1
        }
    }

    if ([math]::Abs($difference.TotalDays) -gt 365) {
        Write-Warning 'BIOS date differs from online time by more than one year.'
    } elseif ([math]::Abs($difference.TotalHours) -gt 24) {
        Write-Warning 'BIOS date differs from online time by more than 24 hours.'
    } else {
        Write-Host 'BIOS date is within 24 hours of online time.'
    }
} catch {
    Write-Error $_
    exit 1
}

#iex(irm https://raw.githubusercontent.com/Justin-Swets/OSD/refs/heads/main/Cloud-Provision.ps1)
Start-Process powershell.exe -ArgumentList @(
    "-NoProfile", 
    "-ExecutionPolicy Bypass", 
    "-WindowStyle Normal", 
    "-Command", "iex(irm https://raw.githubusercontent.com/Justin-Swets/OSD/refs/heads/main/Cloud-Provision.ps1)" -wait -windowstyle normal
)