#iex(irm https://raw.githubusercontent.com/Justin-Swets/OSD/refs/heads/main/Cloud-Provision.ps1)
Start-Process powershell.exe -ArgumentList @(
    "-NoProfile", 
    "-ExecutionPolicy Bypass", 
    "-WindowStyle Normal", 
    "-Command", "iex(irm https://raw.githubusercontent.com/Justin-Swets/OSD/refs/heads/main/Cloud-Provision.ps1)" -wait -windowstyle normal
)