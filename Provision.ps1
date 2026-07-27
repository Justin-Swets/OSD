
#iex(irm https://raw.githubusercontent.com/Justin-Swets/OSD/refs/heads/main/Cloud-Provision.ps1)
Start-Process powershell.exe -ArgumentList @(
    "-NoProfile", 
    "-ExecutionPolicy Bypass", 
    "-WindowStyle Normal", 
    "-Command `"iex (irm 'https://raw.githubusercontent.com/Justin-Swets/OSD/refs/heads/main/Cloud-Provision.ps1')`""
    ) -wait -windowstyle normal
If ((get-ciminstance -Class "Win32_ComputerSystem").Model -like "*Surface Laptop for Business 7th*") {
   Start-Process powershell.exe -ArgumentList @(
    "-NoProfile", 
    "-ExecutionPolicy Bypass", 
    "-WindowStyle Normal", 
    "-Command `"iex (irm 'https://raw.githubusercontent.com/Justin-Swets/OSD/refs/heads/main/OOB-KB5121767.ps1')`""
    ) -wait -windowstyle normal
# Do something
}