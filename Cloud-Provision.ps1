if ($null -eq $(Get-OSDCatalogDriverPack).name){iex(irm https://raw.githubusercontent.com/Justin-Swets/OSD/refs/heads/main/Get-SurfaceDriversv4.ps1)}
Deploy-OSDCloud