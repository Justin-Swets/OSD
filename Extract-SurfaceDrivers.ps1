# Define your paths
$msiPath = "C:\Users\jswet\Downloads\SurfaceLaptop_13in_1st_Edition_Win11_26100_25.124.43263.0.msi"
$outPath = "C:\WinPEDrivers\Arm64\full"

# Create the folder if it doesn't exist
New-Item -ItemType Directory -Force -Path $outPath

# Run the administrative installation (extraction)
Start-Process "msiexec.exe" -ArgumentList "/a `"$msiPath`" /qb TARGETDIR=`"$outPath`"" -Wait