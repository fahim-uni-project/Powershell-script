# Get the registry key for Chrome
$registryPath32 = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall"

$installedApps = @()

# Query 32-bit registry path
Get-ChildItem -Path $registryPath32 | Where-Object { $_.GetValue("DisplayName") -like "Google Chrome*" } | ForEach-Object {
    $app = New-Object -TypeName PSObject
    $app | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $_.GetValue("DisplayName")
    $app | Add-Member -MemberType NoteProperty -Name "Version" -Value $_.GetValue("DisplayVersion")
    $installedApps += $app
}

# Check if Chrome version is not equal to the specified version
$versionToCheck = "123.0.6312.59"
foreach ($app in $installedApps) {
    if ($app.Version -ne $versionToCheck) {
        Write-Host "$($app.DisplayName) version $($app.Version) is not the required version."
        # Define the URL for the Chrome MSI
$url = "https://dl.google.com/dl/chrome/install/googlechromestandaloneenterprise64.msi"

# Define the path where you want to save the downloaded MSI file
$downloadPath = "C:\Temp\googlechromestandaloneenterprise64.msi"

# Create the directory if it doesn't exist
if (-not (Test-Path -Path "C:\Temp")) {
    New-Item -ItemType Directory -Path "C:\Temp"
}

# Download the MSI file
Invoke-WebRequest -Uri $url -OutFile $downloadPath

# Install Chrome silently
Start-Process -FilePath msiexec.exe -ArgumentList "/i `"$downloadPath`" /quiet /qn /norestart" -Wait


    }
    else{ Write-Host "$($app.DisplayName) version $($app.Version) is the  version."}
    exit 0
}


