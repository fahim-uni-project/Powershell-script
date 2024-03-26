# Define the product name of the driver
$productName = "Microsoft OLE DB Driver for SQL Server"

# Search for the uninstall string in the registry
$uninstallString = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" |
    Where-Object { $_.DisplayName -eq $productName } |
    Select-Object -ExpandProperty UninstallString
    # Search for the uninstall string in the registry
$Displayname = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" |
    Where-Object { $_.DisplayName -eq $productName } |
    Select-Object -ExpandProperty DisplayName

# Check if the uninstall string is found
if ($uninstallString) {
    # Replace "MsiExec.exe /I{" with the desired command
    $uninstallCommand = $uninstallString.Replace("MsiExec.exe /I{", "MsiExec.exe /x{")

    # Execute the uninstall string silently
    Start-Process -FilePath cmd.exe -ArgumentList "/C $uninstallCommand /qn" -WindowStyle Hidden -Wait
    Write-Host "$Displayname has been uninstalled successfully."
} else {
    Write-Host "$productName is not installed on this system."
}
