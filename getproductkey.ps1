# Get the Windows product key
$windowsProductKey = (Get-WmiObject -query 'select * from SoftwareLicensingService').OA3xOriginalProductKey

# Display the product key
Write-Host "Windows Product Key: $windowsProductKey"
