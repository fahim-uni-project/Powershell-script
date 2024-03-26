$paths = @("C:\Program Files\KeePass Password Safe 2\unins000.exe", "C:\Program Files (x86)\KeePass Password Safe 2\unins000.exe")
$uninstall = $paths | Where-Object { Test-Path $_ } | Select-Object -First 1
if ($uninstall) {
    Start-Process -FilePath $uninstall -ArgumentList "/VERYSILENT /NORESTART" -Wait -NoNewWindow
} else {
    Write-Host "The file does not exist in either path."
}