# Get the registry key for Chrome
$registryPath32 = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall"

$installedApps = @()

# Query 32-bit registry path
Get-ChildItem -Path $registryPath32 | Where-Object { $_.GetValue("DisplayName") -like "Java 8 Update 401*" } | ForEach-Object {
    $app = New-Object -TypeName PSObject
    $app | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $_.GetValue("DisplayName")
    $app | Add-Member -MemberType NoteProperty -Name "Version" -Value $_.GetValue("DisplayVersion")
    $installedApps += $app
}
$installedApps 