$registryPath32 = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall"
$registryPath64 = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
$registryPathLocal = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall"

$installedApps = @()

# Query 32-bit registry path
Get-ChildItem -Path $registryPath32 | ForEach-Object {
    $app = New-Object -TypeName PSObject
    $app | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $_.GetValue("DisplayName")
    $app | Add-Member -MemberType NoteProperty -Name "UninstallString" -Value $_.GetValue("UninstallString")
    $app | Add-Member -MemberType NoteProperty -Name "Version" -Value $_.GetValue("DisplayVersion")
    $installedApps += $app
}

# Query 64-bit registry path
Get-ChildItem -Path $registryPath64 | ForEach-Object {
    $app = New-Object -TypeName PSObject
    $app | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $_.GetValue("DisplayName")
    $app | Add-Member -MemberType NoteProperty -Name "UninstallString" -Value $_.GetValue("UninstallString")
    $app | Add-Member -MemberType NoteProperty -Name "Version" -Value $_.GetValue("DisplayVersion")
    $installedApps += $app
}

# Query local registry path
Get-ChildItem -Path $registryPathLocal | ForEach-Object {
    $app = New-Object -TypeName PSObject
    $app | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $_.GetValue("DisplayName")
    $app | Add-Member -MemberType NoteProperty -Name "UninstallString" -Value $_.GetValue("UninstallString")
    $app | Add-Member -MemberType NoteProperty -Name "Version" -Value $_.GetValue("DisplayVersion")
    $installedApps += $app
}

# Output the installed apps
$installedApps | Format-Table -AutoSize DisplayName, Version, UninstallString
#$installedApps | Export-Csv -Path "C:\Users\Public\Downloads\InstalledApps.csv" -NoTypeInformation
Write-Output "Data saved in to: C:\Users\Public\Downloads\InstalledApps.csv"
