# Define the folder paths and file paths
$folderPath = "C:\sqltemp"
$vcRedist_x86_FilePath = Join-Path $folderPath "vc_redist.x86.exe"
$vcRedist_x64_FilePath = Join-Path $folderPath "vc_redist.x64.exe"
$msoledbsqlFilePath = Join-Path $folderPath "msoledbsql.msi"

# Create the folder
New-Item -ItemType Directory -Force -Path $folderPath
# Define the content of the batch file
$content = @"
@echo off
setlocal

REM Define folder paths and file paths
set "folderPath=C:\sqltemp"
set "logFilePathole=%folderPath%\installation_log_ole.txt"
set "logFilePath64=%folderPath%\installation_log_64.txt"
set "logFilePath86=%folderPath%\installation_log_86.txt"
set "vcRedistFolderPath=%folderPath%\vc_redist"
set "vcRedist_x86_FilePath=%folderPath%\vc_redist.x86.exe"
set "vcRedist_x64_FilePath=%folderPath%\vc_redist.x64.exe"
set "msoledbsqlFilePath=%folderPath%\msoledbsql.msi"

REM Create the folder if it doesn't exist
if not exist "%folderPath%" mkdir "%folderPath%"

REM Check if Visual C++ Redistributable x86 is installed
reg query "HKLM\Software\Microsoft\VisualStudio\14.0\VC\Runtimes\x86" >nul 2>&1
if %errorlevel% neq 0 (
    REM Download Visual C++ Redistributable x86
    powershell -command "(New-Object Net.WebClient).DownloadFile('https://aka.ms/vs/17/release/vc_redist.x86.exe', '%vcRedist_x86_FilePath%')"
    REM Install Visual C++ Redistributable x86
    start /wait "" "%vcRedist_x86_FilePath%" /install /quiet /norestart /l*v "%logFilePath86%"
)

REM Check if Visual C++ Redistributable x64 is installed
REM reg query "HKLM\Software\Microsoft\VisualStudio\14.0\VC\Runtimes\x64" >nul 2>&1
REM if %errorlevel% neq 0 (
    REM Download Visual C++ Redistributable x64
    powershell -command "(New-Object Net.WebClient).DownloadFile('https://aka.ms/vs/17/release/vc_redist.x64.exe', '%vcRedist_x64_FilePath%')"
    REM Install Visual C++ Redistributable x64
    start /wait "" "%vcRedist_x64_FilePath%" /install /quiet /norestart /l*v "%logFilePath64%"
REM )

REM Download the OLE DB Driver for SQL Server
powershell -command "(New-Object Net.WebClient).DownloadFile('https://go.microsoft.com/fwlink/?linkid=2248728', '%msoledbsqlFilePath%')"

REM Install the OLE DB Driver for SQL Server
start /wait "" msiexec.exe /i "%msoledbsqlFilePath%" IACCEPTMSOLEDBSQLLICENSETERMS=YES /qn /l*v "%logFilePathole%"

REM Remove the folder
rd /s /q "%folderPath%"

echo Installation completed.
exit /b 0
"@

# Define the path for the batch file
$batchFilePath = "$folderPath\Install_MicrosoftOLEDB.bat"

# Save content to the batch file
$content | Out-File -FilePath $batchFilePath -Encoding ascii

# Confirm that the batch file has been saved
if (Test-Path $batchFilePath) {
    Write-Host "Batch file saved successfully at $batchFilePath."
    # Run the batch file
    Start-Process -FilePath $batchFilePath -WindowStyle Hidden -Wait
    # Get all versions of the Microsoft OLE DB Driver for SQL Server
$oldDrivers = Get-WmiObject -Class Win32_Product | Where-Object {$_.Name -like "Microsoft OLE DB Driver*"}

if ($oldDrivers) {
    Write-Host "Old versions of Microsoft OLE DB Driver for SQL Server found:"
    foreach ($driver in $oldDrivers) {
        Write-Host "Product Name: $($driver.Name), Version: $($driver.Version)"
    }
} else {
    Write-Host "No old versions of Microsoft OLE DB Driver for SQL Server found."
}

} else {
    Write-Host "Failed to save the batch file."
}
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

$oldDrivers = Get-WmiObject -Class Win32_Product | Where-Object {$_.Name -like "Microsoft OLE DB Driver*"}

if ($oldDrivers) {
    Write-Host "Old versions of Microsoft OLE DB Driver for SQL Server found:"
    foreach ($driver in $oldDrivers) {
        Write-Host "Product Name: $($driver.Name), Version: $($driver.Version)"
    }
} else {
    Write-Host "No old versions of Microsoft OLE DB Driver for SQL Server found."
}
