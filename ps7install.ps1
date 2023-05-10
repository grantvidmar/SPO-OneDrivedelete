# Get the path to the PowerShell 7 executable
$pwsh7Path = "C:\Program Files\PowerShell\7\pwsh.exe"

# Check if PowerShell 7 is already installed
if (-not (Test-Path $pwsh7Path)) {
    # PowerShell 7 is not installed, so install it
    $url = "https://aka.ms/powershell7"
    $installerPath = "$env:TEMP\PowerShell7Installer.exe"
    Invoke-WebRequest -Uri $url -OutFile $installerPath
    Start-Process -FilePath $installerPath -Verb RunAs -Wait


    # Check if the installation was successful
    if (-not (Test-Path $pwsh7Path)) {
        Write-Output "PowerShell 7 installation failed."
        exit
    }
}

# Define the script filename to search for
$scriptFileName = "CombinedSPODeleteV4.ps1"

# Get the current script's directory
$scriptDirectory = $PSScriptRoot

# Construct the full path to the script
$scriptPath = Join-Path -Path $scriptDirectory -ChildPath $scriptFileName

# Wait until PowerShell 7 is available
while (-not (Test-Path $pwsh7Path)) {
    Write-Output "Waiting for PowerShell 7 to become available..."
    Start-Sleep -Seconds 5
}

# Launch the PowerShell script using PowerShell 7
Start-Process -FilePath $pwsh7Path -ArgumentList "-File `"$scriptPath`""