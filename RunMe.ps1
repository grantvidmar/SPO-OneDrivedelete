# Specify the path to the PowerShell 7 executable
$ps7ExecutablePath = "C:\Program Files\PowerShell\7\pwsh.exe"

# Check if the PowerShell 7 executable file exists
if (Test-Path $ps7ExecutablePath) {
    # PowerShell 7 is installed
    Write-Host "PowerShell 7 is already installed."

    $currentPath = (Get-Location).Path
    $combinedSPODeleteScript = Join-Path -Path $currentPath -ChildPath 'CombinedSPODeletev4.ps1'

    if (Test-Path $combinedSPODeleteScript) {
        & $combinedSPODeleteScript
    }
    else {
        Write-Host "The CombinedSPODeletev4.ps1 script was not found in the current path."
    }

    Write-Host "Press any key to continue..."
    $null = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
else {
    # PowerShell 7 is not installed
    Write-Host "PowerShell 7 is not installed."

    $currentPath = (Get-Location).Path
    $ps7InstallScript = Join-Path -Path $currentPath -ChildPath 'ps7install.ps1'

    if (Test-Path $ps7InstallScript) {
        Start-Process powershell.exe -Verb RunAs -ArgumentList "-File '$ps7InstallScript'"
    }
    else {
        Write-Host "The ps7install.ps1 script was not found in the current path."
    }
}


