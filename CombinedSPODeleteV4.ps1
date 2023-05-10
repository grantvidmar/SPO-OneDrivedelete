# Check PowerShell version
if ($PSVersionTable.PSEdition -ne 'Core') {
    Write-Host "This script must be run in PowerShell 7 or later."
    $ps7InstallScript = Join-Path -Path $PSScriptRoot -ChildPath "RunMe.ps1"
    if (Test-Path $ps7InstallScript) {
        Write-Host "To install PowerShell 7, please run the following script: $ps7InstallScript"
    } else {
        Write-Host "The PowerShell 7 installation script 'RunMe.ps1' was not found in the same directory."
    }
    exit
}

# Check if PnP.PowerShell module is installed
$pnPModule = Get-Module -ListAvailable -Name "PnP.PowerShell"
if (-not $pnPModule) {
    Write-Host "PnP.PowerShell module is not installed. Installing..." -ForegroundColor Green
    Install-Module -Name "PnP.PowerShell" -Force
} else {
    Write-Host "PnP.PowerShell module is already installed. Version: $($pnPModule.Version)" -ForegroundColor Green
}

# Check if Microsoft.Online.SharePoint.PowerShell module is installed
$sharePointModule = Get-Module -ListAvailable -Name "Microsoft.Online.SharePoint.PowerShell"
if (-not $sharePointModule) {
    Write-Host "Microsoft.Online.SharePoint.PowerShell module is not installed. Installing..." -ForegroundColor Green
    Install-Module -Name "Microsoft.Online.SharePoint.PowerShell" -Force
} else {
    Write-Host "Microsoft.Online.SharePoint.PowerShell module is already installed. Version: $($sharePointModule.Version)" -ForegroundColor Green
}

# Import required modules
try {
    if (-not (Get-Module -ListAvailable -Name "PnP.PowerShell")) {
        Import-Module -Name "PnP.PowerShell" -DisableNameChecking
    }
    if (-not (Get-Module -ListAvailable -Name "Microsoft.Online.SharePoint.PowerShell")) {
        Import-Module -Name "Microsoft.Online.SharePoint.PowerShell" -DisableNameChecking
    }

    Write-Host "Modules imported successfully." -ForegroundColor Green
}
catch {
    Write-Host "Error importing modules: $($_.Exception.Message)" -ForegroundColor Red
    return
}

# Check installed version of PnP.PowerShell
$PnPModule = Get-Module -Name PnP.PowerShell -ListAvailable | Select-Object -First 1
Write-Host "PnP.PowerShell Version: $($PnPModule.Version)"

# Check installed version of Microsoft.Online.SharePoint.PowerShell
$SharePointModule = Get-Module -Name Microsoft.Online.SharePoint.PowerShell -ListAvailable | Select-Object -First 1
Write-Host "Microsoft.Online.SharePoint.PowerShell Version: $($SharePointModule.Version)"

# Define the country code options
$countryCodes = @{
    "NAM" = "North America";
    "CAN" = "Canada";
    "APC" = "Asia Pacific";
    "AUS" = "Australia";
    "EUR" = "EMEA Region";
    "FRA" = "France";
    "DEU" = "Germany";
    "ZAF" = "South Africa";
    "IND" = "India";
    "BRA" = "Brazil";
    ""    = "United Kingdom"
}

# Display the available country codes
Write-Host "Available country codes:"
foreach ($code in $countryCodes.GetEnumerator() | Sort-Object Value) {
    Write-Host "$($code.Value) = $($code.Name)"
}

# Prompt the user to select a country code
do {
    $CountryCode = Read-Host "Please enter the country code" 
} until ($countryCodes.ContainsKey($CountryCode))

# Prompt for user's UPN and replace non-letter characters with underscores
$UserUPN = Read-Host "Enter user's UPN" 
#$UserUnderscore = "melissa_tang2_clydeco_com" #Use this for non standard UPN
$UserUnderscore = $UserUPN -replace "[^a-zA-Z]", "_"
$SitePath = "/personal/$UserUnderscore"

# Define SharePoint Online URL
$tenantMap = Read-Host "Enter Tenant Map"
$tenantName = "$tenantMap$CountryCode"
$baseUrl = "https://$tenantName-my.sharepoint.com"
Write-Host "$baseUrl"

# Get the site
$site = "$baseUrl/personal/$UserUnderscore"

# Display site information
Write-Host "Site URL: $site"
$DeletePath = "/Documents/Delete"
$deleteSite = $site + $DeletePath
$FolderSiteRelativeURL = $SitePath + $DeletePath
write-host "$deleteSite"-foregroundcolor magenta
Read-Host -Prompt "Press enter to verify the site is correct"

# Connect to SharePoint Online
Connect-PnPOnline -Url $site -Interactive 

# Check if the connection is successful
if (Get-PnPConnection) {
    Write-Host "Connected to SharePoint Online"-ForegroundColor Green
} else {
    Write-Host "Failed to connect to SharePoint Online"-ForegroundColor Red
}

#Count and display numer of items in the delete folder
$itemCount = (Get-PnPFolderItem $deleteSite).Count
Write-Host "Total items in Delete Folder: $itemCount"

$Web = Get-PnPWeb
$Folder = Get-PnPFolder -Url $FolderSiteRelativeURL
     
# Function to recursively remove files and folders from the path given.
Function Clear-PnPFolder([Microsoft.SharePoint.Client.Folder]$Folder) {
    $InformationPreference = 'Continue'
    If ($Web.ServerRelativeURL -eq '/') {
        $FolderSiteRelativeURL = $Folder.ServerRelativeUrl
    } Else {       
        $FolderSiteRelativeURL = $Folder.ServerRelativeUrl.Replace($Web.ServerRelativeURL, [string]::Empty)
    }
    # First remove all files in the folder.
    $Files = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderSiteRelativeURL -ItemType File
    $TotalFiles = $Files.Count
    $ProcessedFiles = 0
    ForEach ($File in $Files) {
        # Delete the file.
        Remove-PnPFile -ServerRelativeUrl $File.ServerRelativeURL -Force -Recycle
        Write-Information ("Deleted File: '{0}' at '{1}'" -f $File.Name)
        
        $ProcessedFiles++
        $PercentageComplete = ($ProcessedFiles / $TotalFiles) * 100
        $ProgressStatus = "Deleting Files: {0}/{1} ({2}%)" -f $ProcessedFiles, $TotalFiles, $PercentageComplete
        Write-Progress -Activity "Clearing Folder" -Status $ProgressStatus -PercentComplete $PercentageComplete
    }
    # Second loop through subfolders and remove them - unless they are "special" or "hidden" folders.
    $SubFolders = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderSiteRelativeURL -ItemType Folder
    $TotalFolders = $SubFolders.Count
    $ProcessedFolders = 0
    Foreach ($SubFolder in $SubFolders) {
        If (($SubFolder.Name -ne 'Forms') -and (-Not($SubFolder.Name.StartsWith('_')))) {
            # Recurse into children.
            Clear-PnPFolder -Folder $SubFolder -Force
            # Finally delete the now empty folder.
            Remove-PnPFolder -Name $SubFolder.Name -Folder ($Site + $Folder) -Force
            Write-Information ("Deleted Folder: '{0}' at '{1}'" -f $SubFolder.Name)
        }
        
        $ProcessedFolders++
        $PercentageComplete = ($ProcessedFolders / $TotalFolders) * 100
        $ProgressStatus = "Deleting Folders: {0}/{1} ({2}%)" -f $ProcessedFolders, $TotalFolders, $PercentageComplete
        Write-Progress -Activity "Clearing Folder" -Status $ProgressStatus -PercentComplete $PercentageComplete
    }
    $InformationPreference = 'SilentlyContinue'
}     
# Call the function to empty folder if it exists.
if ($null -ne $Folder) {
    Clear-PnPFolder -Folder $Folder -Force
} Else {
    Write-Error ("Folder '{0}' not found" -f $FolderSiteRelativeURL)
}
Read-Host -Prompt "Press any key to continue..."

# Delete all subfolders at the specified location and display their names
$FoldersToDelete = Get-PnPFolder -FolderSiteRelativeUrl $deleteSite -Recursive
$TotalFolders = $FoldersToDelete.Count
$ProcessedFolders = 0

ForEach ($Folder in $FoldersToDelete) {
    $FolderName = $Folder.Name
    $FolderUrl = $Folder.ServerRelativeUrl

    Remove-PnPFolder -Name $FolderName -ServerRelativeUrl $FolderUrl -Force

    $ProcessedFolders++
    $PercentageComplete = ($ProcessedFolders / $TotalFolders) * 100
    $ProgressStatus = "Deleting Folders: {0}/{1} ({2}%)" -f $ProcessedFolders, $TotalFolders, $PercentageComplete
    Write-Progress -Activity "Deleting Subfolders" -Status $ProgressStatus -PercentComplete $PercentageComplete
}

Write-Host "All subfolders deleted."

# Display the total number of deleted items
$DeletedItemCount = $items.Count
Write-Host "Deleted $DeletedItemCount items." -ForegroundColor Yellow

# Count the number of items in the recycle bins
$firstStageRecycleBinCount = Get-PnPRecycleBinItem | Where-Object { $_.ItemType -ne "Folder" } | Measure-Object | Select-Object -ExpandProperty Count
$secondStageRecycleBinCount = Get-PnPRecycleBinItem -SecondStage | Where-Object { $_.ItemType -ne "Folder" } | Measure-Object | Select-Object -ExpandProperty Count
Write-Host "There are currently $firstStageRecycleBinCount items in the first stage recycle bin and $secondStageRecycleBinCount items in the second stage recycle bin." -ForegroundColor Cyan

#Empty Reycle Bins: Both 1st stage and 2nd Stage
Clear-PnPRecycleBinItem -All -Force

# Get the Preservation Hold library
$preservationHoldLibrary = Get-PnPList -Identity "Preservation Hold Library"

# Count the number of items in the library
$itemCount = (Get-PnPListItem -List $preservationHoldLibrary).Count

# Display the item count
Write-Host "The Preservation Hold library contains $itemCount items." -ForegroundColor Cyan

# Get all the items in the Preservation Hold Library
$items = Get-PnPListItem -List $preservationHoldLibrary

# Delete all the items in the Preservation Hold library using batch deletion with progress tracker
Write-Host "Deleting all items in the Preservation Hold library..." -ForegroundColor Yellow
$itemsToDelete = Get-PnPListItem -List $preservationHoldLibrary
for ($i = 0; $i -lt $itemsToDelete.Count; $i++) {
    $item = $itemsToDelete[$i]
    Write-Progress -Activity "Deleting items" -Status "Progress: $i / $itemCount" -PercentComplete ($i / $itemCount * 100)
    Remove-PnPListItem -List $preservationHoldLibrary -Identity $item.Id -Force
}
Write-Progress -Activity "Deleting items" -Status "Progress: $itemCount / $itemCount" -PercentComplete 100
Write-Host "All items deleted from the Preservation Hold library." -ForegroundColor Green