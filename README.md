# SPO-OneDrivedelete
This is a set PowerShell script used to delete files from a specified folder in SPO/OneDrive.
I have created this as a result of needing to delete a large amount of files from a SPO users OneDrive. 

#RunMe.ps1
The first script that sould be run is the "RunMe.ps1"
This script shouyld check if PowerShell7 is installed, if it is it will run "CombinedSPODelete.ps1"
If Powershell7 is not installed it should run "ps7install.ps1"

#ps7install.ps1
This file should be launched as an admin as you may potentially need to install PS7 as an admin.
It should preform a check for PowerShell7 again and at that point download and install PowerShell7 via admin.

#CombinedSPODeleteV4.ps1
This is the file that should do the brunt of the work.
It will ask you for a bunch of information to enter The big ones are as follows
Country Code: this is if you have a multigeo company and have different location.
UPN: this would usually be something along the lines of "firtst.last@companyname.com"
The big thing to remember is that the script is looking for a location "/Documents/Delete". This will need to be created in the users OneDrive and all files that need to be deleted should be moved here.
After that it should clear out both First and Second Stage Recycle Bins.
Finally it will clear all items for the Preservation Hold Library. If your orginization has this setup you will need to make sure the user is in the aappropriate exception group for this.

This set of scripts are my first real dive into Powershell. I did not write this code myself I have used ChatGPT and prompt engineering to create this script. It probably could be better and made into one file but I am still working on that.
