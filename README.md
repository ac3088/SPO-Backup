# Requirements

 - PowerShell version 7 or higher (7.4 is the latest stable build when this was written): https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7.4
 - PnP PowerShell: https://pnp.github.io/powershell/articles/installation.html

# How to Use

Before the script can be used, line 7 needs to be edited to have the correct SharePoint URL. Just change it to whatever your SharePoint URL is, for example:
`$SiteURL = "https://xxxx.sharepoint.com/sites/$Site` --> `$SiteURL = https://catlovers.sharepoint.com/sites/$Site`
After that's done, just open the "run.bat" file to run it. If you encounter any errors, they should be written in the "log.txt" file.

**The log file is re-created/reset every time the script is run, so don't start the script again if you want to read the log.**
