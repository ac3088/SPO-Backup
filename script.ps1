# Create error log file
New-Item -Path . -Name "log.txt" -ItemType File -Force | Out-Null
$LogPath = "./log.txt"

# Get the desired site and connect to Sharepoint
$Site = Read-Host -Prompt "Enter the name of site to backup"
$SiteURL = "https://xxxx.sharepoint.com/sites/$Site"
try {
    Connect-PnPOnline -Url $SiteURL -UseWebLogin
} catch {
    $Error | Out-File -Path $script:LogPath -Append
    Write-Host "There was an error connecting to the requested site`nCheck the log file for more information" -f Red
    Read-Host "Press the ENTER key to continue"
    Exit
}

# Create a new directory to store the backup
$Date = Get-Date -Format dd-MM-yyyy
$DirName = "$Site $Date"
if (Test-Path -Path "./$DirName") {
    $Input = Read-Host -Prompt "A folder called $DirName already exists. Would you like to overwrite it? [y/n]"
    while ($Input -ne "y" -and $Input -ne "n") {
        $Input = Read-Host -Prompt "Please enter either 'y' or 'n'"
    }
    if ($Input -eq "n") {
        Read-Host "The program will now close. Press the ENTER key to continue"
        Exit
    }
}
New-Item -Path . -Name $DirName -ItemType Directory -Force | Out-Null
$DirPath = "./$DirName"
Write-Host "New folder created ($DirName)Files and folders will be saved to this folder" -f Green

# Variables used to count the total amount of files retrieved
$FilesRetrieved = 0
$FilesNotRetrieved = @()
$TotalFiles = 0

# Gets all the files in the given directory and downloads them to the given path
function Get-Files {
    param (
        $FolderName,
        $Destination
    )
    $Files = Get-PnPFolderItem -Identity $FolderName -ItemType File
    foreach ($File in $Files) {
        $script:TotalFiles++
        Write-Host "Retrieving $($File.Name)... " -f Yellow -NoNewLine
        try {
            Get-PnPFile -ServerRelativeUrl $File.ServerRelativeUrl -Path $Destination -FileName $File.Name -AsFile -Force | Out-Null
            $script:FilesRetrieved++
            Write-Host "Done" -f Green
        } catch {
            $Error | Out-File -Path $script:LogPath -Append
            $script:FilesNotRetrieved += $File.Name
            Write-Host "Failed" -f Red
        }
    }
}

# Gets all folders in the given directory and recreates them in the given path
function Get-Folders {
    param (
        $FolderName,
        $Destination
    )
    $Folders = Get-PnPFolderItem -Identity $FolderName -ItemType Folder
    foreach($Folder in $Folders) {
        Write-Host "Creating $($Folder.Name) folder... " -f Yellow -NoNewLine
        New-Item -Path $Destination -Name $Folder.Name -ItemType Directory -Force | Out-Null
        Write-Host "Done" -f Green
        Get-Folders $Folder.ServerRelativeUrl "$Destination/$($Folder.Name)"
        Get-Files $Folder.ServerRelativeUrl "$Destination/$($Folder.Name)"
    }
}

# Gets all files, folders, and subfolders from the given directory and recreates 
# them at the given path on the local computer
function Get-Site {
    param (
        $FolderName,
        $Destination
    )
    Get-Files $FolderName $Destination
    Get-Folders $FolderName $Destination
}

Measure-Command { Get-Site "Shared Documents" $DirPath | Out-Host }
if ($FilesRetrieved -eq $TotalFiles) {
    Write-Host "All files retrieved successfully ($FilesRetrieved/$TotalFiles)" -f Green
} else {
    Write-Host "The following files couldn't be retrieved ($FilesRetrieved/$TotalFiles):" -f Red
    $FilesNotRetrieved | ForEach-Object { Write-Host $_ -f Red }
    Write-Host "Check the log file for more information" -f Red
}
Write-Host "Local backup created in folder $DirName"
Read-Host "Finished. Press the ENTER key to continue"