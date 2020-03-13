# Loads .NET framework to memory and suppresses the echo
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null

# File path and Prefix variable values are read from Input boxes
$path = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the folder path `n`nNote: This folder and any subfolders within will be affected", "Folder Path")
$prefix = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the prefix `n`nNote: All the Pdfs within this folder and subfolders will be affected", "Prefix")

# Using the Variables, change directory and add prefix to Pdf files only
cd $path
Get-ChildItem -Filter "*pdf*" -Recurse | Rename-Item -NewName {$prefix + $_.basename + $_.extension}