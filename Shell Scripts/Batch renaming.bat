::Set location for powershell:
cd "T:\Private\A&B REPORTS\Batch file renaming"

::Start the transcript:
Start-Transcript

::See the effect of renaming operation:
Get-ChildItem -Filter "*0118*" -Recurse | Rename-Item -NewName {$_.name -replace '0118','0218' } -whatif

::End the transcript:
Stop-Transcript

::Clear the screen:
cls

::Execute the renaming operation:
Get-ChildItem -Filter "*0118*" -Recurse | Rename-Item -NewName {$_.name -replace '0118','0218' } 

:: If you want to append a string ("_AP Units") to all the files in a directory
Get-ChildItem | Rename-Item -NewName {$_.basename + "_AP Units" + $_.extension}
