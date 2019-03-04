:::: Objective: Call Cisco VPN file, Call the essential windows programs

:: Stop the terminal from printing these statements out
@echo off

:: Call VB Script (named StartCiscoVPN) 
Start C:\Public\"Batch Files"\StartCiscoVPN.vbs

:: Wait a few seconds (It would look cool if every application starts one after another and you are able see the process)
timeout /t 6

:: Open Outlook application
Start C:\"Program Files (x86)\Microsoft Office"\root\Office16\OUTLOOK.EXE

:: Wait and Start Skype for business
timeout /t 5
start C:\"Program Files (x86)\Microsoft Office"\root\Office16\lync.exe

:: Wait and Open a file explorer window
timeout /t 3
START C:\Windows\explorer.EXE

:: Wait and Open Personal emails, these have passwords saved. So, you can get away with not entering credentials
timeout /t 4
start chrome https://mail.google.com/mail/u/0/#inbox
start chrome https://mail.google.com/mail/u/1/#inbox

Exit


