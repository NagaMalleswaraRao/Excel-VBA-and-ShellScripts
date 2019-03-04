:::: Objective: Call Cisco VPN file, Call the essential windows programs

:: Stop the terminal from printing these statements out
@echo off

:: Call VB Script (named StartCiscoVPN) 
Start C:\Public\"Batch Files"\StartCiscoVPN.vbs

:: Wait a few seconds, it would look cool if every application starts one after another and you can see the process
timeout /t 6

:: Open Outlook application
Start C:\"Program Files (x86)\Microsoft Office"\root\Office16\OUTLOOK.EXE

timeout /t 5

:: Start Skype for business
start C:\"Program Files (x86)\Microsoft Office"\root\Office16\lync.exe

timeout /t 3

:: Open a file explorer window
START C:\Windows\explorer.EXE

timeout /t 4

:: Open Personal emails, these have passwords saved. So, you can get away with not entering credentials
start chrome https://mail.google.com/mail/u/0/#inbox

start chrome https://mail.google.com/mail/u/1/#inbox

Exit


