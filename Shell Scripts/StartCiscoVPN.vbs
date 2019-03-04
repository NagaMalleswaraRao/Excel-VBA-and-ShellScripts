'''Objective: Open Cisco VPN and enter Credentials

Set WshShell = WScript.CreateObject("WScript.Shell")

  'Explanation for a Layman: Does the same as clicking on CiscoVPN application to start
WshShell.Run """%PROGRAMFILES(x86)%\Cisco\Cisco AnyConnect Secure Mobility Client\vpnui.exe"""

  'Wait for few seconds (2) till you can see the UI
WScript.Sleep 2000

WshShell.AppActivate "Cisco AnyConnect Secure Mobility Client"

  'Navigate towards "Connect" button
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{ENTER}"		
  
  'Wait (5s) till it opens credentials window which has Username saved but not password.
WScript.Sleep 5000

  'Enter Password and press Enter
WshShell.SendKeys "Password123"
WshShell.SendKeys "{ENTER}"
  
  
