'Open Cisco VPN and enter Credentials
Set WshShell = WScript.CreateObject("WScript.Shell")

WshShell.Run """%PROGRAMFILES(x86)%\Cisco\Cisco AnyConnect Secure Mobility Client\vpnui.exe"""

WScript.Sleep 2000

WshShell.AppActivate "Cisco AnyConnect Secure Mobility Client"

WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{ENTER}"		

WScript.Sleep 5000

WshShell.SendKeys "Nagakaju!69"
WshShell.SendKeys "{ENTER}"
