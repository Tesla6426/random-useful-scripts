Set WshShell = WScript.CreateObject("WScript.Shell")
wscript.sleep 5000
WshShell.SendKeys("ABCDEFGHIJKLMNOPQRSTUVWXYZ")
WshShell.SendKeys("{ENTER}")
WshShell.SendKeys("abcdefghijklmnopqrstuvwxyz")
WshShell.SendKeys("{NUMLOCK}{CAPSLOCK}{SCROLLLOCK}")
