	Set oSHELL = CreateObject("WScript.Shell")
	oSHELL.Run "powershell.exe -windowstyle hidden -executionpolicy bypass -file C:\Dell\OutlookFixTask.ps1", 0, True
	Set oSHELL = Nothing
