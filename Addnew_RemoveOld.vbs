'Script to add printers during user login
'Very important otherwise users will get errors on subsequent reconnections
On Error resume Next

'Set Old Server Name
strOldServerName = "test"

'Set New Server Name
strNewServerName = "win7printer"

'Get default WMI namespace on Windows machines
Set objWMIService = GetObject("winmgmts:root\cimv2")

'Get printers connected to a computer
Set colPrinters = objWMIService.ExecQuery("Select * From Win32_Printer")

' Initialise the Network connections object
Set objNet = CreateObject("WScript.Network")

For Each objPrinter in colPrinters
  	If objPrinter.Attributes And 64 Then
  	Else
      	strPrinterType = "Network"
	    If objPrinter.ServerName="\\" & strOldServerName Then
	        objNet.RemovePrinterConnection objPrinter.Name
	    End If
  	End If
Next

newPrinters = Array("HP Color LaserJet 9500 PCL 6","Kyocera Copystar 250ci","HP 915")

For Each Printer in newPrinters
	tmp = "\\" & strNewServerName & "\" & Printer
	objNet.AddWindowsPrinterConnection tmp
Next