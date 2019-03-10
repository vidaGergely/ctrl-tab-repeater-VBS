Set args = Wscript.Arguments
Dim intPID
Dim secundum
Dim messageBox

If  WScript.arguments.length <> 0 Then
	intPID = Wscript.Arguments.Item(0)
	secundum = Wscript.Arguments.Item(1)
	
	messageBox = MsgBox("A Ctrl + Tab Repeater program fut!"& vbcrlf &"Ha be szeretne zarni a programot, nyomjon az OK-ra!! "& vbcrlf & "Beallitott sebesseg: " & secundum &" masodperc" , 0,"CTRL + TAB Repeater leallitasa")
	If messageBox = 1 Then
		strComputer = "."
		Set objWMIService = GetObject ("winmgmts:\\" & strComputer & "\root\cimv2")
		Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where ProcessID = " & intPID & "")
		For Each objProcess in colProcessList
			objProcess.Terminate()
		Next
	End If
	
Else
	MsgBox "Sajnos a program magaban nem tud futni!"
End If

			


