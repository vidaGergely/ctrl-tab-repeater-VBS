Dim merged
Dim WshShell
Dim objShell
Set objShell = Wscript.CreateObject("WScript.Shell")
Dim c
Dim s
Dim ProcessId
ProcessId = CurrProcessId 'will remain valid indefinitely
'--Author-- Gergely V.
Function CurrProcessId
	Dim oShell, sCmd, oWMI, oChldPrcs, oCols, lOut
	lOut = 0
	Set oShell  = CreateObject("WScript.Shell")
	Set oWMI    = GetObject(_
	"winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	sCmd = "/K " & Left(CreateObject("Scriptlet.TypeLib").Guid, 38)
	oShell.Run "%comspec% " & sCmd, 0
	WScript.Sleep 100 'For healthier skin, get some sleep
	Set oChldPrcs = oWMI.ExecQuery(_
	"Select * From Win32_Process Where CommandLine Like '%" & sCmd & "'",,32)
	For Each oCols In oChldPrcs
		lOut = oCols.ParentProcessId 'get parent
		oCols.Terminate 'process terminated
		Exit For
	Next
	CurrProcessId = lOut
End Function
'-------------------------------------------------------------------------------------------------------------------
Set WshShell = WScript.CreateObject("WScript.Shell")
c = InputBox("- CTRL + TAB billentyu kombinacio ismetlodese megadott intervallum alapjan " & vbcrlf &vbcrlf&_
"- Egy bongeszo ablakon belul nyisd meg a lapfuleket"& vbcrlf & vbcrlf& "- A script leallitasahoz hasznald a hatterben megnyilo az 'OK' gombot" &vbcrlf & vbcrlf &_
"Tabulator valtas kozotti sebesseg[sec]: ","CTRL + TAB repeater - Bongeszo lapfulek valtasahoz","10")

If IsNumeric(c) = True AND c <> "" AND CInt(c)>0 Then
	merged = "passedPidKiller.vbs " & ProcessId & " " & c
	s =CInt(c)*1000
	objShell.Run merged 
	Do while s > 0 
		WScript.Sleep s		
		WshShell.SendKeys"^{TAB}"
		WScript.Sleep 50
	Loop
Else
	
	MsgBox "A program kilep!" 
	WScript.Quit
End If








