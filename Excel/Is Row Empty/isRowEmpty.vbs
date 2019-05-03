Set appXL = CreateObject("Excel.application")
appXL.Application.WorkBooks.Open(WScript.Arguments.Item(0))
appXL.Application.Visible = True
Set excelSheet = appXL.Worksheets(WScript.Arguments.Item(1))
result = 0
If appXL.CountA(WScript.Arguments.Item(2)) = 0 Then
	result = 1
End If
appXL.Quit
WScript.StdOut.Write(result)
WScript.Quit
