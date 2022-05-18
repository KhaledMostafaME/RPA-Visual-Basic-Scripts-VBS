InputFile = Wscript.Arguments(0)
SheetName = Wscript.Arguments(1)
OutputFile = Wscript.Arguments(2)
Ext =  Wscript.Arguments(3)
Dim password
If WScript.Arguments.Count > 4 then
	password = Wscript.Arguments(4)
End If
Set objExcel = CreateObject("Excel.application")
objExcel.application.visible=false
objExcel.application.displayalerts=false
If Not isEmpty(password) then
	set objExcelBook = objExcel.Workbooks.Open(InputFile,,,,password)
Else
	set objExcelBook = objExcel.Workbooks.Open(InputFile)
End If
objExcel.Sheets(SheetName).Select
objExcelBook.SaveAs OutputFile, Ext
objExcel.Application.Quit
objExcel.Quit
