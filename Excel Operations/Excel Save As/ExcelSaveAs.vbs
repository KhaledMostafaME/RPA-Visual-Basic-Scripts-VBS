InputFile = Wscript.Arguments(0)
SheetName = Wscript.Arguments(1)
OutputFile = Wscript.Arguments(2)
Ext =  Wscript.Arguments(3)
Set objExcel = CreateObject("Excel.application")
objExcel.application.visible=false
objExcel.application.displayalerts=false
set objExcelBook = objExcel.Workbooks.Open(InputFile)
objExcel.Sheets(SheetName).Select
objExcelBook.SaveAs OutputFile, Ext
objExcel.Application.Quit
objExcel.Quit
