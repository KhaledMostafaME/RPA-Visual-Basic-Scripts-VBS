ExcelFile = Wscript.Arguments(0)
CsvFile = Wscript.Arguments(1)
SheetName = Wscript.Arguments(2)
Set objExcel = CreateObject("Excel.application")
objExcel.application.visible=false
objExcel.application.displayalerts=false
set objExcelBook = objExcel.Workbooks.Open(ExcelFile)
objExcel.Sheets(SheetName).Select
objExcelBook.SaveAs CsvFile, 23
objExcel.Application.Quit
objExcel.Quit