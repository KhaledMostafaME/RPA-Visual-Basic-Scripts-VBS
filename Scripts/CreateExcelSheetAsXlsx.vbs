strFileName = WScript.Arguments.Item(0)
Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Add()
objWorkbook.SaveAs(strFileName)
objExcel.Quit