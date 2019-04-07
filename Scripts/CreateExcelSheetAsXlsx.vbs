'Developed and Edited by Khaled Mostafa
'####### Usage #######
'Pass full file path with extension as a parameter 
'For More info: https://github.com/KhaledMostafaME

strFileName = WScript.Arguments.Item(0)
Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Add()
objWorkbook.SaveAs(strFileName)
objExcel.Quit