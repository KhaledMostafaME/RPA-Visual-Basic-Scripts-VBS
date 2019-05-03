'Developed and Edited by Khaled Mostafa
'
'####### Usage #######
'Input (String): Full file path with extension as a parameter and without quotations  
'Example: Input: C:\Users\Khaled\Desktop\Test.xlsx
'
'####### License #######
'MIT License
'Copyright (c) 2019 Khaled Mostafa
'For More info: https://github.com/KhaledMostafaME

strFileName = WScript.Arguments.Item(0)
Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Add()
objWorkbook.SaveAs(strFileName)
objExcel.Quit