Set objExcel = CreateObject("Excel.Application") 
Set objWorkbook = objExcel.Workbooks.Open(Wscript.Arguments(0)) 
objExcel.Application.Visible = False 
objExcel.Application.DisplayAlerts = False 
objExcel.ActiveWorkbook.SaveAs Wscript.Arguments(1), 56 
objExcel.ActiveWorkbook.Close 
objExcel.Application.DisplayAlerts = True 
objExcel.Application.Quit 
WScript.Quit