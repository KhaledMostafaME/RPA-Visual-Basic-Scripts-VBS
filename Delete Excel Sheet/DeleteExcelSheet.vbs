Set objFSO = CreateObject("Scripting.FileSystemObject")
src_file = objFSO.GetAbsolutePathName(Wscript.Arguments.Item(0))
sheetname = Wscript.Arguments.Item(1)
Dim oExcel
Set oExcel = CreateObject("Excel.Application")
Dim oBook
Set oBook = oExcel.Workbooks.Open(src_file)    
oExcel.DisplayAlerts = False    
oExcel.ScreenUpdating = False    
oBook.sheets(sheetname).Delete
oBook.save
oBook.close
oExcel.DisplayAlerts = True 
oExcel.ScreenUpdating = True 
oExcel.Quit