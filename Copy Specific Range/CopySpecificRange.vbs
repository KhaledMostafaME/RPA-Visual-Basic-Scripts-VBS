Dim appXL
Dim wbSource
Dim wbDest
Dim wksSource
Dim wksDest
Set appXL = CreateObject("Excel.Application")
appXL.Visible = True
Set wbSource = appXL.Workbooks.Open(WScript.Arguments.Item(0))
Set wksSource = wbSource.Worksheets(WScript.Arguments.Item(1))
Set wbDest = appXL.Workbooks.Open(WScript.Arguments.Item(3))
Set wksDest = wbDest.Worksheets(WScript.Arguments.Item(4))

wksDest.Range(WScript.Arguments.Item(5)).Value = wksSource.Range(WScript.Arguments.Item(2)).Value

wbDest.Close True
wbSource.Close False
appXL.Quit

Set appXL = Nothing
Set wbSource = Nothing
Set wksSource = Nothing
Set wbDest = Nothing
Set wksDest = Nothing
WScript.Quit


