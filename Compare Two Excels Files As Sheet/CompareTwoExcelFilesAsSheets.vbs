Set xl = CreateObject("Excel.Application")
xl.Visible = "False"
Set Compare1 = xl.Workbooks.Open(Wscript.Arguments(0))
Set Compare2 = xl.Workbooks.Open(Wscript.Arguments(2))
 
Set CompareSheet1 = Compare1.Sheets(Wscript.Arguments(1))
Set compareSheet2 = Compare2.Sheets(Wscript.Arguments(3))
 
For Each Cell In CompareSheet1.UsedRange
  If Cell.Value <> compareSheet2.Range(Cell.Address).Value Then
                ' highlight the cell if mismatch exist
        Cell.Interior.ColorIndex = 6
        End If
Next
Compare1.Save
Compare2.Save
Compare1.Close
Compare2.Close