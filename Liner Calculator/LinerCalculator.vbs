' Turn on error Handling
On Error Resume Next

Dim ii, sOperator, strExpr, y 
strExpr = WScript.Arguments.Item(0)
' insert spaces around all operators
strExpr = Replace(strExpr,"=","") 
For Each sOperator in Array("+","-","*","/","%")
  strExpr = Trim( Replace( strExpr, sOperator, Space(1) & sOperator & Space(1)))
Next
' replace all multi spaces with a single space 
Do While Instr( strExpr, Space(2))
  strExpr = Trim( Replace( strExpr, Space(2), Space(1)))
Loop
result = Eval(strExpr)

' Error Handler
If Err.Number <> 0 Then
	WScript.StdOut.Write("Error: " &  Err.Description)
	Err.Clear
	WScript.Quit
	Else
	WScript.StdOut.Write(result)
End If
WScript.Quit
