Input = WScript.Arguments.Item(0)
StartDate = WScript.Arguments.Item(1)
EndDate = WScript.Arguments.Item(2)

If (Input >= StartDate AND Input <= EndDate) Then
		Result = 1
		Else
		Result = 0 
End If
WScript.StdOut.Write(Result)