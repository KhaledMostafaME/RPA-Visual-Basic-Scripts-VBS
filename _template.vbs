' Turn on error Handling
On Error Resume Next



' Error Handler
If Err.Number <> 0 Then
	WScript.StdOut.Write("Error: " &  Err.Description)
	Err.Clear
	WScript.Quit
	Else
	WScript.StdOut.Write("Success")
End If
WScript.Quit
