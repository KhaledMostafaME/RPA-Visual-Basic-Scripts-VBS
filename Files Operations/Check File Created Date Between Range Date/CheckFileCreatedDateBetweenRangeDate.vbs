Dim Fso, FileObj, FilePath, StartDate, EndDate, FileDate, Result 
Set Fso = CreateObject("Scripting.FileSystemObject")
FilePath = WScript.Arguments.Item(0)
StartDate = WScript.Arguments.Item(1)
EndDate = WScript.Arguments.Item(2)
If (Fso.FileExists(FilePath)) Then
	Set FileObj = Fso.GetFile(FilePath)
	FileDate = FormatDateTime(FileObj.DateCreated,2)
	If (FileDate >= StartDate AND FileDate <= EndDate) Then
		Result = 1
		Else
		Result = 0 
	End If
Else 'File Doesn't Exit.
	Result = -1
End If
WScript.StdOut.Write(Result)
WScript.Quit
