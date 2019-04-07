'Developed and Edited by Khaled Mostafa

Dim Fso, Msg, FileObj, FilePath
Set Fso = CreateObject("Scripting.FileSystemObject") 'Creates "FileSystemObject" Object.
FilePath = WScript.Arguments.Item(0) 'Getting parameter
If (Fso.FileExists(FilePath)) Then 'Checks Whether File Exits At The Specified Path
	Set FileObj = Fso.GetFile(FilePath) 'Returns "File" Object
	Msg = FileObj.DateLastModified
Else 'File Doesn't Exit.
	Msg = "File : " & FilePath & " Doesn't Exist."
End If
WScript.StdOut.WriteLine Msg 'return the msg
