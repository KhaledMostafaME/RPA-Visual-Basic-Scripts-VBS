Dim Fso, Msg, FileObj, FilePath
Set Fso = CreateObject("Scripting.FileSystemObject") 'Creates "FileSystemObject" Object.
FilePath = InputBox("Enter Path Of The File : ","File Name") 'Get The Required File From The User.
If (Fso.FileExists(FilePath)) Then 'Checks Whether File Exits At The Specified Path
	Set FileObj = Fso.GetFile(FilePath) 'Returns "File" Object
	Msg = FileObj.DateLastModified

Else 'File Doesn't Exit.
	Msg = "File : " & FilePath & " Doesn't Exist."
End If
MsgBox Msg