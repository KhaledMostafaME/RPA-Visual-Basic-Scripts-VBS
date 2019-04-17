Set objFSO = CreateObject("Scripting.FileSystemObject")
objStartFolder = WScript.Arguments.Item(0)
Set objFolder = objFSO.GetFolder(objStartFolder)
Dim Files
Set colFiles = objFolder.Files
For Each objFile in colFiles
    Files = Files + objFile.Name + vbCr
Next
WScript.StdOut.WriteLine Files