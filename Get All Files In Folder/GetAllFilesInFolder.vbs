'Developed and Edited by Khaled Mostafa
'
'####### Usage #######
'Input (String): Full folder path as a parameter and without quotations
'Output (String): Listing all the files in one string.
'Example: Input: C:\Users\Khaled\Desktop\
'
'####### License #######
'MIT License
'Copyright (c) 2019 Khaled Mostafa
'For More info: https://github.com/KhaledMostafaME

Set objFSO = CreateObject("Scripting.FileSystemObject")
objStartFolder = WScript.Arguments.Item(0)
Set objFolder = objFSO.GetFolder(objStartFolder)
Dim Files
Set colFiles = objFolder.Files
For Each objFile in colFiles
    Files = Files + objFile.Name + vbCr
Next
WScript.StdOut.WriteLine Files