'Developed and Edited by Khaled Mostafa
'####### Usage #######
'Parameter 1: Pass full TIFF file path with extension
'Parameter 2: Pass full path of save location 
'Parameter 3: Pass extension type of the output image with dot. ie -> .jpg .png .jpeg .tiff .tif
'For More info: https://github.com/KhaledMostafaME

Dim Img, myPage, v, lp 
Set Img = WScript.CreateObject("WIA.ImageFile")

Img.LoadFile WScript.Arguments.Item(0)

For lp = 1 To Img.FrameCount
Img.ActiveFrame = lp
Set v = Img.ARGBData
Set myPage = v.ImageFile(Img.Width, Img.Height)
myPage.SaveFile WScript.Arguments.Item(1) & "\img_" & lp & WScript.Arguments.Item(2)
Next