Dim Img, myPage, v, lp 
Set Img = WScript.CreateObject("WIA.ImageFile")

Img.LoadFile WScript.Arguments.Item(0)

For lp = 1 To Img.FrameCount
Img.ActiveFrame = lp
Set v = Img.ARGBData
Set myPage = v.ImageFile(Img.Width, Img.Height)
myPage.SaveFile WScript.Arguments.Item(1) & "\img_" & lp & WScript.Arguments.Item(2)
Next