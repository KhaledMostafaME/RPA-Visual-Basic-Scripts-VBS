'https://www.w3schools.com/asp/func_formatnumber.asp

Dim vIn, vPoints, vOut
vIn = Wscript.Arguments(0) 
vPoints = Wscript.Arguments(1)
vOut = FormatNumber(vIn,vPoints,,,0)
