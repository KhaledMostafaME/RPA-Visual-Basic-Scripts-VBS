inputString = Wscript.Arguments(0)
inputStringToCount = Wscript.Arguments(1)
Count = len(inputString) - len(replace(inputString, inputStringToCount, ""))
WScript.StdOut.Write(Count)
WScript.Quit