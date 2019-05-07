inputString = Wscript.Arguments(0)
character = Wscript.Arguments(1)
Count = len(inputString) - len(replace(inputString, character, ""))
WScript.StdOut.Write(Count)
WScript.Quit