inputString = Wscript.Arguments(0)
character = Wscript.Arguments(1)
Count = len(inputString) - len(replace(inputString, character, ""))
Result = Count/len(character)
WScript.StdOut.Write(Result)
WScript.Quit
