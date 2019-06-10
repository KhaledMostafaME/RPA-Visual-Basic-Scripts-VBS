Dim i, col, Input, Output
Input = WScript.Arguments.Item(0)
For i = Len(Input) To 1 Step -1
    col = col + (Asc(Mid(Input, i, 1)) - 64) * (26 ^ (i - 1))
Next
Output = col
WScript.StdOut.Write(Output)
'Thanks to https://stackoverflow.com/a/15637514/6906583 