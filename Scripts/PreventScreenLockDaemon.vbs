' Prevent computer lock screen during automation process
' Need to [Open Program/File] Command,not [Run Script] Command 

Set WshShell = WScript.CreateObject("WScript.Shell")
While True
  WScript.Sleep 1000*60
  WshShell.SendKeys "{NUMLOCK}"
  WshShell.SendKeys "{NUMLOCK}"
Wend
