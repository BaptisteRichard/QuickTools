Set WshShell = WScript.CreateObject("WScript.Shell")

For i = 1 To 120
  WshShell.SendKeys "{NUMLOCK}"
  WshShell.SendKeys "{NUMLOCK}"
  WScript.Sleep 290000
Next

x=msgbox("Task scheduler" ,4096, "Task completed")
