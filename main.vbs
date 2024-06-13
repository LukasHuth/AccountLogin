set script = CreateObject("WScript.Shell")

set service = GetObject("winmgmts:")
For Each Process In Service.InstancesOf("Win32_Process")
  If Process.Name = "Riot Client.exe" Then
    MsgBox("League seems to be currently open, please close it and try again")
    WScript.Quit
  End If
Next

script.Run "mshta.exe ""E:\vbs\league\menu.hta"""

