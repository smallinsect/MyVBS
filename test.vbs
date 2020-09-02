

set json = CreateObject("Chilkat_9_5_0.JsonObject")
json.EmitCompact = 0

' Assume the file contains the data as shown above..
success = json.LoadFile("test.json")
If (success <> 1) Then
    msgbox json.LastErrorText
    WScript.Quit
End If

msgbox json.Emit()
