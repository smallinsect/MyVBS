set json = CreateObject("Chilkat_9_5_0.JsonObject")
json.EmitCompact = 0

success = json.LoadFile("test.json")
If (success <> 1) Then
    msgbox json.LastErrorText
    WScript.Quit
End If

msgbox json.Emit()

Dim url, json
url="http://localhost:8879/mail"
set Http=createobject("MSXML2.XMLHTTP")
Http.Open "POST", url, False
http.setRequestHeader "content-type","application/json"
http.Send json.Emit()
msgbox http.responsetext
