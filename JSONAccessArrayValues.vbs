Dim fso, outFile
Set fso = CreateObject("Scripting.FileSystemObject")
Set outFile = fso.CreateTextFile("output.txt", True)

set json = CreateObject("Chilkat_9_5_0.JsonObject")

jsonStr = "{ ""id"": 1, ""name"": ""A green door"", ""tags"": [""home"", 22, ""green""], ""price"": 125 }"

success = json.Load(jsonStr)
If (success <> 1) Then
    outFile.WriteLine(json.LastErrorText)
    WScript.Quit
End If

' Get the "tags" array, which contains "home", 22, "green"
' tagsArray is a Chilkat_9_5_0.JsonArray
Set tagsArray = json.ArrayOf("tags")
If (json.LastMethodSuccess = 0) Then
    outFile.WriteLine("tags member not found.")
    WScript.Quit
End If

' Get the value at each array index. 
' Output will be:
' [0] = home
' [0] as integer = 0
' [1] = 22
' [1] as integer = 22
' [2] = green
' [2] as integer = 0

arraySize = tagsArray.Size

For i = 0 To arraySize - 1

    sValue = tagsArray.StringAt(i)

    outFile.WriteLine("[" & i & "] = " & sValue)

    iValue = tagsArray.IntAt(i)
    outFile.WriteLine("[" & i & "] as integer = " & iValue)

Next

' Note: The StringAt method returns the value as a string regardless of the type.

' The IntAt method returns the value as an integer.  If the value does not convert to 
' an integer, then 0 is returned

outFile.Close