Dim http, fso, outFile

Set fso = CreateObject("Scripting.FileSystemObject")
Set outFile = fso.CreateTextFile("useraddr.json", True)

Set http = CreateObject("Msxml2.ServerXMLHTTP")
http.open "GET", "http://localhost:9520/log/userAddr?page=1&pageNum=20", False
http.send
msgbox http.status
outFile.WriteLine(http.responsetext)
