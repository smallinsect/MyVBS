Dim http
Set http = CreateObject("Msxml2.ServerXMLHTTP")
http.open "GET", "http://localhost:8879/user/list", False
http.send
msgbox http.status
msgbox http.responsetext
