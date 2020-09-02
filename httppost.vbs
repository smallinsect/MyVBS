Dim url, json
url="http://localhost:9520/user/mail"
json="{""uids"":[], ""type"":1, ""title"":""111"", ""content"":""222"", ""res"":""333""}"
set Http=createobject("MSXML2.XMLHTTP")
Http.Open "POST", url, False
http.setRequestHeader "content-type","application/json"
http.Send json
html=http.responsetext
msgbox html
