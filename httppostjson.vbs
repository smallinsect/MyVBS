Dim fso, outFile
Set fso = CreateObject("Scripting.FileSystemObject")
Set outFile = fso.CreateTextFile("output.txt", True)

' This requires the Chilkat API to have been previously unlocked.
' See Global Unlock Sample for sample code

set req = CreateObject("Chilkat_9_5_0.HttpRequest")
set http = CreateObject("Chilkat_9_5_0.Http")

' This example duplicates the HTTP POST shown at
' http://json.org/JSONRequest.html

' Specifically, the request to be sent looks like this:

' POST /request HTTP/1.1
' Accept: application/jsonrequest
' Content-Encoding: identity
' Content-Length: 72
' Content-Type: application/jsonrequest
' Host: json.penzance.org
' 
' {"user":"doctoravatar@penzance.com","forecast":7,"t":"vlIj","zip":94089}

' First, remove default header fields that would be automatically
' sent.  (These headers are harmless, and shouldn't need to 
' be suppressed, but just in case...)
http.AcceptCharset = ""
http.UserAgent = ""
http.AcceptLanguage = ""
' Suppress the Accept-Encoding header by disallowing 
' a gzip response:
http.AllowGzip = 0

' If a Cookie needs to be added, it may be added by calling
' AddQuickHeader:
success = http.AddQuickHeader("Cookie","JSESSIONID=1234")

jsonText = "{""uids"":[], ""type"":2, ""title"":""222"", ""content"":""3333"", ""res"":""4444""}"

' To use SSL/TLS, simply use "https://" in the URL.

' IMPORTANT: Make sure to change the URL, JSON text,
' and other data items to your own values.  The URL used
' in this example will not actually work.

' resp is a Chilkat_9_5_0.HttpResponse
Set resp = http.PostJson("http://localhost:9520/user/mail",jsonText)
If (http.LastMethodSuccess <> 1) Then
    outFile.WriteLine(http.LastErrorText)
Else
    ' Display the JSON response.
    outFile.WriteLine(resp.BodyStr)
End If


outFile.Close