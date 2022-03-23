Attribute VB_Name = "WebFunc"
Option Explicit
Option Compare Text
Option Base 1

'Web Requests Function Library
'Version 1.0.0

'Imports
'Microsoft WinHTTP Services, version 5.1
'Microsoft Scripting Runtime Reference
'Microsoft ActiveX Data Objects 2.8 Library

Public Function MakeWebRequest(ByVal Method As String, ByVal URL As String, PostData) As String
    ' make sure to include the Microsoft WinHTTP Services in the project
    ' tools -> references -> Microsoft WinHTTP Services, version 5.1
    ' http://www.808.dk/?code-simplewinhttprequest
    ' http://msdn.microsoft.com/en-us/library/windows/desktop/aa384106(v=vs.85).aspx
    ' http://www.neilstuff.com/winhttp/

    ' create the request object
    Dim Request As New WinHttpRequest

    ' set timeouts
    ' http://msdn.microsoft.com/en-us/library/windows/desktop/aa384061(v=vs.85).aspx
    ' SetTimeouts(resolveTimeout, ConnectTimeout, SendTimeout, ReceiveTimeout)
    Request.SetTimeouts 60000, 60000, 60000, 60000

    ' make the request, http verb (method), url, false to force syncronous
    ' open(http method, absolute uri to request, async (true: async, false: sync)
    Request.Open Method, URL, False

    ' handle post content type
    If Method = "POST" Then
        Request.SetRequestHeader "Content-type", "application/x-www-form-urlencoded"
    End If

    ' set WinHttpRequestOption enumerations
    ' http://msdn.microsoft.com/en-us/library/windows/desktop/aa384108(v=vs.85).aspx

    ' set user agent
    Request.Option(0) = "Echovoice VBA HTTP Bot v0.1"

    ' set ssl ignore errors
    '   13056: ignore errors
    '   0: break on errors
    Request.Option(4) = 13056

    ' set redirects
    Request.Option(6) = True

    ' allow http to redirect to https
    Request.Option(12) = True

    ' send request
    ' send post data, should be blank for a get request
    Request.Send PostData

    ' read response and return
    MakeWebRequest = Request.ResponseText

End Function

Public Function ParseJSON(data As String) As Object
    ' http://www.ediy.co.nz/vbjson-json-parser-library-in-vb6-xidc55680.html
    ' take JSON and convert to object

    ' change code in cStringBuilder for Win64 systems
    ' https://code.google.com/p/vba-json/issues/detail?id=13

    ' add Microsoft Scripting Runtime Reference for Dictionary data type
    ' add ADO reference Microsoft ActiveX Data Objects 2.8 Library
    ' http://msdn.microsoft.com/en-us/library/aa241766(v=vs.60).aspx

    ' use Set when returning an object
    Set ParseJSON = JSON.Parse(data)

End Function

