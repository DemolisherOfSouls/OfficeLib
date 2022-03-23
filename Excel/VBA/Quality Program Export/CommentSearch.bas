Attribute VB_Name = "CommentSearch"
Option Explicit
Option Compare Text
Option Base 1

Function Comment_Search(ByVal Value As String, ByVal Source As String, Optional ByVal R1 As String, Optional ByVal R2 As String, Optional ByVal R3 As String, Optional ByVal R4 As String)
  On Error GoTo ErrorHandler
  
  Dim Temp As String
  If InStr(1, Source, Value, 1) > 0 Then
    Temp = Mid(Source, InStr(1, Source, Value, 1), 100)
    If InStr(1, Temp, Chr(10), vbTextCompare) > 0 Then
        Temp = Left(Temp, InStr(1, Temp, Chr(10), vbTextCompare) - 1)
    End If
    If InStr(1, Temp, ":", vbTextCompare) > 0 Then
        Temp = Right(Temp, Len(Temp) - InStr(1, Temp, ":", vbTextCompare))
    End If
    Temp = Replace(Temp, R1, "")
    Temp = Replace(Temp, R2, "")
    Temp = Replace(Temp, R3, "")
    Temp = Replace(Temp, R4, "")
    Temp = Replace(Temp, " ", "")
    Comment_Search = Temp
  Else
    Comment_Search = ""
  End If
Exit Function

ErrorHandler:
  MsgBox "Error with Comment Search Function"
End Function
