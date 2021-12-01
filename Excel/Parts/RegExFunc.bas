Attribute VB_Name = "RegExFunc"
Option Explicit
Option Compare Text
Option Base 1

'RegExp Function Library

Private RegExO As New RegExp

Public Function RegExTest(ByVal Source As String, ByVal p As String, Optional ByVal i As Integer = 0, Optional ByVal c As Integer = 0) As Boolean
  
  
  With RegExO
    .IgnoreCase = True
    .Global = True
    .MultiLine = True
    .Pattern = p
  End With

  RegExTest = RegExO.Test(Source)

End Function

Public Function RegExExecute(ByVal Source As String, ByVal p As String, Optional ByVal i As Integer = 0, Optional ByVal c As Integer = 0) As String

  With RegExO
    .IgnoreCase = True
    .Global = True
    .MultiLine = True
    .Pattern = p
  End With
  
  RegExExecute = RegExO.Execute(Source).Item(i).SubMatches(c)

End Function

Public Function RegExQuick(ByVal Source As String, ByVal p As String) As String

  With RegExO
    .IgnoreCase = True
    .Global = False
    .MultiLine = True
    .Pattern = p
  End With
  
  RegExQuick = RegExO.Execute(Source).Item(0)
  
End Function

Public Function RegExReplace(ByVal Source As String, ByVal p As String, ByVal ReplaceWith) As String
  
  With RegExO
    .IgnoreCase = True
    .Global = True
    .MultiLine = True
    .Pattern = p
  End With

  RegExReplace = RegExO.Replace(Source, ReplaceWith)

End Function

Public Function ParseFraction(ByVal s As String) As Double
  On Error GoTo Invalid

  Dim whole, upper, lower
  
  With RegExO
    .Global = True
    .IgnoreCase = True
    .MultiLine = True
    .Pattern = "([\d\.]+)[  \-]+([\d\.]+)[\/\\  ]+([\d\.]+)"
  
    With .Execute(s).Item(0).SubMatches
      whole = .Item(0)
      upper = .Item(1)
      lower = .Item(2)
    End With
  End With
  
  ParseFraction = CDbl(upper) / CDbl(lower) + CInt(whole)
  
  Exit Function
Invalid:
  ParseFraction = CVErr(xlErrNum)
  
End Function
