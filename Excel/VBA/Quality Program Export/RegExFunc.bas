Attribute VB_Name = "RegExFunc"
Option Explicit
Option Compare Text
Option Base 1

'RegExp Function Library

Private RegExO As New RegExp

Public Function RegExTest(ByVal Source As String, ByVal p As String, Optional ByVal i As Integer = 0, Optional ByVal C As Integer = 0) As Boolean
  
  With RegExO
    .IgnoreCase = True
    .Global = True
    .MultiLine = True
    .Pattern = p
  End With

  RegExTest = RegExO.Test(Source)

End Function

Public Function RegExExecute(ByVal Source As String, ByVal p As String, Optional ByVal i As Integer = 0, Optional ByVal C As Integer = 0) As String

  With RegExO
    .IgnoreCase = True
    .Global = True
    .MultiLine = True
    .Pattern = p
  End With
  
  RegExExecute = RegExO.Execute(Source).Item(i).SubMatches(C)

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

Public Function ParseFraction(ByVal S As String, Optional ByRef Out As Double) As Double
  On Error GoTo Invalid

  Dim Whole, upper, lower
  
  With RegExO
    .Global = True
    .IgnoreCase = True
    .MultiLine = True
    .Pattern = "([\d\.]+)[  \-]+([\d\.]+)[\/\\  ]+([\d\.]+)"
  
    With .Execute(S).Item(0).SubMatches
      Whole = .Item(0)
      upper = .Item(1)
      lower = .Item(2)
    End With
  End With
  
  Out = CDbl(upper) / CDbl(lower) + CInt(Whole)
  ParseFraction = Out
  
Exit Function
  
Invalid:
  Out = 0
  ParseFraction = CVErr(xlErrNum)
End Function

Public Function TryParseFraction(ByVal V As Variant) As Double

  If IsNumeric(V) Or IsEmpty(V) Then
    TryParseFraction = IIf(IsNumeric(V), CDbl(V), 0)
    Exit Function
  End If
  
  Dim result: result = ParseFraction(CStr(V))
  
  TryParseFraction = IIf(IsError(result), 0, result)
    
End Function
