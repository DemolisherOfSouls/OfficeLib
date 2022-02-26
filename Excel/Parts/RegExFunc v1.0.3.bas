Attribute VB_Name = "RegExFunc"
Option Explicit
Option Compare Text
Option Base 1

'`RegExp Function Library
'Version 1.0.3

'Imports
'Microsoft VBScript Regular Expressions 5.5

'History
' 1.0.3 - Added function for CodeFunc Library
'         Removed RegExQuick

'Current

Private RegExO As New RegExp

Private Const RXFullMatch = -1

Private Const RXL_PARSEFRAC = "([\d\.]+)[  \-]+([\d\.]+)[\/\\  ]+([\d\.]+)"
Private Const RXL_FORMSTART = "^\s*=\s*("
Private Const RXL_FORMNAMES = "[a-zA-Z][a-zA-Z0-9]*"

Public Function RegExTest(ByVal Source As String, ByVal Pattern As String) As Boolean
  
  With RegExO
    .IgnoreCase = True
    .Global = True
    .MultiLine = True
    .Pattern = Pattern
  End With

  RegExTest = RegExO.Test(Source)
End Function

Public Function RegExExecute(ByVal Source As String, ByVal Pattern As String, Optional ByVal i As Integer = 0, Optional ByVal c As Integer = RXFullMatch) As String

  With RegExO
    .IgnoreCase = True
    .Global = True
    .MultiLine = True
    .Pattern = Pattern
  End With
  
  if c = RXFullMatch then
    RegExExecute = RegExO.Execute(Source).Item(i)
  else
    RegExExecute = RegExO.Execute(Source).Item(i).SubMatches(c)
  end if
End Function

Public Function RegExReplace(ByVal Source As String, ByVal Pattern As String, ByVal ReplaceWith) As String
  
  With RegExO
    .IgnoreCase = True
    .Global = True
    .MultiLine = True
    .Pattern = Pattern
    RegExReplace = .Replace(Source, ReplaceWith)
  End With
End Function

Public Function ParseFraction(ByVal Source As String, Optional ByRef Out As Double) As Double
  On Error GoTo Invalid
  
  With RegExO
    .Global = True
    .IgnoreCase = True
    .MultiLine = True
    .Pattern = RXL_PARSEFRAC
  
    With .Execute(Source).Item(0).SubMatches
      Out = CDbl(.Item(1)) / CDbl(.Item(2)) + CInt(.Item(0))
    End With
  End With
  
  ParseFraction = Out
Exit Function
  
Invalid:
  Out = 0
  ParseFraction = CVErr(xlErrNum)
End Function

Public Function TryParseFraction(ByVal Val) As Double

  If IsNumeric(Val) Or IsEmpty(Val) Then
    TryParseFraction = IIf(IsNumeric(Val), CDbl(Val), 0)
    Exit Function
  End If
  
  Dim result: result = ParseFraction(CStr(Val))
  
  TryParseFraction = IIf(IsError(result), 0, result)
End Function

Public Function CheckFormulaFunction (ByRef Source As string, Optional ByVal Funct as String = RXL_FORMNAMES) as Boolean

  CheckFormulaFunction = RegExTest(Source, RXL_FORMSTART & Funct & ")")
End Function
