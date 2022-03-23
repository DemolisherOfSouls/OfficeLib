Attribute VB_Name = "RegExFunc"
Option Explicit
Option Compare Text
Option Base 1

'RegExp Function Library

Public RegExO As Object

Public Function RegExTest(ByVal Source, ByVal p As String, Optional ByVal i As Integer = 0, Optional ByVal c As Integer = 0) As Boolean
  
  Set RegExO = New RegExp
  
  With RegExO
        .IgnoreCase = True
        .Global = (i > 0 Or c > 0)
        .MultiLine = True
        .Pattern = p
    End With

    RegExTest = RegExO.Test(Source).Length > 0

End Function

Public Function RegExExecute(ByVal Source, ByVal p As String, Optional ByVal i As Integer = 0, Optional ByVal c As Integer = 0)

  Set RegExO = New RegExp

  With RegExO
        .IgnoreCase = True
        .Global = (i > 1 Or c > 1)
        .MultiLine = True
        .Pattern = p
    End With

    Dim Result As MatchCollection :  Set Result = RegExO.Execute(Source)
  
  RegExExecute = Result.Item(i).SubMatches.Item(c)

End Function

Public Function RegExReplace(ByVal Source, ByVal p As String, ByVal ReplaceWith) As String

  Set RegExO = New RegExp
  
  With RegExO
        .IgnoreCase = True
        .Global = (i > 1 Or c > 1)
        .MultiLine = True
        .Pattern = p
    End With

    RegExReplace = RegExO.Replace(Source, ReplaceWith)

End Function
