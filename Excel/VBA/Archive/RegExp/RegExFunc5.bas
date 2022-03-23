Attribute VB_Name = "RegExFunc"
Option Explicit
Option Compare Text
Option Base 1

'RegExp Function Library

Public RegExO As Object

Public Function RegExTest(ByVal Source, ByVal p As String, Optional ByVal i As Integer = 0, Optional ByVal c As Integer = 0) As Boolean
  
  If RegExO Is Nothing Then
  
    Set RegExO = CreateObject("vbscript.regexp")
    
    With RegExO
      .IgnoreCase = True
      .Global = True
      .MultiLine = True
      .Pattern = p
    End With
    
  ElseIf RegExO.Pattern <> p Then
  
    Let RegExO.Pattern = p
    
  End If

  RegExTest = RegExO.Test(Source).Length > 0

End Function

Public Function RegExExecute(ByVal Source, ByVal p As String, Optional ByVal i As Integer = 0, Optional ByVal c As Integer = 0)

  If RegExO Is Nothing Then
  
    Set RegExO = CreateObject("vbscript.regexp")
    
    With RegExO
      .IgnoreCase = True
      .Global = True
      .MultiLine = True
      .Pattern = p
    End With
    
  ElseIf RegExO.Pattern <> p Then
  
    Let RegExO.Pattern = p
    
  End If
  
  RegExExecute = RegExO.Execute(Source).Item(i).SubMatches.Item(c)

End Function

Public Function RegExReplace(ByVal Source, ByVal p As String, ByVal ReplaceWith) As String

  If RegExO Is Nothing Then
  
    Set RegExO = CreateObject("vbscript.regexp")
    
    With RegExO
      .IgnoreCase = True
      .Global = True
      .MultiLine = True
      .Pattern = p
    End With
    
  ElseIf RegExO.Pattern <> p Then
  
    Let RegExO.Pattern = p
    
  End If

  RegExReplace = RegExO.Replace(Source, ReplaceWith)

End Function
