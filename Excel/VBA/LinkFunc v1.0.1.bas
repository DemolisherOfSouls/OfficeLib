Attribute VB_Name = "LinkFunc"
Option Explicit
Option Compare Text
Option Base 1

'`Link Generating Function Library
'Version 1.0.1

'History
' 1.0.1 - Allowed File Prefixes

'Current

Public Sub LinkFromContent()
Attribute LinkFromContent.VB_ProcData.VB_Invoke_Func = "L\n14"

  Dim cs As Range, C As Range
  Set cs = Selection.Cells
  
  If cs.Count = 0 Then Exit Sub
  
  For Each C In cs
    Dim add As String
    If Contains(Right(Left(C, 3), 2), ":/") Then
      add = ""
    ElseIf Not Contains(C, "://") Then
      add = "https://"
    End If
    add = add & C
    
    If Not IsEmpty(C) Then ActiveSheet.Hyperlinks.add Anchor:=C, Address:=add, TextToDisplay:=C
  Next C

End Sub
