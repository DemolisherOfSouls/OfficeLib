Attribute VB_Name = "LinkFunc"
Option Explicit
Option Compare Text
Option Base 1

'Link Generating Function Library
'Version 1.0.0

Public Sub LinkFromContent()
Attribute LinkFromContent.VB_ProcData.VB_Invoke_Func = "L\n14"

  Dim cs As Range, C As Range
  Set cs = Selection.Cells
  
  If cs.Count = 0 Then Exit Sub
  
  For Each C In cs
    Dim add As String
    add = IIf(Contains(C, "://"), "", "https://") & C
    
    If Not IsEmpty(C) Then ActiveSheet.Hyperlinks.add Anchor:=C, Address:=add, TextToDisplay:=C
  Next C

End Sub
