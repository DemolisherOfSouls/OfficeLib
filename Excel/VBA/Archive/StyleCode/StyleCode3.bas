Attribute VB_Name = "StyleCode"
Option Explicit
Option Compare Text
Option Base 1

Private Const DefStyle = "Normal"
Private Const DefTitleText = "Added Title"

Private Function RemoveTop(r As Range) As Range

  If r.Rows.Count > 1 Then
    Set RemoveTop = r.Offset(1, 0).Resize(r.Rows.Count - 1, r.Columns.Count)
  Else
    Set RemoveTop = r
  End If
  
End Function

Private Function SelStyle(r As Range) As String

  SelStyle = r.Resize(1, 1).Style.Name

End Function

Private Function GetTCHeader(r As Range) As Range
  
  Set GetTCHeader = Intersect(r.EntireColumn, RemoveTop(r.CurrentRegion)).Rows(1)

End Function

Private Function GetTCBody(r As Range) As Range

  Set GetTCBody = RemoveTop(Intersect(r.EntireColumn, RemoveTop(r.CurrentRegion)))

End Function

Private Sub TableColumn(ByVal HStyle As String, ByVal BStyle As String, Ref As Range)

  Dim Sel As Range
  Dim TColumn As Range
  
  Set Sel = Ref.Resize(1, 1)
  
  Set TColumn = Intersect(Sel.EntireColumn, Sel.CurrentRegion)
  
  RemoveTop(TColumn).Style = BStyle
  
  
  TColumn.Rows(1).Style = HStyle
  TColumn.Select

End Sub

'Public Functions & Subroutines

'Remove Built in Styles
Public Sub RemoveBadStyles()

  Dim s As Style
  For Each s In ActiveWorkbook.Styles
  
 If STARTSWITH(s.Name, "20") Then GoTo Del
 If STARTSWITH(s.Name, "40") Then GoTo Del
 If STARTSWITH(s.Name, "20") Then GoTo Del
 If STARTSWITH(s.Name, "20") Then GoTo Del
 If STARTSWITH(s.Name, "20") Then GoTo Del
 If STARTSWITH(s.Name, "20") Then GoTo Del
 
 Exit Sub
 
Del:
 
 ActiveWorkbook.Styles(s.Name).Delete
  
  Next

End Sub

Public Sub LookupColumn()

  TableColumn HStyle:="LkpHd", BStyle:="LkpCell", Ref:=Selection

End Sub

Public Sub CalcColumn()

  TableColumn HStyle:="CalcHd", BStyle:="CalcCell", Ref:=Selection

End Sub

Public Sub FixColumn()

  Dim Sel, STyp, HTyp, BTyp

  Sel = SelStyle(Selection)
  
  If Sel Like "Lkp*" Or Sel Like "Int*" Or Sel Like "Inp*" Then
    STyp = Left(Sel, 3)
    
  ElseIf Sel Like "Calc*" Then
    STyp = Left(Sel, 4)
  
  ElseIf Sel Like "Act*" Or Sel = DefStyle Then
    Exit Sub
  
  End If
  
  If Sel Like "*Hd" Then
    HTyp = "Hd"
    BTyp = "Cell"
    
  ElseIf Sel Like "*Cell" Or Sel Like "*Date" Then
    HTyp = "Hd"
    BTyp = Right(Sel, 4)
    
  ElseIf Sel Like "*Val" Then
    HTyp = "Hd"
    BTyp = Right(Sel, 3)
    
  ElseIf Sel Like "*HdKey" Then
    HTyp = "HdKey"
    BTyp = "Key"
    
  End If

  TableColumn HStyle:=(STyp & HTyp), BStyle:=(STyp & BTyp), Ref:=Selection

End Sub

Public Sub AddTitle()

  With Selection
    .EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    .Offset(-1, 0).Select
  End With
  
  With Selection
    .Style = "BoxTitle"
    .Merge
    .Value = DefTitleText
  End With
   
End Sub

