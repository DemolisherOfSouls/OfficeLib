Attribute VB_Name = "StyleFunc"
Option Explicit
Option Compare Text
Option Base 1

'Style Function Library
'Version 1.0.1

Private Const DefStyle = "Normal"
Private Const DefTitleText = "Added Title"

Private Sub TableColumn(ByVal Typ As String, ByRef Ref As Range, Optional ByVal Body As String = "Cell", Optional ByVal Head As String = "Hd")
  Dim Bd: Set Bd = Intersect(Ref.ListObject.DataBodyRange, Ref.EntireColumn)
  Dim Hd: Set Hd = Intersect(Ref.ListObject.HeaderRowRange, Ref.EntireColumn)
  Bd.Style = Typ + Body
  Hd.Style = Typ + Head
End Sub

Public Sub LookupColumn()
  TableColumn "Lkp", Selection
End Sub

Public Sub CalcColumn()
  TableColumn "Calc", Selection
End Sub

Public Sub InputColumn()
  TableColumn "Inp", Selection
End Sub

Public Sub InternalColumn()
  TableColumn "Int", Selection
End Sub

Public Sub FixColumn()
  Dim Sel, STyp, HTyp, BTyp
  
  Set Sel = Selection.Style
  
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
  
  TableColumn STyp + HTyp, STyp + BTyp, Selection
  
End Sub

Public Sub AddTitle()

  With Selection.Resize(1).Offset(-1, 0)
    .Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromLeftOrAbove
    .Style = "BoxTitle"
    .Merge
    .Text = DefTitleText
  End With
  
End Sub

Public Sub RemoveAllDefStyles()

  Dim Item As Style
  For Each Item In ActiveWorkbook.Styles

    If .Name Like "*Accent*" Then GoTo Delete
    If .Name Like "Heading*" Then GoTo Delete
    GoTo Skip
Delete:
    Item.Delete
Skip:
  Next Item
End Sub

Public Sub UpdateStyles()

  RemoveAllDefStyles

  Dim Item As Style
  For Each Item In ActiveWorkbook.Styles
    
    With Item
      If EndsWith(.Name, "Title") Then
        .Font.Size = Range("TitleFontSize_Override")
      ElseIf EndsWith(.Name, "Hd") Or EndsWith(.Name, "HdKey") Then
        .Font.Size = Range("HeaderFontSize_Override")
      ElseIf EndsWith(.Name, "Cell") Or EndsWith(.Name, "Box") Or EndsWith(.Name, "Key") Or _
             EndsWith(.Name, "Val") Or EndsWith(.Name, "Date") Then
       .Font.Size = Range("CellFontSize_Override")
      End If
      
      If .Name = "Normal" And Range("ChangeNormalSize_Override") Then
        .Font.Size = Range("CellFontSize_Override")
      End If
      
      .IncludeFont = Range("SetsFont_Override")
      .IncludeNumber = Range("SetsFormat_Override")

    End With
  Next Item
End Sub

Public Sub ReformCell()

  With Selection
    .UnMerge
    .BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
    .Merge
    .Style = .Style
  End With

End Sub
