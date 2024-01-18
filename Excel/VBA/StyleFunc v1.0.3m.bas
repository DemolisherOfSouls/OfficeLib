Attribute VB_Name = "StyleFunc"
Option Explicit
Option Compare Text
Option Base 1

'`Style Function Library
'Version 1.0.3m

Private Const DefStyle = "Normal"
Private Const DefTitleText = "Added Title"

Private Sub TableColumn(ByVal Typ As String, ByRef Ref As Range, Optional ByVal Body As String = "Cell", Optional ByVal Head As String = "Hd")
  Dim Bd: Set Bd = Intersect(Ref.ListObject.DataBodyRange, Ref.EntireColumn)
  Dim Hd: Set Hd = Intersect(Ref.ListObject.HeaderRowRange, Ref.EntireColumn)
  Bd.Style = Typ + Body
  Hd.Style = Typ + Head
End Sub

Public Sub LookupColumn()
  TableColumn "Lkp", Selection.Range
End Sub

Public Sub CalcColumn()
  TableColumn "Calc", Selection.Range
End Sub

Public Sub InputColumn()
  TableColumn "Inp", Selection.Range
End Sub

Public Sub InternalColumn()
  TableColumn "Int", Selection.Range
End Sub

Public Sub ErrorColumn()
  TableColumn "Err", Selection.Range
End Sub

Public Sub QueryColumn()
  TableColumn "Que", Selection.Range
End Sub

Public Sub FixColumn()
  Dim Sel As Style, STyp, HTyp, BTyp
  Set Sel = Selection.Style
  
  If Sel Like "Act*" Or Sel = DefStyle Then Exit Sub
  If Sel Like "Calc*" Then
    STyp = Left(Sel, 4)
  Else
    STyp = Left(Sel, 3)
  End If
  
  If Sel Like "*Hd" Then
    HTyp = "Hd"
    BTyp = "Cell"
  ElseIf Sel Like "*Key" Then
    HTyp = "HdKey"
    BTyp = "Key"
  ElseIf Sel Like "*Cell" Or Sel Like "*Date" Then
    HTyp = "Hd"
    BTyp = Right(Sel, 4)
  ElseIf Sel Like "*Val" Then
    HTyp = "Hd"
    BTyp = Right(Sel, 3)
  End If
  
  TableColumn STyp, Selection.Range, BTyp, HTyp
  
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

    If .Name Like "*Accent*" Or _
       .Name Like "Heading*" Or _
       .Name Like "*put" Or _
       .Name Like "Curr*" Or _
       (.Name Like "* *" And Not .Name Like "*Link*") Or _
       .Name Like "Comm*" _
        Then GoTo Delete
        
    GoTo Skip
Delete:
    Item.Delete
Skip:
  Next Item
End Sub

Public Sub UpdateStyles()

  RemoveAllDefStyles

  Dim Item As Style: For Each Item In ActiveWorkbook.Styles: With Item
    
    Dim ChgNorm As Boolean: ChgNorm = Range("ChangeNormalSize_Override")
    Dim ChgFont As Boolean: ChgFont = Range("SetsFont_Override")
    Dim ChgForm As Boolean: ChgForm = Range("SetsFormat_Override")
    Dim BdySize As Integer: BdySize = Range("CellFontSize_Override")
    Dim BdyFont As String:  BdyFont = Range("FontTable[Body]")
    
    If EndsWith(.Name, "Title") Then
      .Font.Size = Range("TitleFontSize_Override")
      .Font.Name = Range("FontTable[Head]")
    ElseIf EndsWithAny(.Name, "Hd", "HdKey") Then
      .Font.Size = Range("HeaderFontSize_Override")
      .Font.Name = Range("FontTable[Head]")
    ElseIf EndsWithAny(.Name, "Cell", "Box", "Key") Then
      .Font.Size = BdySize
      .Font.Name = BdyFont
    ElseIf EndsWithAny(.Name, "Val", "Date") Then
      .Font.Size = BdySize
      .Font.Name = Range("FontTable[Mono]")
    End If
    
    If .Name = "Normal" And ChgNorm Then .Font.Size = BdySize
    If .Name = "Normal" And ChgNorm Then .Font.Name = BdyFont
    
    If .Name = "xCond" Then .Font.Name = Range("FontTable[Cond]")
    If .Name = "xMono" Then .Font.Name = Range("FontTable[Mono]")
    
    .IncludeFont = ChgFont
    .IncludeNumber = ChgForm
  End With: Next Item
End Sub

Public Sub ReformCell()

  With Selection
    .UnMerge
    .BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
    .Merge
    .Style = .Style
  End With

End Sub
