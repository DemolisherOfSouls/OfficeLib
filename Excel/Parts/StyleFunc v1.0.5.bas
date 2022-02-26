Attribute VB_Name = "StyleFunc"
Option Explicit
Option Compare Text
Option Base 1

'Style Function Library
'Version 1.0.5

Private Const DefStyle = "Normal"
Private Const DefTitleText = "Added Title"
Private ActiveFontSet As FontSet

Public Type FontSet
  Body As String
  Mono As String
  Cond As String
  Head As String
  BSize As Integer
  HSize As Integer
  TSize As Integer
  SetFont As Boolean
  SetForm As Boolean
  ChgDef As Boolean
End Type

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

Public Sub DeacColumn()
  TableColumn "Deac", Selection.Range
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
  Dim Sel As Style, STyp As String, HTyp As String, BTyp As String
  Set Sel = Selection.Style
  
  If Sel Like "Act*" Or Sel = DefStyle Then Exit Sub
  If Sel Like "Calc*" Or Sel Like "Deac*" Then
    STyp = Left(Sel, 4)
  Else
    STyp = Left(Sel, 3)
  End If
  
  If Sel Like "*HdKey" Then
    HTyp = "HdKey"
    BTyp = "Key"
  ElseIf Sel Like "*Hd" Then
    HTyp = "Hd"
    BTyp = "Cell"
  ElseIf Sel Like "*Cell" Or Sel Like "*Date" Then
    HTyp = "Hd"
    BTyp = Right(Sel, 4)
  ElseIf Sel Like "*Val" Then
    HTyp = "Hd"
    BTyp = "Val"
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

'Leave Percent, Normal, Hyperlink, Followed Hyperlink
Private Sub RemoveAllDefStyles()
  Dim Item As Style
  For Each Item In ActiveWorkbook.Styles: With Item
    If .Name Like "*Accent*" Or _
       .Name Like "Heading*" Or _
       .Name Like "*put" Or _
       .Name Like "Curr*" Or _
       (.Name Like "* *" And Not .Name Like "*Hyperlink*") Or _
       .Name Like "Comm*" Or _
       .Name = "Title" Or _
       .Name = "Total" Or .Name = "Good" Or .Name = "Title" Or .Name = "Bad" Or .Name = "Neutral" _
        Then .Delete
    End With
  Next Item
End Sub

Private Sub LoadFontSet()
  With ActiveFontSet
    .Head = ListSheet.Range("FontTable[Head]")
    .Body = ListSheet.Range("FontTable[Body]")
    .Mono = ListSheet.Range("FontTable[Mono]")
    .Cond = ListSheet.Range("FontTable[Cond]")
    .BSize = ListSheet.Range("SizeOverrideTable[BSize]")
    .HSize = ListSheet.Range("SizeOverrideTable[HSize]")
    .TSize = ListSheet.Range("SizeOverrideTable[TSize]")
    .SetFont = ListSheet.Range("SizeOverrideTable[SetFont]")
    .SetForm = ListSheet.Range("SizeOverrideTable[SetForm]")
    .ChgDef = ListSheet.Range("SizeOverrideTable[ChgDef]")
  End With
End Sub

Public Sub UpdateStyles()
  RemoveAllDefStyles
  LoadFontSet

  'All Styles except Normal
  Dim Item As Style: For Each Item In ActiveWorkbook.Styles: With Item
    If StartsWith(.Name, "Act") Then
      .Font.Size = ActiveFontSet.BSize
      .Font.Name = ActiveFontSet.Body
    ElseIf EndsWith(.Name, "Title") Then
      .Font.Size = ActiveFontSet.TSize
      .Font.Name = ActiveFontSet.Head
    ElseIf EndsWithAny(.Name, "Hd", "HdKey", "Head") Then
      .Font.Size = ActiveFontSet.HSize
      .Font.Name = ActiveFontSet.Head
    ElseIf EndsWithAny(.Name, "Val", "Date") Then
      .Font.Size = ActiveFontSet.BSize
      .Font.Name = ActiveFontSet.Mono
    ElseIf .Name <> DefStyle Then
      .Font.Size = ActiveFontSet.BSize
      .Font.Name = ActiveFontSet.Body
    End If
    
    'Box and Title Styles
    If StartsWith(.Name, "Box") Then
      .IncludeAlignment = True
    Else
      .IncludeAlignment = False
    End If
    
    'Normal Style
    If .Name = DefStyle And ActiveFontSet.ChgDef Then
      .Font.Name = ActiveFontSet.Body
      .Font.Size = ActiveFontSet.BSize
    End If
    
    'Font Setter Styles
    If StartsWith(.Name, "x") Then
      .Font.Name = Range("FontTable[" & Right(.Name, 4) & "]")
    Else
      .IncludeFont = ActiveFontSet.SetFont
    End If
    
    'Format Setter Styles
    If EndsWithAny(.Name, "Date", "Percent") Then
      .IncludeNumber = True
    Else
      .IncludeNumber = ActiveFontSet.SetForm
    End If
    
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
