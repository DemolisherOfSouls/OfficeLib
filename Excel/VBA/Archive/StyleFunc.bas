Attribute VB_Name = "StyleFunc"
Option Explicit
Option Compare Text
Option Base 1

'Style Function Library
'Version 0.9.0

Private Const DefStyle = "Normal"
Private Const DefTitleText = "Added Title"

Private Function SelStyle(ByRef r As Range) As String
  SelStyle = r.Cells(1).Style.Name
End Function

Private Function GetTCHeader(ByRef r As Range) As Range
  GetTCHeader = Intersect(r.EntireColumn, r.ListObject.HeaderRowRange)
End Function

Private Function GetTCBody(ByRef r As Range) As Range
  GetTCBody = Intersect(r.EntireColumn, r.ListObject.DataBodyRange)
End Function

Private Sub TableColumn(ByVal Typ As String, ByRef Ref As Range, Optional ByVal Body As String = "Cell", Optional ByVal Head As String = "Hd")
  GetTCBody(Ref).Style = Typ + Body
  GetTCHeader(Ref).Style = Typ + Head
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
  
  TableColumn (STyp & HTyp), (STyp & BTyp), Selection
End Sub

Public Sub AddTitle()
  Dim I As Range: Set I = Intersect(Selection.Resize(1).Offset(-1, 0).EntireRow, Selection.EntireColumn)
  With I
    .EntireRow.Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromLeftOrAbove
    .Style = "BoxTitle"
    .Merge
    .Value = DefTitleText
  End With
End Sub

Public Sub UpdateStyles()

  Dim Item: For Each Item In ActiveWorkbook.Styles
    
    With Item
      If EndsWith(.Name, "Title") Then
        Let .Font.Size = Range("TitleFontSize_Override")
      ElseIf EndsWith(.Name, "Hd") Or EndsWith(.Name, "HdKey") Then
        Let .Font.Size = Range("HeaderFontSize_Override")
      ElseIf EndsWith(.Name, "Cell") Or EndsWith(.Name, "Box") Or EndsWith(.Name, "Key") Or _
             EndsWith(.Name, "Val") Or EndsWith(.Name, "Date") Then
       Let .Font.Size = Range("CellFontSize_Override")
      End If
      
      If .Name = "Normal" And Range("ChangeNormalSize_Override").Value Then
        Let .Font.Size = Range("CellFontSize_Override")
      End If
      
      Let .IncludeFont = Range("SetsFont_Override")
      Let .IncludeNumber = Range("SetsFormat_Override")
    End With
  
  Next Item

End Sub

Public Sub ReformCell()

  With Selection
    .UnMerge
    .BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
    .Merge
  End With

End Sub
