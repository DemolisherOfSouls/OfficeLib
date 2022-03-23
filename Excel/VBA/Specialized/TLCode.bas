Attribute VB_Name = "TLCode"
Option Explicit
Option Compare Text
Option Base 1

Public Origin As Range
Public StartDate As Range
Public EndDate As Range
Public Resource As Range
Public Company As Range
Public Resolution As Range
Public TLabels As Range

Private _
  DataO As ListObject, _
  StCol As Range, _
  EnCol As Range, _
  JbCol As Range, _
  TArea As Range, _
  FLabl As Range, _
  TLabl As Range, _
  CompR As Range

'Refresh Query
Sub Refresh()

  ActiveWorkbook.RefreshAll

End Sub

'Return biggest
Function Max(ByVal x, ByVal y)

  Max = IIf(x > y, x, y)

End Function

'Return unbiggest
Function Min(ByVal x, ByVal y)

  Min = IIf(x < y, x, y)

End Function

'Reset
Sub Setup()

  Set Origin = Range("Origin")
  Set StartDate = Range("StartDate")
  Set Resolution = Range("Resolution")
  Set TLabl = Range("TimelineLabels")
  Set FLabl = Range("FLabel")
  
  Set DataO = DataSheet.ListObjects(1)
  Set JbCol = DataO.ListColumns(1).DataBodyRange
  Set StCol = DataO.ListColumns(3).DataBodyRange
  Set EnCol = DataO.ListColumns(4).DataBodyRange
  
  TLabl.Font.Color = vbBlack
  FLabl = Max(StCol.Value2(1, 1), StartDate)
  
  With Range("TimelineArea")
    .ClearFormats
    .ClearContents
    .ColumnWidth = 2
  End With

End Sub

'I hate this
Function Flip(Optional ByVal silent = False) As Byte
  
  Static var As Byte
  
  If var = 0 And Not silent Then
    var = 1
  ElseIf var = 1 And Not silent Then
    var = 0
  End If
  
  Flip = var

End Function

Sub MakeBlock(ByRef cells As Range, ByVal job As String)

  With cells
    
    If (.Row Mod 2 = 1) Then
      .Style = "QueKey"
    Else
      .Style = "LkpKey"
    End If
    
    .Merge
    .Value = job
    .Font.Size = 11
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    
  End With

End Sub

'Refresh Timeline
Sub Timeline()

  Setup
  
  Dim _
    RIndex As Range, _
    LastDX As Integer, _
    CellDX As Integer, _
    SizeDX As Integer, _
    ZeroStack As Integer, _
    JobRow As ListRow, _
    LastDt As Date
  
  For Each JobRow In DataO.ListRows
  
    Dim _
      StDt As Date, _
      EnDt As Date, _
      Diff As Double, _
      JobS As String, _
      Row As Range
    
    Set Row = JobRow.Range
    StDt = Max(Intersect(Row, StCol), StartDate)
    EnDt = Intersect(Row, EnCol)
    JobS = Intersect(Row, JbCol)
    
    If (LastDt <> 0) Then
    
      Diff = CDbl(StDt) - CDbl(LastDt)
      CellDX = CellDX + CInt(Diff * Resolution)
      If CellDX > 0 Then ZeroStack = 0
    
    End If
    
    Diff = CDbl(EnDt) - CDbl(StDt)
    
    LastDX = CellDX
    SizeDX = Diff * Resolution
    CellDX = CellDX + SizeDX
    
    If (LastDX = 0 And Flip(True) = 1) Then Flip
    
    If SizeDX <= 1 Then
      Set RIndex = Origin.Offset(1 + ZeroStack, LastDX)
      Set RIndex = RIndex.Resize(1, 1)
      RIndex.ColumnWidth = 8
      ZeroStack = ZeroStack + 1
    Else
      Set RIndex = Origin.Offset(Flip, LastDX)
      Set RIndex = RIndex.Resize(1, SizeDX)
      RIndex.ColumnWidth = Max(8# / SizeDX, 2)
      ZeroStack = 0
    End If
    
    MakeBlock(RIndex, JobS)
    
    LastDt = EnDt

  Next
  
  'Hide Excess Timeline
  RIndex.Offset(-5 - Flip, 1).Resize(1, 800).Font.Color = vbWhite
  
End Sub
