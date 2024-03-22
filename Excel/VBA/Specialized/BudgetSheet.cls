VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BudgetSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
Option Base 1

'Budget Table Columns
Const BT_TRowCol = 1
Const BT_DateCol = 2
Const BT_TDayCol = 3
Const BT_CodeCol = 4
Const BT_DTORCol = 5
Const BT_ATORCol = 6
Const BT_AmntCol = 7
Const BT_DescCol = 8
Const BT_TotlCol = 9
Const BT_ThrsCol = 10

Const BT_FirstCell = "A11"

Const BT_EndDummyDate = #12/31/2025#

'Savings Table Columns
Const ST_DateCol = 1
Const ST_AmntCol = 2
Const ST_DescCol = 3
Const ST_GoalCol = 4
Const ST_DiscCol = 5
Const ST_GTotCol = 6
Const ST_TotlCol = 7
Const ST_PercCol = 8
Const ST_DTAdCol = 9

Const ST_FirstCell = "N11"

Public Function GetBT(ByVal col As Integer, ByVal row As Integer)
  Dim thiscell As Range
  Set thiscell = Range(BT_FirstCell).Offset(row - 1, col - 1)
  GetBT = thiscell.Value
End Function

Public Function GetST(ByVal col As Integer, ByVal row As Integer)
  Dim thiscell As Range
  Set thiscell = Range(ST_FirstCell).Offset(row - 1, col - 1)
  GetST = thiscell.Value
End Function

Public Function NextBTRow() As Integer
  NextBTRow = Range("BudgetRowCount").Value + 1
End Function

Public Function NextSTRow() As Integer
  NextSTRow = Range("SavingsTable[Date]").count + 1
End Function

Public Sub SetBT(ByVal col As Integer, ByVal row As Integer, ByVal val)
  Range(BT_FirstCell).Offset(row - 1, col - 1).Value = val
End Sub

Public Sub SetST(ByVal col As Integer, ByVal row As Integer, ByVal val)
  Range(ST_FirstCell).Offset(row - 1, col - 1).Value = val
End Sub

Public Sub AddSTRow(ByVal transdate)

  Dim addedSRow As Integer
  addedSRow = NextSTRow()
  
  Call SetST(ST_DateCol, addedSRow, CLng(transdate))
  Call SetST(ST_DescCol, addedSRow, "Transfer")
  Call SetST(ST_GoalCol, addedSRow, "Cushion")

End Sub

Public Sub DeleteBTRow(ByVal row As Integer)
  BudgetSheet.ListObjects("BudgetTable").ListRows(row).Delete
End Sub

Public Sub DeleteSTRow(ByVal row As Integer)
  BudgetSheet.ListObjects("SavingsTable").ListRows(row).Delete
End Sub


Public Sub SortBT()
  With BudgetSheet.ListObjects("BudgetTable").Sort
    .SortFields.Clear
    .SortFields.Add Range("BudgetTable[Date]")
    .Header = xlYes
    .Apply
  End With
End Sub

Private Sub BudgetFilterButton_Click()

  Dim table As ListObject
  Set table = BudgetSheet.ListObjects("BudgetTable")
  
  If table.ShowAutoFilter Then
    table.ShowAutoFilter = False
  End If
  
  If Not table.AutoFilter Is Nothing Then
    table.AutoFilter.ShowAllData
  End If
  
  If (Len(DescriptionComboBox.Text) > 0) Then
    table.Range.AutoFilter Field:=BT_CodeCol, Criteria1:=DescriptionComboBox.Text
  End If
  
  'TODO: Add Date Filter
  'If (Len(DescriptionComboBox.Text) > 0) Then
  '  table.Range.AutoFilter Field:=BT_DescCol, Criteria1:=DescriptionComboBox.Text
  'End If
  

End Sub

Private Sub ClearButton_Click()

  DescriptionComboBox.Text = ""
  BudgetFilterButton_Click

End Sub

Private Sub SavingsUpdate()

  Application.ScreenUpdating = False
  
  Dim TableSize As Integer
  Dim counter As Integer
  
  TableSize = Range("SavingsRowCount").Value
  
  For counter = 1 To TableSize
    If IsError(GetST(ST_AmntCol, counter)) Then DeleteSTRow (counter)
  Next
  
  Application.ScreenUpdating = True

End Sub

Private Sub UpdateButton_Click()

  Application.ScreenUpdating = False

  Dim TableSize As Integer
  Dim MaxSavingsDate
  Dim counter As Integer
  
  TableSize = Range("BudgetRowCount").Value
  MaxSavingsDate = Range("MaxSavingsDate").Value
  
  For counter = 1 To TableSize
  
    Dim ThisRowCode As String
    ThisRowCode = GetBT(BT_CodeCol, counter)
    
    'Delete if date is past and is a one-time entry
    Dim IsOneTime As Boolean
    IsOneTime = ThisRowCode = "TransExt" Or _
                ThisRowCode = "Jackie" Or _
                ThisRowCode = "CSRp175" Or _
                ThisRowCode = "CSLn175" Or _
                ThisRowCode = "Refund" Or _
                ThisRowCode = "Theater" Or _
                ThisRowCode = "Taxes"
    If CLng(GetBT(BT_DateCol, counter)) < Date And IsOneTime Then
      DeleteBTRow (counter)
    End If
    
    'Clear if date is past and not a transfer
    If CLng(GetBT(BT_DateCol, counter)) < Date And ThisRowCode <> "TransSav" Then
      Call SetBT(BT_DTORCol, counter, "")
      Call SetBT(BT_ATORCol, counter, "")
    End If
    
    'Add missing transfers to savings table
    If CLng(GetBT(BT_DateCol, counter)) > MaxSavingsDate And ThisRowCode = "TransSav" Then
      AddSTRow (GetBT(BT_DateCol, counter))
    End If
    
    'Mark if date is past and is a transfer
    If GetBT(BT_DateCol, counter) < Date And ThisRowCode = "TransSav" Then
      Call SetBT(BT_DTORCol, counter, BT_EndDummyDate)
      Call SetBT(BT_ATORCol, counter, 200)
    End If
  
  Next

  SortBT
  
  BudgetSheet.ListObjects("BudgetTable").Range.Calculate
  
  For counter = 1 To TableSize
  
    'Reset mark on transfer
    If GetBT(BT_DTORCol, counter) = BT_EndDummyDate And GetBT(BT_CodeCol, counter) = "TransSav" Then
      Call SetBT(BT_DTORCol, counter, "")
    End If
  
  Next
  
  SortBT
  
  SavingsUpdate
  
  Application.ScreenUpdating = True

End Sub

Public Sub EndButton_Click()

  Dim TableSize As Integer
  Dim LastRow As Integer
  Dim ScrollAmount As Integer
  
  TableSize = Range("BudgetRowCount").Value
  LastRow = TableSize + 10
  ScrollAmount = TableSize - 15
  
  ActiveWindow.SmallScroll Down:=ScrollAmount
  
  Range("A" & LastRow & ":L" & LastRow).Select
  
End Sub

Private Sub AddLineButton_Click()

  Dim addedRow As Integer
  addedRow = NextBTRow()

  Call SetBT(BT_CodeCol, addedRow, DescriptionComboBox.Text)
  
  If DescriptionComboBox.Text = "TransSav" Then
    AddSTRow (GetBT(BT_DateCol, addedRow))
  End If

End Sub
