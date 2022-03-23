Attribute VB_Name = "JobResultsCode"
Option Explicit
Option Compare Text
Option Base 1

Function DisplayResults(ByVal InspName As String, ByVal Columns As String, ByVal DBColumns As String)

  Dim UD10 As New ADODB.Recordset
  Dim Field_Temp As String
  Dim Temp As String
  Dim Row As Integer
  Dim NumCol As Integer: NumCol = 0
  Dim CSize As String
  Dim EndCol As String
  Dim Col As Integer
  
  ListPopSheet.Cells.Clear
  
  Field_Temp = Left(Columns, InStr(1, Columns, Chr(44)) - 1)
  Temp = Right(Columns, Len(Columns) - InStr(1, Columns, Chr(44)))
  Row = 2
  Col = 1
  While Field_Temp > ""
    EndCol = Atize(Col)
    Col = Col + 1
    ListPopSheet.Range(EndCol & Row) = Field_Temp
    If CSize > "" Then
      CSize = CSize & "," & Len(Field_Temp) * 5 + 40
    Else
      CSize = Len(Field_Temp) * 5 + 40
    End If
    If InStr(1, Temp, Chr(44)) > 0 Then
      Field_Temp = Left(Temp, InStr(1, Temp, Chr(44)) - 1)
      Temp = Right(Columns, Len(Temp) - InStr(1, Temp, Chr(44)))
    Else
      Field_Temp = Temp
      Temp = ""
    End If
  Wend
  NumCol = Col - 1
  EndCol = Atize(NumCol)
  
  With Job_Results
    With .Results
      .ColumnCount = NumCol
      .ColumnWidths = CSize
      .ColumnHeads = True
    End With
    .Lbl_Main = InspName
  End With
  
  DBEpicor.Open
  Set UD10.ActiveConnection = DBEpicor
  UD10.Open "SELECT * FROM " & "ice.UD10" & " WHERE Company = 200 AND Key1 = '" & JobNum & "' AND Key2 LIKE '" & InspName & "%" & "'"

  If UD10.EOF Or UD10.BOF Then
    MsgBox "No Results Exist for This Job"
    Exit Function
  End If
  
  UD10.MoveFirst
  Row = 3
  While UD10.EOF = False
    Field_Temp = Left(DBColumns, InStr(1, DBColumns, Chr(44)) - 1)
    Temp = Right(DBColumns, Len(DBColumns) - Len(Field_Temp) - 1)
    Col = 1
    While Field_Temp > ""
      ListPopSheet.Range(Atize(Col) & Row) = Replace(UD10.Fields(Field_Temp), InspName, "")
      If InStr(1, Temp, Chr(44)) > 0 Then
        Field_Temp = Left(Temp, InStr(1, Temp, Chr(44)) - 1)
        Temp = Right(Temp, Len(Temp) - Len(Field_Temp) - 1)
      Else
        Field_Temp = Temp
        Temp = ""
      End If
      Col = Col + 1
    Wend
    UD10.MoveNext
    Row = Row + 1
  Wend
  
  DBEpicor.Close
  
  Row = Row - 1
  Job_Results.Results.RowSource = "'ListPop'!A3:" & EndCol & Row
  Job_Results.Show
  
End Function
