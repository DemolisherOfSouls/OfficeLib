Attribute VB_Name = "PMCode"
Option Explicit
Option Compare Text
Option Base 1

Sub Refresh_Click()

    Workbooks(ThisWorkbook.Name).RefreshAll
    
    Range("RefreshDate").Value = Format(Date, "M-D-YY")

End Sub

Public Function FLAG(ByVal val As String) As String

    FLAG = Left(LCase(val), 4)

End Function
