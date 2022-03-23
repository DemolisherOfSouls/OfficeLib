Attribute VB_Name = "ButtonCode"
Option Explicit
Option Compare Text
Option Base 1

Public Sub RefreshButton_Click()
    
    'Timestamp
    With MachineSheet
    
        .Range("RefDate").Value = Date
        .Range("RefTime").Value = Time
        
    End With
    
    'Schedule Tab Filter
    With ScheduleSheet.ListObjects("ScheduleTable")
    
        .Range.AutoFilter Field:=1, Criteria1:=">=0"
        .ShowAutoFilterDropDown = False
        
    End With

    'Refresh Data / Supress Deletion Alerts
    With Application
    
        .ActiveWorkbook.RefreshAll
        .DisplayAlerts = False
        
    End With
    
    'Delete Generated Sheets
    Dim sh As Worksheet
    For Each sh In Worksheets
    
        If STARTSWITH(sh.Name, "Sheet") Then
          sh.Delete
        End If
        
    Next
    
    'Resume Deletion Alerts
    Application.DisplayAlerts = True
    
End Sub

Sub WeekComboBox_Change()

    'Filter Forecast ComboBox
    Range("TD").AutoFilter Field:=1, Criteria1:=Range("FCWeeks").Value
    Range("CD").AutoFilter Field:=1, Criteria1:=Range("FCWeeks").Value

End Sub


Sub ToggleSelection_MS()

    
    
    

End Sub
