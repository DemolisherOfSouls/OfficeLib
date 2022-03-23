Attribute VB_Name = "ButtonCode"
Option Explicit
Option Compare Text
Option Base 1


Public LastActive As Date
Public WarnTime   As Date
Public KickTime   As Date
Public KickQueued As Boolean

Const MaxIdle As Date = #12:30:00 AM# '30 minutes - Kick Time
Const PreIdle As Date = #12:20:00 AM# '20 minutes - Prompt Show
Const NxtChkT As Date = #12:05:00 AM# '5 minutes  - Next Check

'Refresh Button
Sub RefreshButton_Click()
  
  'Update timestamp
  With MachineSheet
  
    .Range("RefDate").Value = Date
    .Range("RefTime").Value = Time
      
  End With
  
  'Filter out past weeks
  Dim filter As Variant
  filter = SetupSheet.ListObjects("WkTable").Range.Value
  With ScheduleSheet.ListObjects("ScheduleTable")
  
    .Range.AutoFilter Field:=1, Criteria1:=filter
    .ShowAutoFilterDropDown = False
  
  End With

  'Refresh tables
  With Application
  
    .DisplayAlerts = False
    .StatusBar = "Refreshing Content"
    .ThisWorkbook.RefreshAll
    .StatusBar = "Removing Temporary Sheets"
  
  End With
  
  'Delete temp sheets
  Dim Sh As Integer
  For (Sh = 1 To 10)
    If Worksheets.Count > 4 Then
      Worksheets("Sheet" & Sh).Delete
    End If
  Next
  
  'Reset settings
  With Application
  
    .DisplayAlerts = True
    .StatusBar = False
  
  End With
  
  SetIdleTimer
    
End Sub

'Forecast Weeks
Public Sub WeekComboBox_Change()

  Dim filter As Variant
  filter = Range("FCWeeks").Value
  Range("TD").AutoFilter Field:=1, Criteria1:=Range("FCWeeks").Value
  Range("CD").AutoFilter Field:=1, Criteria1:=Range("FCWeeks").Valu

End Sub

Public Sub SetIdleTimer()

  LastActive = Time
  CheckIdle

End Sub


'Set Timer for kick
Public Sub CheckIdle()
  
  Dim TimeOut As Date
  Dim ChkTime As Date
  Dim Warn As Boolean
  WarnTime = LastActive + PreIdle
  TimeOut = LastActive + MaxIdle
  ChkTime = Time + NxtChkT
  Warn = WarnTime < Time
  
  
  If Warn Then
    Set IdleDialog = New IdleForm
    KickQueued = True
    KickTime = TimeOut
    Application.OnTime EarliestTime:=KickTime, Procedure:="KickExecute"
    IdleDialog.Show
  Else
    Application.OnTime EarliestTime:=ChkTime, Procedure:="CheckIdle"
  End If
  
End Sub

Public Sub KickExecute()

    IdleDialog.Hide
    Application.DisplayAlerts = False
    ThisWorkbook.Close SaveChanges:=True
    Application.DisplayAlerts = True
    Application.Quit

End Sub

Public Sub AbortKick()

    IdleDialog.Hide
    Application.OnTime EarliestTime:=KickTime, Procedure:="KickExecute", Schedule:=False
    SetIdleTimer

End Sub
