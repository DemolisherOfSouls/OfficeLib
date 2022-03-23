Attribute VB_Name = "IdleFunc"
Option Explicit
Option Compare Text
Option Base 1

'Idle Timer Function Library
'Version 1.0.1m

'Current

Private LastActive As Date
Private WarnTime   As Date
Private KickTime   As Date
Private IDialog    As New IdleForm
Private Closed     As Boolean

Private Const MaxIdle As Date = #12:30:00 AM# '30 minutes - Kick Time
Private Const PreIdle As Date = #12:20:00 AM# '20 minutes - Prompt Show

'Call This Subroutine to Activate the Idle Timer
Public Sub SetIdleTimer()
  LastActive = Time
  CancelTimer
  WarnTime = LastActive + PreIdle
  KickTime = LastActive + MaxIdle
  Application.OnTime KickTime, "KickExecute"
  Application.OnTime WarnTime, "WarnExecute"
End Sub

'Cancels the Timer
Public Sub CancelTimer(Optional ByVal final As Boolean = False)
  If WarnTime <> Empty Then
    Application.OnTime KickTime, "KickExecute", Schedule:=False
    Application.OnTime WarnTime, "WarnExecute", Schedule:=False
  End If
  If final Then Closed = True
End Sub

'Call This Subroutine to Abort the prompt
'  Reset the Timer if true (default)
'  Stop the idle timer if false
Public Sub AbortKick(Optional ByVal reset As Boolean = True)
  IDialog.Hide
  If reset Then
    SetIdleTimer
  Else
    CancelTimer
  End If
End Sub

'Show Kick Warning
Private Sub WarnExecute()
  IDialog.Show False
End Sub

'Execute Kick
Private Sub KickExecute()
  If Closed Then Exit Sub
  IDialog.Hide
  With Application
    .DisplayAlerts = False
    .ThisWorkbook.Close True
    .DisplayAlerts = True
    .Quit
  End With
  CancelTimer True
End Sub
