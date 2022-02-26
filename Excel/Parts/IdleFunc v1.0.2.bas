Attribute VB_Name = "IdleFunc"
Option Explicit
Option Compare Text
Option Base 1

'`Idle Timer Function Library
'Version 1.0.2

'History
' 1.0.2 - Changed to GenericIdleForm
'         Warn now triggers kick, not timer

'Current

Private LastActive As Date
Private WarnTime   As Date
Private Stage      As String
Private IDialog    As New GenericIdleForm
Private Closed     As Boolean

Private Const MaxIdle As Date = #12:30:00 AM# '30 minutes - Kick
Private Const PreIdle As Date = #12:20:00 AM# '20 minutes - Prompt Show

Private Sub IdleWait(ByVal cancel as Boolean)
  Application.OnTime WarnTime, Stage, Schedule:=(Not cancel)
End Sub

'Call This Subroutine to Activate the Idle Timer
Public Sub SetIdleTimer()
  LastActive = Time
  If WarnTime <> Empty Then IdleWait(true)
  Stage = "WarnExecute"
  WarnTime = LastActive + PreIdle
  IdleWait(false)
End Sub

'Call This Subroutine to Abort the prompt
'  Reset the Timer if true (default)
'  Stop the idle timer if false
Public Sub AbortKick(Optional ByVal reset As Boolean = True)
  IDialog.Hide
  If reset Then
    SetIdleTimer
  Else
    If WarnTime <> Empty Then IdleWait(true)
    If final Then Closed = True
  End If
End Sub

'Show Kick Warning
Private Sub WarnExecute()
  IDialog.Show False
  Stage = "KickExecute"
  WarnTime = LastActive + MaxIdle
  IdleWait false
End Sub

'Execute Kick
Private Sub KickExecute()
  If Closed Then Exit Sub
  IDialog.Hide
  With Application
    .DisplayAlerts = False
    .ThisWorkbook.Close True
    .DisplayAlerts = True
  End With
  IdleWait true
  Closed = True
End Sub
