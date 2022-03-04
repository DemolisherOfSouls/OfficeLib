Attribute VB_Name = "IdleFunc"
Option Explicit
Option Compare Text
Option Base 1

'`Idle Timer Function Library
'Version 1.0.3

'History
' 1.0.2 - Changed to GenericIdleForm
'         Warn now triggers kick, not timer
' 1.0.3 - Changed to IdleForm
'         Added UseIdleTimer Document Switch

'Current

Public UseIdleTimer As Boolean

Private LastActive As Date
Private WarnTime   As Date
Private Stage      As String
Private IDialog    As New IdleForm
Private Closed     As Boolean

Private Const MaxIdle As Date = #12:30:00 AM# '30 minutes - Kick
Private Const PreIdle As Date = #12:20:00 AM# '20 minutes - Prompt Show

Private Sub IdleWait(ByVal Cancel As Boolean)
  Application.OnTime WarnTime, Stage, Schedule:=(Not Cancel)
End Sub

'Call This Subroutine to Activate the Idle Timer
Public Sub SetIdleTimer()
  If UseIdleTimer <> True Then Exit Sub
  LastActive = Time
  If WarnTime <> Empty Then IdleWait (True)
  Stage = "WarnExecute"
  WarnTime = LastActive + PreIdle
  IdleWait (False)
End Sub

'Call This Subroutine to Abort the prompt
Public Sub AbortKick()
  IDialog.Hide
  SetIdleTimer
End Sub

'Show Kick Warning
Private Sub WarnExecute()
  IDialog.Show False
  Stage = "KickExecute"
  WarnTime = LastActive + MaxIdle
  IdleWait False
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
  IdleWait True
  Closed = True
End Sub
