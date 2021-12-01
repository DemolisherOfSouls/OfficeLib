Attribute VB_Name = "IdleFunc"
Option Explicit
Option Compare Text
Option Base 1

'Idle Timer Function Library

Private LastActive As Date
Private WarnTime   As Date
Private KickTime   As Date
Private KickQueued As Boolean
Private IdleDialog As New IdleForm
Private Closed As Boolean

Private Const MaxIdle As Date = #12:30:00 AM# '30 minutes - Kick Time
Private Const PreIdle As Date = #12:20:00 AM# '20 minutes - Prompt Show
Private Const NxtChkT As Date = #12:05:00 AM# '5 minutes  - Next Check

'Call This Subroutine to Activate the Idle Timer
Public Sub SetIdleTimer()
  LastActive = Time
  CheckIdle
End Sub

'Call This Subroutine to Abort the prompt
'  Reset the Timer if true (default)
'  Stop the idle timer if false
Public Sub AbortKick(Optional ByVal temp As Boolean = True)
  IdleDialog.Hide
  Call Application.OnTime(KickTime, "KickExecute", Schedule:=False)
  If (temp) Then SetIdleTimer
End Sub


Private Sub CheckIdle()
  If Closed Then Exit Sub
  
  Dim TimeOut As Date
  Dim ChkTime As Date
  Dim Warn As Boolean
  WarnTime = LastActive + PreIdle
  TimeOut = LastActive + MaxIdle
  ChkTime = Time + NxtChkT
  Warn = WarnTime < Time
  
  If Warn Then
    KickQueued = True
    KickTime = TimeOut
    Call Application.OnTime(KickTime, "KickExecute")
    IdleDialog.Show (False)
  Else
    Call Application.OnTime(ChkTime, "CheckIdle")
  End If
  
End Sub

Private Sub KickExecute()
  On Error GoTo Quit
  If Closed Then Exit Sub
  IdleDialog.Hide
  
  With Application
    .DisplayAlerts = False
    .ThisWorkbook.Close (True)
    .DisplayAlerts = True
    .Quit
  End With
  
  Closed = True
  
Quit:

End Sub
