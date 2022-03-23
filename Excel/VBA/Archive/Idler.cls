VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Idler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
Option Base 1

Const MaxIdle As Date = #12:30:00 AM# '30 minutes
Const PreIdle As Date = #12:20:00 AM# '20 minutes

Public LastActive As Date
Static TimeOut As Date
Static TimePre As Date

Private Sub CloseSave()

  Application.DisplayAlerts = False
  ThisWorkbook.Close SaveChanges:=True
  Application.DisplayAlerts = True

End Sub

Private Sub NotIdle()

  LastActive = Now

End Sub

Private Sub LogOpen()
    
  Console.Log ("WARNING CLOSING IN 10 MINUTES")

End Sub

Private Sub SetEventTimer()

  Application.OnTime EarliestTime:=TimeOut, Procedure:="CheckIdle"

End Sub
 
Public Sub CheckIdle()
  
  TimeOut = LastActive + MaxIdle
  Warn = LastActive + PreIdle < Now
  
  If TimeOut < Now Then CloseSave
  ElseIf Warn Then LogOpen
  
  SetEventTimer
    
End Sub

Public Property Get TRemains() As Date
  TRemains = Now - LastActive
End Property
 
Private Sub Idler_Initialize()
  NotIdle
  CheckIdle
End Sub
 
Private Sub Idler_Terminate()
  Set Me = Nothing
End Sub
