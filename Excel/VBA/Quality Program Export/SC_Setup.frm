VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SC_Setup 
   Caption         =   "ABI Inspection Program"
   ClientHeight    =   10980
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   20415
   OleObjectBlob   =   "SC_Setup.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SC_Setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = 0 Then
    Cancel = True
    MsgBox MBExitDisabled, vbCritical
  End If
End Sub

Private Sub CommandButton1_Click()
  Unload S_C_Setup
End Sub

Private Sub UserForm_Activate()
  On Error GoTo MissingData

  Dim Row As Integer: Row = 32
  Dim Spec As String
  Dim YMin As String
  Dim Targ As String
  Dim Temp As Variant
  Dim YMax As String
  
  Spec = CalcSheet.Range("J" & i)
  Temp = CalcSheet.Range("L" & i)
  Y_Min = CStr(CalcSheet.Range("N" & i) + Temp)
  Targ = CStr(Temp)
  Y_Max = CStr(CalcSheet.Range("Q" & i) + Temp)

  For Row = 33 To 35
    Spec = Spec + vbNewLine + CalcSheet.Range("J" & i)
    If CalcSheet.Range("J" & i) = "Rod Length (Visual)" Or CalcSheet.Range("J" & i) = "Straightness" Then
      Y_Min = Y_Min + vbNewLine + "Pass"
      Targ = Targ + vbNewLine + "Pass"
      Y_Max = Y_Max + vbNewLine + "Pass"
    Else
      Temp = CalcSheet.Range("L" & i)
      Y_Min = Y_Min + vbNewLine + CStr(CalcSheet.Range("N" & i) + Temp)
      Targ = Targ + vbNewLine + CStr(Temp)
      Y_Max = Y_Max + vbNewLine + CStr(CalcSheet.Range("Q" & i) + Temp)
    End If
  Next Row

  SpecText = Spec
  Yellow_Min = YMin
  Target = Targ
  Yellow_Max = YMax
  Operation_Comment = "[STRAIGHT AND CUT COMMENTS]" & vbNewLine & vbNewLine & CalcSheet.Range("Operation_Comment")
Exit Sub
    
MissingData:
  MsgBox MBDataMissingContact
End Sub
