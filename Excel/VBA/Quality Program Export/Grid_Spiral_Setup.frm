VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Grid_Spiral_Setup 
   Caption         =   "ABI Inspection Program"
   ClientHeight    =   10980
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   20415
   OleObjectBlob   =   "Grid_Spiral_Setup.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Grid_Spiral_Setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = 0 Then
    Cancel = True
    MsgBox MBExitDisabled, vbCritical
  End If
End Sub

Private Sub CommandButton1_Click()
  Unload Grid_Spiral_Setup
End Sub

Private Sub UserForm_Activate()
  On Error GoTo MissingData
    
  Dim Row As Integer: Row = 7
  Dim Spec As String
  Dim YMin As String
  Dim Targ As String
  Dim Temp As Variant
  Dim YMax As String

  Spec = CalcSheet.Range("J" & Row)
  Temp = CalcSheet.Range("L" & Row)
  YMin = CStr(CalcSheet.Range("N" & Row) + Temp)
  Targ = CStr(Temp)
  YMax = CStr(CalcSheet.Range("Q" & Row) + Temp)

  For Row = 8 To 24
    Temp = CalcSheet.Range("J" & Row)
    Spec = Spec + vbNewLine + Temp
    If Temp = "Dog Leg" Or Temp = "Burrs" Or Temp = "Spiral Twist" Then
      YMin = YMin + vbNewLine + "None"
      Targ = Targ + vbNewLine + "None"
      YMax = YMax + vbNewLine + "None"
    Else
      Temp = CalcSheet.Range("L" & i)
      YMin = YMin + vbNewLine + CStr(CalcSheet.Range("N" & Row) + Temp)
      Targ = Targ + vbNewLine + CStr(Temp)
      YMax = YMax + vbNewLine + CStr(CalcSheet.Range("Q" & Row) + Temp)
    End If
  Next Row

  SpecLabel = Spec
  YMinLabel = YMin
  TargetLabel = Targ
  YMaxLabel = YMax
  OpComLabel = "[SPIRAL FORMING COMMENTS]" & vbCrLf & vbCrLf & CalcSheet.Range("Operation_Comment")
Exit Sub

MissingData:
  MsgBox MBDataMissingContact
End Sub
