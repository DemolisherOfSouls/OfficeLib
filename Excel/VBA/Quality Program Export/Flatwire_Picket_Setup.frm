VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Flatwire_Picket_Setup 
   Caption         =   "ABI Inspection Program"
   ClientHeight    =   10980
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   20415
   OleObjectBlob   =   "Flatwire_Picket_Setup.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Flatwire_Picket_Setup"
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

Private Sub CloseButton_Click()
  Unload Flatwire_Picket_Setup
End Sub

Private Sub UserForm_Activate()
  On Error GoTo MissingData
  
  Dim Row As Integer: Row = 87
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

  For Row = 88 To 99
    Temp = CalcSheet.Range("J" & Row)
    Spec = Spec + vbNewLine + Temp
    If Temp = "Rod Length (Visual)" Or Temp = "Straightness" Then
      YMin = YMin + vbNewLine + "Pass"
      Targ = Targ + vbNewLine + "Pass"
      YMax = YMax + vbNewLine + "Pass"
    Else
      Temp = CalcSheet.Range("L" & Row)
      YMin = YMin + vbNewLine + CStr(CalcSheet.Range("N" & Row) + Temp)
      Targ = Targ + vbNewLine + CStr(Temp)
      YMax = YMax + vbNewLine + CStr(CalcSheet.Range("Q" & Row) + Temp)
    End If
  Next Row

  SpecText = Spec
  Yellow_Min = YMin
  Target = Targ
  Yellow_Max = YMax
  Operation_Comment = "[WEAVING COMMENTS]" & vbNewLine & vbNewLine & CalcSheet.Range("Operation_Comment")
Exit Sub
  
MissingData:
  MsgBox MBDataMissingContact
End Sub
