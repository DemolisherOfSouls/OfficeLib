VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Weaving_Spiral_Setup 
   Caption         =   "ABI Inspection Program"
   ClientHeight    =   10980
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   20415
   OleObjectBlob   =   "Weaving_Spiral_Setup.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Weaving_Spiral_Setup"
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

Private Sub CloseButton_Click()
  Unload Weaving_Spiral_Setup
End Sub

Private Sub UserForm_Activate()
  On Error GoTo MissingData
    
  Dim Row As Integer: Row = 51
  Dim Spec As String
  Dim YMin As String
  Dim Targ As String
  Dim JRow, LRow, NRow, QRow
  Dim YMax As String
  
  Spec = CalcSheet.Range("J" & Row)
  LRow = CalcSheet.Range("L" & Row)
  NRow = CalcSheet.Range("N" & Row)
  QRow = CalcSheet.Range("Q" & Row)
  YMin = CStr(NRow + LRow)
  Targ = CStr(LRow)
  YMax = CStr(CalcSheet.Range("Q" & Row) + LRow)

  For Row = 52 To 55
  
    JRow = CalcSheet.Range("J" & Row)
    LRow = CalcSheet.Range("L" & Row)
    NRow = CalcSheet.Range("N" & Row)
    QRow = CalcSheet.Range("Q" & Row)
  
    Spec = Spec + vbNewLine + JRow
    If JRow = "Rod Length (Visual)" Or JRow = "Straightness" Then
      YMin = YMin + vbNewLine + "Pass"
      Targ = Targ + vbNewLine + "Pass"
      YMax = YMax + vbNewLine + "Pass"
    Else
      YMin = YMin + vbNewLine + CStr(NRow + LRow)
      Targ = Targ + vbNewLine + CStr(LRow)
      YMax = YMax + vbNewLine + CStr(QRow + LRow)
    End If
  Next Row

  Spec = Spec
  Yellow_Min = YMin
  Target = Targ
  Yellow_Max = YMax
  Operation_Comment = "[WEAVING COMMENTS]" & vbNewLine & vbNewLine & CalcSheet.Range("Operation_Comment")
Exit Sub
    
MissingData:
  MsgBox MBDataMissingContact
End Sub
