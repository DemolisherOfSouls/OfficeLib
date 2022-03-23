VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Weaving_CB5_Setup 
   Caption         =   "ABI Inspection Program"
   ClientHeight    =   10980
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   20415
   OleObjectBlob   =   "Weaving_CB5_Setup.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Weaving_CB5_Setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SpecText_Click()

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = 0 Then
    Cancel = True
    MsgBox "The X is disabled, please use a button on the form.", vbCritical
  End If
End Sub

Private Sub CommandButton1_Click()
  Unload CB5_Weave_Setup
End Sub

Private Sub UserForm_Activate()
  On Error GoTo MissingData
    
    Dim Row As Integer: Row = 74
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

    For Row = 75 To 77
        Spec = Spec + vbNewLine + CalcSheet.Range("J" & Row)
        If CalcSheet.Range("J" & Row) = "Rod Length (Visual)" Or CalcSheet.Range("J" & Row) = "Straightness" Then
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
    
    Spec = Spec + vbNewLine + "Fabric Width" + vbNewLine + "Roll"
    YMin = YMin + vbNewLine + "Pass" + vbNewLine + "Pass"
    Targ = Targ + vbNewLine + "Pass" + vbNewLine + "Pass"
    YMax = YMax + vbNewLine + "Pass" + vbNewLine + "Pass"

    SpecText = Spec
    Yellow_Min = YMin
    Target = Targ
    Yellow_Max = YMax
    Operation_Comment = "[WEAVING COMMENTS]" & vbNewLine & vbNewLine & CalcSheet.Range("Operation_Comment")
Exit Sub
    
MissingData:
  MsgBox MBDataMissingContact
End Sub
