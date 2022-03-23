VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Weaving_Crimp_Setup 
   Caption         =   "ABI Inspection Program"
   ClientHeight    =   10980
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   20415
   OleObjectBlob   =   "Weaving_Crimp_Setup.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Weaving_Crimp_Setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SpecText_Click()

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        MsgBox MBExitDisabled, vbCritical
    End If
End Sub

Private Sub CloseButton_Click()
  Unload Crimp_Setup
End Sub

Private Sub UserForm_Activate()
    On Error GoTo MissingData
    
    Dim i As Integer: i = 63
    Dim Spec As String
    Dim Y_Min As String
    Dim Targ As String
    Dim Temp As Variant
    Dim Y_Max As String

    'Resize_Screen(Crimp_Setup)
    
    Spec = CalcSheet.Range("J" & i)
    Temp = CalcSheet.Range("L" & i)
    Y_Min = CStr(CalcSheet.Range("N" & i) + Temp)
    Targ = CStr(Temp)
    Y_Max = CStr(CalcSheet.Range("Q" & i) + Temp)
    i = i + 1

    While i < 67
        Spec = Spec + vbNewLine + CalcSheet.Range("J" & i)
        If CalcSheet.Range("J" & i) = "Dog Leg" Or CalcSheet.Range("J" & i) = "Burrs" Or CalcSheet.Range("J" & i) = "Spiral Twist" Then
            Y_Min = Y_Min + vbNewLine + "None"
            Targ = Targ + vbNewLine + "None"
            Y_Max = Y_Max + vbNewLine + "None"
        Else
            Temp = CalcSheet.Range("L" & i)
            Y_Min = Y_Min + vbNewLine + CStr(CalcSheet.Range("N" & i) + Temp)
            Targ = Targ + vbNewLine + CStr(Temp)
            Y_Max = Y_Max + vbNewLine + CStr(CalcSheet.Range("Q" & i) + Temp)
        End If
        i = i + 1
    Wend

    SpecText = Spec
    Yellow_Min = Y_Min
    Target = Targ
    Yellow_Max = Y_Max
    Operation_Comment = "[SPIRAL FORMING COMMENTS]" & vbCrLf & vbCrLf & CalcSheet.Range("Operation_Comment")
Exit Sub
    
MissingData:
  MsgBox MBDataMissingContact
End Sub
