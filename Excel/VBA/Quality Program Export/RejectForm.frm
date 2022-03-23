VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RejectForm 
   Caption         =   "ABI Inspection Program"
   ClientHeight    =   10980
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   20415
   OleObjectBlob   =   "RejectForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RejectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = 0 Then
    Cancel = True
    MsgBox "The X is disabled, please use a button on the form.", vbCritical
  End If
End Sub

Private Sub OkButton_Click()
  Unload RejectForm
End Sub
