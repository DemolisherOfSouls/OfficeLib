VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GenericIdleForm 
   Caption         =   "Idle Timer"
   ClientHeight    =   1575
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   4365
   OleObjectBlob   =   "GenericIdleForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GenericIdleForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DismissButton_Click()

  AbortKick

End Sub
