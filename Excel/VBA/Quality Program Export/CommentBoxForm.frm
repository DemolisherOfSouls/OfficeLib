VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CommentBoxForm 
   Caption         =   "ABI Inspection Program"
   ClientHeight    =   10980
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11460
   OleObjectBlob   =   "CommentBoxForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CommentBoxForm"
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

Private Sub Submit_Click()
  CommentBoxReturn = Answer_Box
  Unload CommentBoxForm
End Sub

Private Sub UserForm_Activate()
  Answer_Box.SetFocus
End Sub

