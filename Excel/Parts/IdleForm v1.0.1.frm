VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} IdleForm 
   Caption         =   "Idle Timer"
   ClientHeight    =   1575
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   4365
   OleObjectBlob   =   "IdleForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "IdleForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Compare Text
Option Base 1

'`IdleForm Form Library
' Version - 1.0.1

' Current

Private Sub DismissButton_Click()
  AbortKick
End Sub
