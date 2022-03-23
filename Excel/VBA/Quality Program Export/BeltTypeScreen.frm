VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BeltTypeScreen 
   Caption         =   "ABI Inspection Program"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11355
   OleObjectBlob   =   "BeltTypeScreen.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BeltTypeScreen"
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
    MsgBox "The X is disabled, please use a button on the form.", vbCritical
  End If
End Sub

Private Sub Submit_Click()
  CalcSheet.Range("LPartNum") = Input_Belt_Type
  BeltType = Input_Belt_Type
  BeltTypeScreen.Hide
End Sub

Private Sub UserForm_Initialize()
  With Input_Belt_Type
    .AddItem "BW"
    .AddItem "CTB"
    .AddItem "FWA1"
    .AddItem "FWA2"
    .AddItem "FWA3"
    .AddItem "FWA5"
    .AddItem "FWA5S"
    .AddItem "FWA6"
    .AddItem "FWB1"
    .AddItem "FWB2"
    .AddItem "FWB3"
    .AddItem "FWB4"
    .AddItem "FWB5"
    .AddItem "FWB6"
    .AddItem "FWC1"
    .AddItem "FWC1C"
    .AddItem "FWC2C"
    .AddItem "FWC6"
    .AddItem "OFE1"
    .AddItem "OFE2"
    .AddItem "OG075"
    .AddItem "OG100"
    .AddItem "RROG100"
    .AddItem "SROFG1"
    .AddItem "SROFG3"
    .AddItem "SROG075"
    .AddItem "SROG100"
    .AddItem "SSOG075"
  End With
End Sub
