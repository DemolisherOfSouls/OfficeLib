VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Grid_Welding_Inspection 
   Caption         =   "ABI Inspection Program"
   ClientHeight    =   8445.001
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7035
   OleObjectBlob   =   "Grid_Welding_Inspection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Grid_Welding_Inspection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1

Private Sub Inspection_Num_Click()

End Sub

Private Sub UserForm_QueryClose(ByRef Cancel As Integer, CloseMode As Integer)
  If CloseMode = 0 Then
    Cancel = True
    MsgBox MBExitDisabled, vbCritical
  End If
End Sub

Private Sub Home_Click()
  Unload Grid_Welding_Inspection
  StartScreen.Show
End Sub

Private Sub Submit_Click()
  On Error GoTo GeneralError
  
  With CalcSheet
    Const InspName As String = "Grid_Welding_Inspection"
    Dim InspPlan As String: InspPlan = .Range(InspName & "_Plan")
    Dim InspSpec As String: InspSpec = .Range(InspName & "_Spec")
    Dim InspComm As String: InspComm = .Range(InspName & "_Comment")
      
    ClearCalcSheet Written:=True
  
    If Not IsNumeric(LengthTextBox) And Not IsEmpty(LengthTextBox) Then
      GoTo GeneralError
    End If
      
    If Not (PassOption Or FailOption) Or IsEmpty(LengthTextBox) Then
      GoTo BlankForm
    End If
  
    .Range("Insp_Plan") = InspPlan
    .Range("Spec_ID") = InspSpec
    .Range("Data1") = TryParseFraction(LengthTextBox)
    .Range("Check2") = Radio(PassOption, FailOption)
  
    If InspComm = Empty Then
      .Range("Passed") = 1
      .Range("Value") = Empty
      .Range("Failed_Comment") = Empty
    Else
      .Range("Passed") = 0
      .Range("Value") = "Weld Rejected"
      .Range("Failed_Comment") = Replace(InspComm, "?", ".  ")
      RejectForm Replace(InspComm, "?", vbNewLine)
    End If
  End With

  WriteToSQL "ice.UD10"
  
  PassOption = False
  FailOption = False
  LengthTextBox = ""
  
  SetNextSampleNum
  Inspection_Num = "Inspection Num: " & CStr(SampleNum)
    
Exit Sub

GeneralError:
  MsgBox MBDataErrorsResbmit
Exit Sub

BlankForm:
  MsgBox MBFillOutResubmit
End Sub

Private Sub UserForm_Activate()

  ClearCalcSheet Written:=False

  SetNextSampleNum
  Inspection_Num = "Inspection Num: " & CStr(SampleNum)
End Sub
