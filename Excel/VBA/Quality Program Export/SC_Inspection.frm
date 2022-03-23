VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SC_Inspection 
   Caption         =   "ABI Inspection Program"
   ClientHeight    =   10980
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   20415
   OleObjectBlob   =   "SC_Inspection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SC_Inspection"
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

Private Sub Get_Results_Click()
  DisplayResults "Straight and Cut Inspection", _
    "#,Date    ,Type,Time      ,Employ,Spec       ,Part #,Machine #,Rod Length(Measured),Rod Length(Visual),Straightness,Wire Diam", _
    "Key3,Date01,Key2,ShortChar06,ShortChar02,Character07,Character02,ShortChar03,Number01,Checkbox02,Checkbox03,Number02"
End Sub

Private Sub Home_Click()
  Unload S_C_Inspection_Form
  Start_Screen.Show
End Sub

Private Sub Image1_Click()
  Validator
  S_C_Setup.Show
End Sub

Private Sub Gleen_Info()
  Dim Job_Comments As String: Job_Comments = CalcSheet.Range("JobComments")

  If BeltWidth = "" Then
    BeltWidth = Comment_Search("Width", Job_Comments, "Inches", "in.", "", "")
    If IsNumeric(BeltWidth) And BeltWidth > 0 Then
      CalcSheet.Range("BeltWidth") = BeltWidth
    Else
      BeltWidth = TryParseFraction(ShowCommentBox("Belt Width", Job_Comments))
      
      CalcSheet.Range("BeltWidth") = BeltWidth
    End If
  End If
  If CalcSheet.Range("LPartNum") = "PDCE" Then
    Rod_Diam = TryParseFraction(ShowCommentBox("Rod Diameter", Job_Comments))
  End If
End Sub

Private Sub Entry_Test()
  If IsBad(BeltWidth) Then Run_Setup
  If CalcSheet.Range("LPartNum") = "PDCE" And IsBad(Rod_Diam) Then Run_Setup
End Sub

Private Sub Validator()
  While BeltWidth = "" Or (CalcSheet.Range("LPartNum") = "PDCE" And Rod_Diam = "")
    Call Gleen_Info
  Wend
  
  Entry_Test
End Sub

Private Sub Submit_Click()
    On Error GoTo GeneralError
    Dim InspName As String: InspName = "SC_Inspection"
    Dim InspPlan As String: InspPlan = CalcSheet.Range(InspName & "_Plan")
    Dim InspSpec As String: InspSpec = CalcSheet.Range(InspName & "_Spec")
    Dim InspComm As String: InspComm = CalcSheet.Range(InspName & "_Comment")

    ClearCalcSheet True
    Validator

    If Operation = "Setup" And (Rod_Length = "" Or Wire_Diameter = "" Or Machine_No = "" Or Not (Straight_Fail Or Straight_Pass)) Then GoTo BlankForm
    If Operation = "Run" And (Not (Rod_Length_Fail Or Rod_Length_Pass) Or Not (Straight_Fail Or Straight_Pass) Or Machine_No = "") Then GoTo BlankForm

    CalcSheet.Range("Insp_Plan") = InspPlan
    CalcSheet.Range("Spec_ID") = InspSpec
    CalcSheet.Range("Check3") = Radio(Straight_Pass, Straight_Fail)
    CalcSheet.Range("Schar3") = Machine_No
    If Operation = "Setup" Then
        CalcSheet.Range("Data1") = TryParseFraction(Rod_Length)
        CalcSheet.Range("Data2") = TryParseFraction(Wire_Diameter)
    Else
        CalcSheet.Range("Check2") = Radio(Rod_Length_Pass, Rod_Length_Fail)
    End If
    
    If InspComm = Empty Then
        CalcSheet.Range("Passed") = 1
        CalcSheet.Range("Value") = ""
        CalcSheet.Range("Failed_Comment") = ""
    Else
        CalcSheet.Range("Passed") = 0
        CalcSheet.Range("Value") = "Rod Rejected"
        CalcSheet.Range("Failed_Comment") = Replace(InspComm, "?", ".  ")
        RejectForm Replace(InspComm, "?", vbNewLine)
    End If
    
    WriteToSQL "ice.UD10"

    Rod_Length = ""
    Rod_Length_Fail = False
    Rod_Length_Pass = False
    Straight_Fail = False
    Straight_Pass = False
    Wire_Diameter = ""
    
  If Operation = "Setup" Then
    Rod_Length.SetFocus
  Else
    Rod_Length_Pass.SetFocus
  End If

  Inspection_Num = "Inspection Num: " + CStr(SampleNum)
Exit Sub

GeneralError:
  MsgBox MBDataErrorsResbmit
Exit Sub

BlankForm:
  MsgBox MBFillOutResubmit
End Sub

Private Sub UserForm_Activate()
    
    ClearCalcSheet

    Set JobOper = New ADODB.Recordset

    DBEpicor.Open
    Set JobOper.ActiveConnection = DBEpicor
    JobOper.Open "SELECT * FROM " & "erp.JobOper" & " WHERE Company = " & Company & " AND JobNum = '" & JobNum & "' AND (OpCode = 'GBDSTR01' OR OpCode = 'FWDSTR01')"
    If JobOper.EOF Or JobOper.BOF Then
      MsgBox MBErrorOpComments
    End If
    CalcSheet.Range("Operation_Comment") = Compact(JobOper.Fields("CommentText"))
    DBEpicor.Close
    SetNextSampleNum
    Inspection_Num = "Inspection Num: " + CStr(SampleNum)
    Machine_No = ""
    
    If Operation = "Setup" Then
      .Rod_Length_Fail.Visible = False
      .Rod_Length_Pass.Visible = False
      .Rod_Length.Visible = True
      .Wire_Diameter.Visible = True
      .Lbl_Wire_Diam.Visible = True
    Else
      .Rod_Length.Visible = False
      .Wire_Diameter.Visible = False
      .Lbl_Wire_Diam.Visible = False
      .Rod_Length_Fail.Visible = True
      .Rod_Length_Pass.Visible = True
    End If
End Sub
