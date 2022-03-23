VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Weaving_Crimp_Inspection 
   Caption         =   "ABI Inspection Program"
   ClientHeight    =   11235
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   20415
   OleObjectBlob   =   "Weaving_Crimp_Inspection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Weaving_Crimp_Inspection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1

Private Sub UserForm_Activate()

  ClearCalcSheet Written:=False
  
  Dim JobOper As New ADODB.Recordset
  DBEpicor.Open
  Set JobOper.ActiveConnection = DBEpicor
  JobOper.Open "SELECT * FROM " & "erp.JobOper" & " WHERE Company = " & Company & " AND JobNum = '" & JobNum & "' AND OpCode = 'WBDCRI01'"
  If JobOper.EOF Or JobOper.BOF Then
    MsgBox MBErrorOpComments
  End If
  
  CalcSheet.Range("Operation_Comment") = Compact(JobOper.Fields("CommentText"))
  
  DBEpicor.Close
  
  SetNextSampleNum
  Inspection_Num = "Inspection Num: " & CStr(SampleNum)
  
  Machine_No = ""
  Validator
  
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = 0 Then
    Cancel = True
    MsgBox "The X is disabled, please use a button on the form.", vbCritical
  End If
End Sub

Private Sub Get_Results_Click()
  DisplayResults "Crimp Inspection", _
    "#,Date t,Type,Time      ,Employ,Spec       ,Part #,Wire Diam,Pitch,Crimp Depth,AOI,Nom Width,Flatness", _
    "Key3,Date01,Key2,ShortChar06,ShortChar02,Character07,Character02,Number04,ShortChar01,Number02,Checkbox02,Number03,Checkbox03"
End Sub

Private Sub Home_Click()
  Unload Crimp_Inspection
  Start_Screen.Show
End Sub

Private Sub Image1_Click()
  Validator
  Crimp_Setup.Show
End Sub

Private Sub Target_CrimpDepth_AfterUpdate()
  If Target_CrimpDepth >= 0 Then
    CalcSheet.Range("CrimpDepth") = Target_CrimpDepth
    CD = Target_CrimpDepth
    Validator
  Else
    CD = ""
  End If
End Sub

Private Sub Validator()

  While IsBad(BeltWidth) Or IsBad(CrimpDepth)
  
    Dim JobCom As String: JobCom = CalcSheet.Range("JobComments")
    Dim OpCom As String: OpCom = CalcSheet.Range("Operation_Comment")
  
    If BeltWidth = "" Then
      BeltWidth = RegExExecute(JobCom, RegEx.BeltWidth, 0, 1)
      If IsNumeric(BeltWidth) And BeltWidth > 0 Then
        CalcSheet.Range("BeltWidth") = BeltWidth
      Else
        BeltWidth = TryParseFraction(ShowCommentBox("Belt Width", JobCom))
        CalcSheet.Range("BeltWidth") = BeltWidth
      End If
    End If
    
    If CrimpDepth = "" Then
      CrimpDepth = Comment_Search("Crimp Depth", OpCom, "", "", "", "")
      If IsNumeric(CrimpDepth) Then
        CalcSheet.Range("CrimpDepth") = CrimpDepth
        Target_CrimpDepth = CrimpDepth
      Else
        CrimpDepth = TryParseFraction(ShowCommentBox("Crimp Depth", OpCom))
        If IsNumeric(CrimpDepth) Then
          CalcSheet.Range("CrimpDepth") = CrimpDepth
          Target_CrimpDepth = CrimpDepth
        Else
          CrimpDepth = ""
        End If
      End If
    End If
    
  Wend
  
End Sub

Private Sub Submit_Click()
  On Error GoTo GeneralError
  
  Const InspName As String = "Crimp_Inspection"
  
  ClearCalcSheet Written:=True
  Validator

  If Len(Machine_No) = 0 Or Wire_Diameter = 0 Or Pitch = 0 Or CrimpDepth = 0 Or Nominal_Width = 0 Or _
    Not (AO_Pass Or AO_Fail) Or Not (Flatness_Pass Or Flatness_Fail) Then GoTo BlankForm

  With CalcSheet
    .Range("CrimpDepth") = TryParseFraction(Target_CrimpDepth)
    .Range("Insp_Plan") = .Range(InspectName & "_Plan")
    .Range("Spec_ID") = .Range(InspectName & "_Spec")
    .Range("Schar3") = Machine_No
    .Range("Data4") = TryParseFraction(Wire_Diameter)
    .Range("Schar1") = Pitch
    .Range("Data2") = TryParseFraction(CrimpDepth)
    .Range("Check2") = Radio(AO_Pass, AO_Fail)
    .Range("Data3") = TryParseFraction(Nominal_Width)
    .Range("Check3") = Radio(Flatness_Pass, Flatness_Fail)
    
    'Calculate Results
    If .Range(InspName & "_Comment") = Empty Then
      .Range("Passed") = 1
      .Range("Value") = ""
      .Range("Failed_Comment") = ""
    Else
      .Range("Passed") = 0
      .Range("Value") = "Rod Rejected"
      .Range("Failed_Comment") = Replace(.Range(InspName & "_Comment"), "?", ".  ")
      Rejection_Form Replace(.Range(InspName & "_Comment"), "?", vbNewLine)
    End If
  End With
  
  'Send Data To SQL
  WriteToSQL "ice.UD10"
  
  SetNextSampleNum
  ClearCalcSheet Written:=True
  
  Pitch = ""
  CrimpDepth = ""
  Nominal_Width = ""
  AO_Pass = False
  AO_Fail = False
  Flatness_Pass = False
  Flatness_Fail = False
  Inspection_Num = "Inspection Num: " + CStr(CalcSheet.Range("SampleNum"))
  Pitch.SetFocus
Exit Sub

GeneralError:
  MsgBox MBDataErrorsResbmit
Exit Sub

BlankForm:
  MsgBox MBFillOutResubmit
End Sub
