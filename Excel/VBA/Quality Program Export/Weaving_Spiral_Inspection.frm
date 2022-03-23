VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Weaving_Spiral_Inspection 
   Caption         =   "ABI Inspection Program"
   ClientHeight    =   10980
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   20415
   OleObjectBlob   =   "Weaving_Spiral_Inspection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Weaving_Spiral_Inspection"
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
  DisplayResults "Weaving Spiral Inspection", _
    "#,Date    ,Type,Time      ,Employ,Spec       ,Part #,R/S Spiral,Spiral Hand,Thickness,Width,Loop Count(Actual),Loop Count(Visual),Linear Pitch,AOI", _
    "Key3,Date01,Key2,ShortChar06,ShortChar02,Character07,Character02,ShortChar07,ShortChar04,Number01,Number02,Number03,CheckBox03,ShortChar05,Checkbox02"
End Sub

Private Sub Home_Click()
  Unload Weaving_Spiral_Inspection
  Start_Screen.Show
End Sub

Private Sub Image1_Click()
  Run_Setup
  Weaving_Spiral_Setup.Show
End Sub

Private Sub Reg_Spiral_Click()
  SpiralSize = ""
  Run_Setup
End Sub

Private Sub Spec_Spiral_Click()
  SpiralSize = ""
  Run_Setup
End Sub

Private Sub Validator()

  Const OpComment As String = CalcSheet.Range("Operation_Comment")
  
  Dim novalue, toohigh
  
Start:

  While SpiralSize = "" Or Loops = 0
    If IsEmpty(SpiralSize) Then
      SpiralSize = RegExExecute(OpComment, RegEx.SpiralSize)
      If IsEmpty(SpiralSize) Or IsError(SpiralSize) Then
        SpiralSize = Replace(ShowCommentBox("Spiral Size (Example .250x.125)", OpComment), " ", "")
        
        If IsError(SpiralSize) Then SpiralSize = ""
      End If
      
      CalcSheet.Range("Spiral_Size") = SpiralSize
    End If
    
    If IsBad(Loops) Then
      Loops = RegExExecute(OpComment, "Loops")
      If Not IsBad(Loops) Then
        CalcSheet.Range("Loop_Count") = Loops
      Else
        Loops = ""
      End If
    End If
  Wend
  
  novalue = Not IsNumeric(CalcSheet.Range("Spiral_Thick")) Or Not IsNumeric(CalcSheet.Range("Spiral_Width"))
  toohigh = CalcSheet.Range("Spiral_Thick") > 1.75 Or CalcSheet.Range("Spiral_Width") > 2.5
  
  If toohigh Then MsgBox "The Spiral Size entered is too large. Please Check your numbers.", vbOKOnly, "Error"
  If novalue Or toohigh Then SpiralSize = ""
    
  If Loops = 0 Or novalue Or toohigh Then GoTo Start
  
End Sub

Private Sub Submit_Click()
  On Error GoTo GeneralError
  
  Const InspName As String = "Weaving_Inspection"
  Dim InspPlan As String: InspPlan = Range(InspName & "_Plan")
  Dim InspSpec As String: InspSpec = Range(InspName & "_Spec")
  Dim InspComm As String: InspComm = Range(InspPlan & "_Comment")
  
  ClearCalcSheet Written:=True
  Validator
    
    With CalcSheet
    .Range("Insp_Plan") = InspPlan
    .Range("Spec_ID") = InspSpec
    .Range("Schar7") = Radio(Reg_Spiral, Spec_Spiral, "Regular Spiral", "Special Spiral")
    .Range("Schar4") = Radio(RH_Spiral, LH_Spiral, "RH Spiral", "LH Spiral")
    .Range("Schar3") = Machine_No
    .Range("Data1") = TryParseFraction(Thickness)
    .Range("Data2") = TryParseFraction(Width_Setup)
    .Range("Check2") = Radio(AO_Pass, AO_Fail, 1, 0)
    If Operation = "Run" Then
        .Range("Check3") = Radio(Loop_Count_Pass, Loop_Count_Fail, 1, 0)
    Else
        .Range("Data3") = TryParseFraction(Loop_Count)
        .Range("Schar8") = Linear_Pitch
    End If
    
    'Check That Form is Complete
    If IsEmpty(.Range("Schar7")) Or _
      IsEmpty(.Range("Schar4")) Or _
      IsEmpty(.Range("Schar3")) Or _
      IsEmpty(.Range("Data1")) Or _
      IsEmpty(.Range("Data2")) Or _
      IsEmpty(.Range("Check2")) Or _
      (Inspection_Type = "Run" And IsEmpty(.Range("Check3"))) Or _
      (Inspection_Type = "Setup" And (IsEmpty(.Range("Data3")) Or IsEmpty(.Range("Schar8")))) Then GoTo BlankForm

    'Calculate Results
    If .Range(Inspection_Name & "_Comment") = Empty Then
        .Range("Passed") = 1
        .Range("Value") = ""
        .Range("Failed_Comment") = ""
    Else
        .Range("Passed") = 0
        .Range("Value") = "Rod Rejected"
        .Range("Failed_Comment") = Replace(.Range(InspName & "_Comment"), "?", ".  ")
        RejectForm Replace(.Range(InspName & "_Comment"), "?", vbNewLine)
    End If
    End With

    WriteToSQL "ice.UD10"
    
    'Clear Form and Update Sample Number
    Thickness = ""
    Width_Setup = ""
    Loop_Count = ""
    Loop_Count_Pass = False
    Loop_Count_Fail = False
    Linear_Pitch = ""
    AO_Pass = False
    AO_Fail = False
    Thickness.SetFocus
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
  
  Set JobOper = New ADODB.Recordset
  
  DBEpicor.Open
  
  Set JobOper.ActiveConnection = DBEpicor
  JobOper.Open "SELECT * FROM " & "erp.JobOper" & " WHERE Company = " & Company & " AND JobNum = '" & JobNum & "' AND OpCode = 'WBDSRF01'"
  
  If JobOper.EOF Or JobOper.BOF Then
    MsgBox "Error Returning the Operation Comments"
  End If
  
  CalcSheet.Range("Operation_Comment") = Comment_Format(JobOper.Fields("CommentText"))
  
  DBEpicor.Close
  
  SetNextSampleNum
  Machine_No = ""
  CalcSheet.Range("Spiral_Size") = ""
  Inspection_Num = "Inspection Num: " & CStr(SampleNum)
  
    If Operation = "Setup" Then
      Loop_Count_Pass.Visible = False
      Loop_Count_Fail.Visible = False
      Loop_Count.Visible = True
      Linear_Pitch.Visible = True
    Else
      Loop_Count_Pass.Visible = True
      Loop_Count_Fail.Visible = True
      Loop_Count.Visible = False
      Linear_Pitch.Visible = False
    End If
End Sub
