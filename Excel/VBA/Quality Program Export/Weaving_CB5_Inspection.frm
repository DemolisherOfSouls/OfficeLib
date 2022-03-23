VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Weaving_CB5_Inspection 
   Caption         =   "ABI Inspection Program"
   ClientHeight    =   11145
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   20415
   OleObjectBlob   =   "Weaving_CB5_Inspection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Weaving_CB5_Inspection"
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

Private Sub Get_Results_Click()
  DisplayResults "CB5 Weaving Inspection", _
    "#,Date    ,Type,Time      ,Employ,Spec       ,Part #,Spiral Hand,Flat Wire Height,Spiral Outside Height,Spiral Outside Width,Fabric Width,C to C,Roll", _
    "Key3,Date01,Key2,ShortChar06,ShortChar02,Character07,Character02,ShortChar04,Number01,Number02,Number03,Checkbox02,Number04,Checkbox03"
End Sub

Private Sub Home_Click()
  Unload CB5_Weave_Inspection
  Start_Screen.Show
End Sub

Private Sub Image1_Click()
  CB5_Weave_Setup.Show
End Sub

Private Sub Submit_Click()
  On Error GoTo GeneralError
  
  Const InspName As String = "CB5_Weave_Inspection"
  Dim InspPlan As String: InspPlan = Range(InspName & "_Plan")
  Dim InspSpec As String: InspSpec = Range(InspName & "_Spec")
  Dim InspComm As String: InspComm = Range(InspPlan & "_Comment")
    
  ClearCalcSheet Written:=True

  With CalcSheet
    .Range("Insp_Plan") = InspPlan
    .Range("Spec_ID") = InspSpec
    .Range("Schar4") = Radio(RH_Spiral, LH_Spiral, "RH Spiral", "LH Spiral")
    .Range("Schar3") = Machine_No
    .Range("Data1") = TryParseFraction(Flat_Wire_Height)
    .Range("Data2") = TryParseFraction(Spiral_Outside_Height)
    .Range("Data3") = TryParseFraction(Spiral_Outside_Width)
    .Range("Data4") = TryParseFraction(Center_To_Center)
    .Range("Check2") = Radio(Fabric_Width_Pass, Fabric_Width_Fail, 1, 0)
    .Range("Check3") = Radio(Roll_Pass, Roll_Fail, 1, 0)
  
    'Check That Form is Complete
    If IsEmpty(.Range("Schar4")) Or _
      IsEmpty(.Range("Schar3")) Or _
      IsEmpty(.Range("Data1")) Or _
      IsEmpty(.Range("Data2")) Or _
      IsEmpty(.Range("Data3")) Or _
      IsEmpty(.Range("Data4")) Or _
      IsEmpty(.Range("Check2")) Or _
      IsEmpty(.Range("Check3")) Then
      GoTo BlankForm
    End If
  
    'Calculate Results
    If InspComm = Empty Then
      .Range("Passed") = 1
      .Range("Value") = ""
      .Range("Failed_Comment") = ""
    Else
      Temp = Replace(InspComm, "?", ".  ")
      .Range("Passed") = 0
      .Range("Value") = "Rod Rejected"
      .Range("Failed_Comment") = Temp
      RejectForm Replace(InspComm, "?", vbNewLine)
    End If
  End With

  WriteToSQL "ice.UD10"
  
  Flat_Wire_Height = ""
  Spiral_Outside_Height = ""
  Spiral_Outside_Width = ""
  Center_To_Center = ""
  Fabric_Width_Pass = False
  Fabric_Width_Fail = False
  Roll_Pass = False
  Roll_Fail = False
  Flat_Wire_Height.SetFocus
  SetNextSampleNum
  Inspection_Num = "Inspection Num: " & CStr(SampleNum)
Exit Sub

GeneralError:
  MsgBox MBDataErrorsResbmit
Exit Sub

BlankForm:
  MsgBox "Please Completely Fill out the Form and Resubmit."
End Sub

Private Sub UserForm_Activate()

  ClearCalcSheet Written:=False
     
  Dim JobOper As New ADODB.Recordset

  DBEpicor.Open
  Set JobOper.ActiveConnection = DBEpicor
  JobOper.Open "SELECT * FROM " & "erp.JobOper" & " WHERE Company = " & Company & " AND JobNum = '" & JobNum & "' AND OpCode = 'WBDSRF01'"

  If JobOper.EOF Or JobOper.BOF Then
    MsgBox "Error Returning the Operation Comments"
  Else
    CalcSheet.Range("Operation_Comment") = Compact(JobOper.Fields("CommentText"))
  End If
  
  DBEpicor.Close
  
  SetNextSampleNum
  
  Inspection_Num = "Inspection Num: " & CStr(InspNum)
  Machine_No = ""
End Sub
