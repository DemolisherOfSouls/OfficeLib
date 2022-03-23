VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Start_Screen 
   Caption         =   "ABI Inspection Program"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   17070
   OleObjectBlob   =   "Start_Screen.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Start_Screen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1

Private Sub UserForm_Activate()

  With CompanySelection
    .AddItem 210
    .AddItem 236
    .AddItem 237
    .AddItem 300
  End With

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = 0 Then
    Cancel = True
    MsgBox MBExitDisabled, vbCritical
  End If
End Sub

Private Sub Clear_Click()
  Job_Num = ""
  Inspection_Type.Clear
  CompanySelection.Clear
  Employee_Num = ""
  Run = False
  Setup = False
  
  Operation = ""
  JobNum = ""
  Inspection = ""
  Company = 200
End Sub

Private Sub Close_Form_Click()
  Unload Start_Screen
  
  If Employee_Num <> 14114 And Not IsText(Employee_Num) Then
    With Application
      .DisplayAlerts = False
      .ActiveWindow.Close False
    End With
  End If
End Sub

Private Sub AddOp(ByVal text As String)
  Dim cb As ComboBox: Set cb = Inspection_Type
  If (cb.ListCount > 0) Then
    If (cb.List(cb.ListCount - 1) = text) Then
      Exit Sub
    End If
  End If
  cb.AddItem text
End Sub

Private Sub Job_Num_Exit(ByVal Cancel As MSForms.ReturnBoolean)
  
  If Job_Num = "" Then
    Exit Sub
  End If
  
  If (DBEpicor.ConnectionString = "") Then
    ProgramInit
  End If

  JobNum = Job_Num
  
  'Get Job from DB
  Dim JobHead As New ADODB.Recordset
  Dim JobOper As New ADODB.Recordset
  Dim JobComment As String
  Dim PartNum As String

  DBEpicor.Open
  
  Set JobHead.ActiveConnection = DBEpicor
  Set JobOper.ActiveConnection = DBEpicor
  
  'Get Comments & Ops
  JobHead.Open "SELECT * FROM " & "erp.JobHead" & " WHERE Company = 200 AND JobNum = '" & JobNum & "'"
  JobOper.Open "SELECT * FROM " & "erp.JobOper" & " WHERE Company = 200 AND JobNum = '" & JobNum & "'"
  
  If JobHead.EOF Or JobHead.BOF Then
    MsgBox "Invalid Job Number Entered"
    Exit Sub
  End If

  PartNum = UCase(JobHead.Fields("PartNum"))
  JobComment = Compact(JobHead.Fields("CommentText"))
  
  CalcSheet.Range("JobComments") = JobComment
  CalcSheet.Range("JobNum") = JobNum
  CalcSheet.Range("PartNum") = PartNum
  SetBeltType PartNum

  If JobOper.EOF Or JobOper.BOF Then
    MsgBox "Error Returning the Operation Comments"
    Exit Sub
  End If
  
  JobOper.MoveFirst
  Inspection_Type.Clear
  
  While Not JobOper.EOF
    Select Case JobOper.Fields("OPCode").Value
      Case "ELDAMWo1" 'Eyelink Assembly & Weld Opt
      Case "ELDELF01" 'Eyelink Forming Operation
      Case "ELDPWD01" 'Eyelink Panel Weld Operation
      Case "FWDBUT01" 'Flatwire Buttoner Operation
      Case "FWDCLI01" 'Flatwire Clincher Operation
        AddOp "Flatwire Clincher Inspection"
      Case "FWDHAM01" 'Flatwire Hand Assembly
      Case "FWDLPS01" 'Flatwire Link Press Opt
      Case "FWMUL01"  'Flatwire MultiSlide Operation
        AddOp "Flatwire Picket Inspection"
      Case "FWDPDL01" 'Flatwire Picket Dimpling Oper
      Case "FWDRCO01" 'Flatwire Reverse Cutoff Oper
      Case "FWDSTR01" 'Flatwire S/C Operation
        AddOp "Straight and Cut Inspection"
      Case "OGBWO01"  'Grid Belt Washing
      Case "GBBUT01"  'Grid Buttoning Operation
        AddOp "Grid Buttoning Inspection"
      Case "GBDHAD01" 'Grid Hand Assembly Operation
      Case "GBDPTL01" 'Grid Pig Tail Forming Opt
      Case "GBDSPR01" 'Grid Spring Forming Operation
        AddOp "Grid Spiral Inspection"
      Case "GBDSTR01" 'Grid Straight and Cut Opt
        AddOp "Straight and Cut Inspection"
      Case "GBDWEL01" 'Grid Welding Operation
        AddOp "Grid Welding Inspection"
      Case "MLDCRA01" 'Matl Crating Operation
        'AddOp "Shipping Sign Off"
      Case "PSDBMH01" 'Plastic Break Modules Opt
      Case "PSDHAS01" 'Plastic Hand Assembly Opt
      Case "PSDASP01" 'Plastic Specialty Operation
      Case "WBDTAC01" 'Tack Weld Flat Wire Belt
      Case "WBDASB01" 'Woven Assembly Operation
      Case "WBDCRI01" 'Woven Crimp Forming Operation
        AddOp "Crimp Inspection"
      Case "WBDSPE01" 'Woven Specialty Operation
      Case "WBDSRF01" 'Woven Spiral Form Operation
        If UCase(CalcSheet.Range("LPartNum")) = "CB5BAND" Then
          AddOp "CB5 Weaving Inspection"
        Else
          AddOp "Weaving Spiral Inspection"
        End If
      Case "WBDWEL01" 'Woven Weld & Trim Operation
      Case Else
    End Select
    
    JobOper.MoveNext
  Wend
  
  DBEpicor.Close
  
  CalcSheet.Range("Operation_Comment") = ""

End Sub

Private Sub Open_Data_Entry_Click()
  
  If Employee_Num = "" Or Not (Setup Or Run) Then GoTo BlankForm

  SetOperation Inspection_Type, Run, Setup
  SetNextSampleNum
  
  CalcSheet.Range("Employee_Num") = Employee_Num
  DiffSpiralCount = CalcSheet.Range("DiffSpiralCount")

  Select Case Inspection
    Case "Grid Spiral Inspection"
      Hide
      Grid_Spiral_Inspection.Show
    Case "Straight and Cut Inspection"
      Hide
      SC_Inspection.Show
    Case "Grid Welding Inspection"
      Hide
      Grid_Welding_Inspection.Show
    Case "Weaving Spiral Inspection"
      Hide
      Weaving_Spiral_Inspection.Show
    Case "Crimp Inspection"
      Hide
      Weaving_Crimp_Inspection.Show
    Case "CB5 Weaving Inspection"
      Hide
      Weaving_CB5_Inspection.Show
    Case "Flatwire Picket Inspection"
      Hide
      Flatwire_Picket_Inspection.Show
  End Select
Exit Sub
    
BlankForm:
  MsgBox MBFillOutResubmit
End Sub
