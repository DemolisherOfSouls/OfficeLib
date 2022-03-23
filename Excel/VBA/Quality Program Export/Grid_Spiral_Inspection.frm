VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Grid_Spiral_Inspection 
   Caption         =   "ABI Inspection Program"
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   20415
   OleObjectBlob   =   "Grid_Spiral_Inspection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Grid_Spiral_Inspection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1

Private Sub Close_Form_Click()
  Unload Spiral_Form
  Unload Start_Screen
End Sub

Private Sub Data_Dump_Click()

  Dim UD10 As New ADODB.Recordset
  Dim strSQL As String
  Dim i As Integer: i = 2

  GraphSheet.Cells.Clear
  
  Dim Step: Step = 0
  Dim Row: Row = 7
  Dim Col: Col = 65
  Dim A: A = ""
  
  Dim X: For X = 1 To 60
  
    If Col = 91 And A = "" Then
      A = "A"
      Col = 65
    ElseIf Col = 91 And A = "A" Then
      A = "B"
      Col = 65
    End If
    
    Select Case Step
    Case 0
      GraphSheet.Range(A & Chr(Col) & "1") = "Job Number"
      Step = 1: Col = Col + 1
    Case 1
      GraphSheet.Range(A & Chr(Col) & "1") = CalcSheet.Range("J" & Row)
      Step = 2: Col = Col + 1: Row = Row + 1
    Case 2
      GraphSheet.Range(A & Chr(Col) & "1") = "Min"
      Step = 3: Col = Col + 1
    Case 3
      GraphSheet.Range(A & Chr(Col) & "1") = "Target"
      Step = 4: Col = Col + 1
    Case 4
      GraphSheet.Range(A & Chr(Col) & "1") = "Max"
      Step = 1: Col = Col + 1
    End Select
        
  Next X

  DBEpicor.Open
  Set UD10.ActiveConnection = DBEpicor
  UD10.Open "SELECT * FROM " & "ice.UD10" & " WHERE Key1 = '" & JobNum & "' AND Key2 = '" & Insp_Type & " " & Operation & "' And Checkbox20 = '0'"
  
  If UD10.EOF Or UD10.BOF Then
    MsgBox MBNoData
    Exit Sub
  End If
  UD10.MoveFirst
  
  While UD10.EOF = False
    
    With GraphSheet
    
      .Range("A" & i) = UD10.Fields("Key1")
      .Range("B" & i) = UD10.Fields("Number01").Value
      .Range("C" & i) = CalcSheet.Range("N" & 7) + CalcSheet.Range("L" & 7)
      .Range("D" & i) = CalcSheet.Range("L" & 7)
      .Range("E" & i) = CalcSheet.Range("Q" & 7) + CalcSheet.Range("L" & 7)
      .Range("F" & i) = UD10.Fields("Number02")
      .Range("G" & i) = CalcSheet.Range("N" & 8) + CalcSheet.Range("L" & 8)
      .Range("H" & i) = CalcSheet.Range("L" & 8)
      .Range("I" & i) = CalcSheet.Range("Q" & 8) + CalcSheet.Range("L" & 8)
      .Range("J" & i) = UD10.Fields("Number03")
      .Range("K" & i) = CalcSheet.Range("N" & 9) + CalcSheet.Range("L" & 9)
      .Range("L" & i) = CalcSheet.Range("L" & 9)
      .Range("M" & i) = CalcSheet.Range("Q" & 9) + CalcSheet.Range("L" & 9)
      .Range("N" & i) = UD10.Fields("Number04")
      .Range("O" & i) = CalcSheet.Range("N" & 10) + CalcSheet.Range("L" & 10)
      .Range("P" & i) = CalcSheet.Range("L" & 10)
      .Range("Q" & i) = CalcSheet.Range("Q" & 10) + CalcSheet.Range("L" & 10)
      .Range("R" & i) = UD10.Fields("Number05")
      .Range("S" & i) = CalcSheet.Range("N" & 11) + CalcSheet.Range("L" & 11)
      .Range("T" & i) = CalcSheet.Range("L" & 11)
      .Range("U" & i) = CalcSheet.Range("Q" & 11) + CalcSheet.Range("L" & 11)
      .Range("V" & i) = UD10.Fields("Number06")
      .Range("W" & i) = CalcSheet.Range("N" & 12) + CalcSheet.Range("L" & 12)
      .Range("X" & i) = CalcSheet.Range("L" & 12)
      .Range("Y" & i) = CalcSheet.Range("Q" & 12) + CalcSheet.Range("L" & 12)
      .Range("Z" & i) = UD10.Fields("Number07")
      .Range("AA" & i) = CalcSheet.Range("N" & 13) + CalcSheet.Range("L" & 13)
      .Range("AB" & i) = CalcSheet.Range("L" & 13)
      .Range("AC" & i) = CalcSheet.Range("Q" & 13) + CalcSheet.Range("L" & 13)
      .Range("AD" & i) = UD10.Fields("Number08")
      .Range("AE" & i) = CalcSheet.Range("N" & 14) + CalcSheet.Range("L" & 14)
      .Range("AF" & i) = CalcSheet.Range("L" & 14)
      .Range("AG" & i) = CalcSheet.Range("Q" & 14) + CalcSheet.Range("L" & 14)
      .Range("AH" & i) = UD10.Fields("Number09")
      .Range("AI" & i) = CalcSheet.Range("N" & 15) + CalcSheet.Range("L" & 15)
      .Range("AJ" & i) = CalcSheet.Range("L" & 15)
      .Range("AK" & i) = CalcSheet.Range("Q" & 15) + CalcSheet.Range("L" & 15)
      .Range("AL" & i) = UD10.Fields("Number10")
      .Range("AM" & i) = CalcSheet.Range("N" & 16) + CalcSheet.Range("L" & 16)
      .Range("AN" & i) = CalcSheet.Range("L" & 16)
      .Range("AO" & i) = CalcSheet.Range("Q" & 16) + CalcSheet.Range("L" & 16)
      .Range("AP" & i) = UD10.Fields("Number11")
      .Range("AQ" & i) = CalcSheet.Range("N" & 17) + CalcSheet.Range("L" & 17)
      .Range("AR" & i) = CalcSheet.Range("L" & 17)
      .Range("AS" & i) = CalcSheet.Range("Q" & 17) + CalcSheet.Range("L" & 17)
      .Range("AT" & i) = UD10.Fields("Number12")
      .Range("AU" & i) = CalcSheet.Range("N" & 18) + CalcSheet.Range("L" & 18)
      .Range("AV" & i) = CalcSheet.Range("L" & 18)
      .Range("AW" & i) = CalcSheet.Range("Q" & 18) + CalcSheet.Range("L" & 18)
      .Range("AX" & i) = UD10.Fields("Number13")
      .Range("AY" & i) = CalcSheet.Range("N" & 19) + CalcSheet.Range("L" & 19)
      .Range("AZ" & i) = CalcSheet.Range("L" & 19)
      .Range("BA" & i) = CalcSheet.Range("Q" & 19) + CalcSheet.Range("L" & 19)
      .Range("BB" & i) = UD10.Fields("Number14")
      .Range("BC" & i) = CalcSheet.Range("N" & 20) + CalcSheet.Range("L" & 20)
      .Range("BD" & i) = CalcSheet.Range("L" & 20)
      .Range("BE" & i) = CalcSheet.Range("Q" & 20) + CalcSheet.Range("L" & 20)
      .Range("BF" & i) = UD10.Fields("Number15")
      .Range("BG" & i) = CalcSheet.Range("N" & 21) + CalcSheet.Range("L" & 21)
      .Range("BH" & i) = CalcSheet.Range("L" & 21)
      .Range("BI" & i) = CalcSheet.Range("Q" & 21) + CalcSheet.Range("L" & 21)
  
    End With
  
    UD10.MoveNext
    i = i + 1
  Wend
    
  DBEpicor.Close
  MsgBox "Complete"
  GraphSheet.Visible = ccc
End Sub

Private Sub Get_Results_Click()
  DisplayResults "Grid Spiral Inspection", _
    "#,Date    ,Type,Time      ,Employ,Spec       ,Part #,Spiral Hand,I/O Spiral   ,Height(B),C+ADJ,Long Leg Len,Tri Leg Len,Height(D),Width(E),Height(B),F+ADJ,Long Leg Len,Tri Leg Len,Height(D),Width(E),Diam(G),Leg Len,Fabric Width,Ref,Dog Leg,Burrs,Spiral Twist", _
    "Key3,Date01,Key2,ShortChar06,ShortChar02,Character07,Character02,ShortChar04,ShortChar05,Number01,Number02,Number03,Number04,Number05,Number06,Number07,Number08,Number09,Number10,Number11,Number12,Number13,Number14,Number15,ShortChar01,CheckBox02,CheckBox03,CheckBox04"
End Sub

Private Sub Home_Click()
  Unload Spiral_Form
  Start_Screen.Show
  GraphSheet.Visible = xlSheetHidden
End Sub

Private Sub Image1_Click()
  Validator
  Grid_Spiral_Setup.Show
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = 0 Then
    Cancel = True
    MsgBox "The X is disabled, please use a button on the form.", vbCritical
  End If
End Sub

Private Sub SpiralIOCheck(ByVal Check As Boolean)
  If Check And CalcSheet.Range("DiffSpiralCount") = 2 Then
    CalcSheet.Range("IO_Spiral") = Radio(Inside_Spiral, Outside_Spiral, "Inside Spiral", "Outside Spiral")
    Fabric_Width = ""
    Validator
  End If
End Sub

Private Sub SpiralLRCheck(ByVal Inside As Boolean)
  If BeltType = "ASB" Or BeltType = "ASB-W" Then
    Inside_Spiral = Inside
    Outside_Spiral = Not Inside
    CalcSheet.Range("IO_Spiral") = Radio(Inside_Spiral, Outside_Spiral, "Inside Spiral", "Outside Spiral")
    Validator
  End If
End Sub

Private Sub Inside_Spiral_Click()
  SpiralIOCheck Inside_Spiral
End Sub

Private Sub Outside_Spiral_Click()
  SpiralIOCheck Outside_Spiral
End Sub

Private Sub RH_Spiral_Click()
  SpiralLRCheck True
End Sub

Private Sub LH_Spiral_Click()
  SpiralLRCheck False
End Sub

Private Sub Validator()

  Dim Job_Comments As String: Job_Comments = CalcSheet.Range("JobComments")
  Dim Op_Comment As String: Op_Comment = CalcSheet.Range("Operation_Comment")

  While IsBad(GlobalVars.BeltWidth) Or (IsBad(Fabric_Width) And (Inside_Spiral Or Outside_Spiral)) Or (IsBad(Center_Link_Location) And DiffSpiralCount > 1)
    
    If BeltWidth = "" Then
        BeltWidth = Comment_Search("Width", Job_Comments, "Inches", "in.", "", "")
        
        'Added 2 More options -JT
        If Len(BeltWidth) = 0 Or IsNumeric(BeltWidth) = False Then
          BeltWidth = Comment_Search("Overall Belt Width:", Job_Comments, "inches", "in.", "in", "")
        End If
        If Len(BeltWidth) = 0 Or IsNumeric(BeltWidth) = False Then
          BeltWidth = Comment_Search("Belt width:", Job_Comments, "inches", "in.", "in", "")
        End If
        
        If IsNumeric(BeltWidth) And BeltWidth > 0 Then
          CalcSheet.Range("BeltWidth") = BeltWidth
        Else
          BeltWidth = TryParseFraction(ShowCommentBox("Belt Width", Job_Comments))
          CalcSheet.Range("BeltWidth") = BeltWidth
        End If
    End If
    
    If DiffSpiralCount > 1 And Center_Link_Location = "" Then
        Center_Link_Location = Comment_Search("Center Link", Job_Comments, "inches", "in.", "in", "")
        'Added 1 More option -JT
        If Len(Center_Link_Location) = 0 Or IsNumeric(Center_Link_Location) = False Then
            Center_Link_Location = Comment_Search("Center Link Location:", Job_Comments, "inches", "in.", "in", "")
        End If
        If IsNumeric(Center_Link_Location) And Center_Link_Location > 0 Then
            CalcSheet.Range("Center_Link_Location") = Center_Link_Location
        Else
            Center_Link_Location = TryParseFraction(ShowCommentBox("Center Link Location", Job_Comments))
            CalcSheet.Range("Center_Link_Location") = Center_Link_Location
        End If
    End If
    
    If (Fabric_Width = "" Or Fabric_Width = 0) And (Inside_Spiral = True Or Outside_Spiral = True) Then
      If DiffSpiralCount = 1 Then
        Fabric_Width = Comment_Search("Fabric Width", Job_Comments, "", "", "", "")
      End If
      If IsNumeric(Fabric_Width) And Fabric_Width > 0 Then
        CalcSheet.Range("Fabric_Width") = Fabric_Width
      Else
        Fabric_Width = TryParseFraction(ShowCommentBox("Fabric Width", Op_Comment))
        CalcSheet.Range("Fabric_Width") = Fabric_Width
      End If
    End If
  Wend
  
End Sub

Private Sub Submit_Click()
  On Error GoTo GeneralError
  
  Const InspName As String = "Spiral_Inspection"
  Dim InspPlan As String: InspPlan = .Range(InspName & "_Plan")
  Dim InspSpec As String: InspSpec = .Range(InspName & "_Spec")
  
  If (Machine_No <> Empty) Or Not (LH_Spiral Or RH_Spiral) Then
    GoTo BlankForm
  End If
    
  ClearCalcSheet Written:=True
  Validator

  With CalcSheet
    .Range("Insp_Plan") = InspPlan
    .Range("Spec_ID") = .Range(InspSpec)
    .Range("Data1") = TryParseFraction(F_B)
    .Range("Data2") = TryParseFraction(F_C_ADJ)
    .Range("Data3") = TryParseFraction(F_Long_Leg)
    .Range("Data4") = TryParseFraction(F_Tri_Leg_Len)
    .Range("Data5") = TryParseFraction(F_D)
    .Range("Data6") = TryParseFraction(F_E)
    .Range("Data7") = TryParseFraction(S_B)
    .Range("Data8") = TryParseFraction(S_C_ADJ)
    .Range("Data9") = TryParseFraction(S_Long_Leg)
    .Range("Data10") = TryParseFraction(S_Tri_Leg_Len)
    .Range("Data11") = TryParseFraction(S_D)
    .Range("Data12") = TryParseFraction(S_E)
    .Range("Data13") = TryParseFraction(P_G)
    .Range("Data14") = TryParseFraction(P_Leg_Len)
    .Range("Data15") = TryParseFraction(O_Fab_Wid)
    .Range("Schar1") = Trim(O_Ref)
    .Range("Schar3") = Trim(Machine_No)
    
    If BeltType = "ASB" Or BeltType = "ASB-W" Then
      .Range("Schar4") = Radio(LH_Spiral, RH_Spiral, "Spiral B", "Spiral A")
    Else
      .Range("Schar4") = Radio(LH_Spiral, RH_Spiral, "LH Spiral", "RH Spiral")
    End If
    
    .Range("Check2") = Radio(Dog_Leg)
    .Range("Check3") = Radio(Burrs)
    .Range("Check4") = Radio(Spiral_Twist)

    'Results
    If .Range(InspName & "_Comment") = Empty Then
      .Range("Passed") = 1
      .Range("Value") = ""
      .Range("Failed_Comment") = ""
    ElseIf IsError(.Range(InspName & "_Comment")) Then
      GoTo MissingData
    Else
      Temp = Replace(.Range(InspName & "_Comment"), "?", ".  ")
      .Range("Passed") = 0
      .Range("Failed_Comment") = Temp
      .Range("Value") = "Spiral Rejected"
      RejectForm Replace(.Range(InspName & "_Comment"), "?", vbNewLine)
    End If
    
  End With

  WriteToSQL "ice.UD10"

  .F_B.Value = ""
  .F_C_ADJ.Value = ""
  .F_Long_Leg.Value = ""
  .F_Tri_Leg_Len = ""
  .F_D.Value = ""
  .F_E.Value = ""
  .S_B.Value = ""
  .S_C_ADJ.Value = ""
  .S_Long_Leg.Value = ""
  .S_Tri_Leg_Len.Value = ""
  .S_D.Value = ""
  .S_E.Value = ""
  .P_G.Value = ""
  .P_Leg_Len.Value = ""
  .O_Fab_Wid.Value = ""
  .O_Ref.Value = ""
  .Dog_Leg.Value = False
  .Burrs.Value = False
  .Spiral_Twist = False
  Inspection_Num = "Inspection Num: " + CStr(SampleNum)
  .F_B.SetFocus
  
Exit Sub

GeneralError:
  MsgBox MBDataErrorsResbmit
Exit Sub

BlankForm:
  MsgBox MB
Exit Sub

MissingData:
  MsgBox MBDataMissingContact
End Sub

Private Sub UserForm_Activate()

  ClearCalcSheet Written:=False
          
  Dim JobOper As New ADODB.Recordset
  DBEpicor.Open
  Set JobOper.ActiveConnection = DBEpicor
  JobOper.Open "SELECT * FROM " & "erp.JobOper" & " WHERE Company = 200 AND JobNum = '" & JobNum & "' AND OpCode = 'GBDSPR01'"
  If JobOper.EOF Or JobOper.BOF Then
    MsgBox MBErrorOpComments
  End If
  CalcSheet.Range("Operation_Comment") = Compact(JobOper.Fields("CommentText"))
  DBEpicor.Close
  
  Inspection_Num = "Inspection Num: " + CStr(SampleNum)

  If DiffSpiralCount = 1 Then
    Inside_Spiral.Visible = False
    Outside_Spiral.Visible = False
    Inside_Spiral.Value = True
  Else
    Inside_Spiral.Visible = True
    Outside_Spiral.Visible = True
    Inside_Spiral = False
  End If
  
  'Setup For Stacker Belts
  If BeltType = "ASB" Or BeltType = "ASB-W" Then
    F_Long_Leg.Visible = False
    F_Tri_Leg_Len.Visible = False
    S_Long_Leg.Visible = False
    S_Tri_Leg_Len.Visible = False
    P_G.Visible = False
    P_Leg_Len.Visible = False
    RH_Spiral = "A Spiral"
    LH_Spiral = "B Spiral"
    RH_Spiral = True
  Else
    F_Long_Leg.Visible = True
    F_Tri_Leg_Len.Visible = True
    S_Long_Leg.Visible = True
    S_Tri_Leg_Len.Visible = True
    P_G.Visible = True
    P_Leg_Len.Visible = True
  End If
  
  Data_Dump.Visible = (Employee = "test" Or Employee = 14114)
  Machine_No = ""

  Validator

End Sub
