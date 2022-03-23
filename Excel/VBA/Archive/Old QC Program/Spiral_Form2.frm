VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Spiral_Form 
   Caption         =   "ABI Inspection Program"
   ClientHeight    =   10980
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   20415
   OleObjectBlob   =   "Spiral_Form2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Spiral_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Belt_Width As Variant
Public Mesh_Desc As String
Public Center_Link_Location As Variant
Public Fabric_Width As Variant
Private Sub Close_Form_Click()
    Unload Spiral_Form
    Unload Start_Screen
End Sub
Private Sub Data_Dump_Click()
    Set objMyConn = New ADODB.Connection
    Set UD10 = New ADODB.Recordset
    Dim strSQL As String
    Dim Table As String: Table = "ice.UD10"
    Dim Item As String
    Dim i As Integer: i = 2
    Dim Temp As String

    Worksheets("Graphical Analysis").Cells.Clear
    
    Worksheets("Graphical Analysis").Range("A1") = "Job Number"
    Worksheets("Graphical Analysis").Range("B1") = CalcSheet.Range("J" & 7)
    Worksheets("Graphical Analysis").Range("C1") = "Min"
    Worksheets("Graphical Analysis").Range("D1") = "Target"
    Worksheets("Graphical Analysis").Range("E1") = "Max"
    Worksheets("Graphical Analysis").Range("F1") = CalcSheet.Range("J" & 8)
    Worksheets("Graphical Analysis").Range("G1") = "Min"
    Worksheets("Graphical Analysis").Range("H1") = "Target"
    Worksheets("Graphical Analysis").Range("I1") = "Max"
    Worksheets("Graphical Analysis").Range("J1") = CalcSheet.Range("J" & 9)
    Worksheets("Graphical Analysis").Range("K1") = "Min"
    Worksheets("Graphical Analysis").Range("L1") = "Target"
    Worksheets("Graphical Analysis").Range("M1") = "Max"
    Worksheets("Graphical Analysis").Range("N1") = CalcSheet.Range("J" & 10)
    Worksheets("Graphical Analysis").Range("O1") = "Min"
    Worksheets("Graphical Analysis").Range("P1") = "Target"
    Worksheets("Graphical Analysis").Range("Q1") = "Max"
    Worksheets("Graphical Analysis").Range("R1") = CalcSheet.Range("J" & 11)
    Worksheets("Graphical Analysis").Range("S1") = "Min"
    Worksheets("Graphical Analysis").Range("T1") = "Target"
    Worksheets("Graphical Analysis").Range("U1") = "Max"
    Worksheets("Graphical Analysis").Range("V1") = CalcSheet.Range("J" & 12)
    Worksheets("Graphical Analysis").Range("W1") = "Min"
    Worksheets("Graphical Analysis").Range("X1") = "Target"
    Worksheets("Graphical Analysis").Range("Y1") = "Max"
    Worksheets("Graphical Analysis").Range("Z1") = CalcSheet.Range("J" & 13)
    Worksheets("Graphical Analysis").Range("AA1") = "Min"
    Worksheets("Graphical Analysis").Range("AB1") = "Target"
    Worksheets("Graphical Analysis").Range("AC1") = "Max"
    Worksheets("Graphical Analysis").Range("AD1") = CalcSheet.Range("J" & 14)
    Worksheets("Graphical Analysis").Range("AE1") = "Min"
    Worksheets("Graphical Analysis").Range("AF1") = "Target"
    Worksheets("Graphical Analysis").Range("AG1") = "Max"
    Worksheets("Graphical Analysis").Range("AH1") = CalcSheet.Range("J" & 15)
    Worksheets("Graphical Analysis").Range("AI1") = "Min"
    Worksheets("Graphical Analysis").Range("AJ1") = "Target"
    Worksheets("Graphical Analysis").Range("AK1") = "Max"
    Worksheets("Graphical Analysis").Range("AL1") = CalcSheet.Range("J" & 16)
    Worksheets("Graphical Analysis").Range("AM1") = "Min"
    Worksheets("Graphical Analysis").Range("AN1") = "Target"
    Worksheets("Graphical Analysis").Range("AO1") = "Max"
    Worksheets("Graphical Analysis").Range("AP1") = CalcSheet.Range("J" & 17)
    Worksheets("Graphical Analysis").Range("AQ1") = "Min"
    Worksheets("Graphical Analysis").Range("AR1") = "Target"
    Worksheets("Graphical Analysis").Range("AS1") = "Max"
    Worksheets("Graphical Analysis").Range("AT1") = CalcSheet.Range("J" & 18)
    Worksheets("Graphical Analysis").Range("AU1") = "Min"
    Worksheets("Graphical Analysis").Range("AV1") = "Target"
    Worksheets("Graphical Analysis").Range("AW1") = "Max"
    Worksheets("Graphical Analysis").Range("AX1") = CalcSheet.Range("J" & 19)
    Worksheets("Graphical Analysis").Range("AY1") = "Min"
    Worksheets("Graphical Analysis").Range("AZ1") = "Target"
    Worksheets("Graphical Analysis").Range("BA1") = "Max"
    Worksheets("Graphical Analysis").Range("BB1") = CalcSheet.Range("J" & 20)
    Worksheets("Graphical Analysis").Range("BC1") = "Min"
    Worksheets("Graphical Analysis").Range("BD1") = "Target"
    Worksheets("Graphical Analysis").Range("BE1") = "Max"
    Worksheets("Graphical Analysis").Range("BF1") = CalcSheet.Range("J" & 21)
    Worksheets("Graphical Analysis").Range("BG1") = "Min"
    Worksheets("Graphical Analysis").Range("BH1") = "Target"
    Worksheets("Graphical Analysis").Range("BI1") = "Max"

    objMyConn.ConnectionString = "Provider=SQLOLEDB;Data Source=esql.ashworth.com;Initial Catalog=Ashworth;User ID=devapp;Password=d3v@PP;" 'Connection String
    objMyConn.Open 'Open Connection
    strSQL = "SELECT * FROM " & Table & " WHERE Key1 = '" & JobNum & "' AND Key2 = '" & Insp_Type & " " & Operation & "' And Checkbox20 = '0'"
    Set UD10.ActiveConnection = objMyConn 'Set Connection
    UD10.Open strSQL 'Open SQL Recordset
    
    If UD10.EOF = True Or UD10.BOF = True Then
        MsgBox ("No Data is Available")
        Exit Sub
    End If
    UD10.MoveFirst
    
    While UD10.EOF = False
    Worksheets("Graphical Analysis").Range("A" & i) = UD10.Fields("Key1")
    Worksheets("Graphical Analysis").Range("B" & i) = UD10.Fields("Number01").Value
    Worksheets("Graphical Analysis").Range("C" & i) = CalcSheet.Range("N" & 7) + CalcSheet.Range("L" & 7)
    Worksheets("Graphical Analysis").Range("D" & i) = CalcSheet.Range("L" & 7)
    Worksheets("Graphical Analysis").Range("E" & i) = CalcSheet.Range("Q" & 7) + CalcSheet.Range("L" & 7)
    Worksheets("Graphical Analysis").Range("F" & i) = UD10.Fields("Number02")
    Worksheets("Graphical Analysis").Range("G" & i) = CalcSheet.Range("N" & 8) + CalcSheet.Range("L" & 8)
    Worksheets("Graphical Analysis").Range("H" & i) = CalcSheet.Range("L" & 8)
    Worksheets("Graphical Analysis").Range("I" & i) = CalcSheet.Range("Q" & 8) + CalcSheet.Range("L" & 8)
    Worksheets("Graphical Analysis").Range("J" & i) = UD10.Fields("Number03")
    Worksheets("Graphical Analysis").Range("K" & i) = CalcSheet.Range("N" & 9) + CalcSheet.Range("L" & 9)
    Worksheets("Graphical Analysis").Range("L" & i) = CalcSheet.Range("L" & 9)
    Worksheets("Graphical Analysis").Range("M" & i) = CalcSheet.Range("Q" & 9) + CalcSheet.Range("L" & 9)
    Worksheets("Graphical Analysis").Range("N" & i) = UD10.Fields("Number04")
    Worksheets("Graphical Analysis").Range("O" & i) = CalcSheet.Range("N" & 10) + CalcSheet.Range("L" & 10)
    Worksheets("Graphical Analysis").Range("P" & i) = CalcSheet.Range("L" & 10)
    Worksheets("Graphical Analysis").Range("Q" & i) = CalcSheet.Range("Q" & 10) + CalcSheet.Range("L" & 10)
    Worksheets("Graphical Analysis").Range("R" & i) = UD10.Fields("Number05")
    Worksheets("Graphical Analysis").Range("S" & i) = CalcSheet.Range("N" & 11) + CalcSheet.Range("L" & 11)
    Worksheets("Graphical Analysis").Range("T" & i) = CalcSheet.Range("L" & 11)
    Worksheets("Graphical Analysis").Range("U" & i) = CalcSheet.Range("Q" & 11) + CalcSheet.Range("L" & 11)
    Worksheets("Graphical Analysis").Range("V" & i) = UD10.Fields("Number06")
    Worksheets("Graphical Analysis").Range("W" & i) = CalcSheet.Range("N" & 12) + CalcSheet.Range("L" & 12)
    Worksheets("Graphical Analysis").Range("X" & i) = CalcSheet.Range("L" & 12)
    Worksheets("Graphical Analysis").Range("Y" & i) = CalcSheet.Range("Q" & 12) + CalcSheet.Range("L" & 12)
    Worksheets("Graphical Analysis").Range("Z" & i) = UD10.Fields("Number07")
    Worksheets("Graphical Analysis").Range("AA" & i) = CalcSheet.Range("N" & 13) + CalcSheet.Range("L" & 13)
    Worksheets("Graphical Analysis").Range("AB" & i) = CalcSheet.Range("L" & 13)
    Worksheets("Graphical Analysis").Range("AC" & i) = CalcSheet.Range("Q" & 13) + CalcSheet.Range("L" & 13)
    Worksheets("Graphical Analysis").Range("AD" & i) = UD10.Fields("Number08")
    Worksheets("Graphical Analysis").Range("AE" & i) = CalcSheet.Range("N" & 14) + CalcSheet.Range("L" & 14)
    Worksheets("Graphical Analysis").Range("AF" & i) = CalcSheet.Range("L" & 14)
    Worksheets("Graphical Analysis").Range("AG" & i) = CalcSheet.Range("Q" & 14) + CalcSheet.Range("L" & 14)
    Worksheets("Graphical Analysis").Range("AH" & i) = UD10.Fields("Number09")
    Worksheets("Graphical Analysis").Range("AI" & i) = CalcSheet.Range("N" & 15) + CalcSheet.Range("L" & 15)
    Worksheets("Graphical Analysis").Range("AJ" & i) = CalcSheet.Range("L" & 15)
    Worksheets("Graphical Analysis").Range("AK" & i) = CalcSheet.Range("Q" & 15) + CalcSheet.Range("L" & 15)
    Worksheets("Graphical Analysis").Range("AL" & i) = UD10.Fields("Number10")
    Worksheets("Graphical Analysis").Range("AM" & i) = CalcSheet.Range("N" & 16) + CalcSheet.Range("L" & 16)
    Worksheets("Graphical Analysis").Range("AN" & i) = CalcSheet.Range("L" & 16)
    Worksheets("Graphical Analysis").Range("AO" & i) = CalcSheet.Range("Q" & 16) + CalcSheet.Range("L" & 16)
    Worksheets("Graphical Analysis").Range("AP" & i) = UD10.Fields("Number11")
    Worksheets("Graphical Analysis").Range("AQ" & i) = CalcSheet.Range("N" & 17) + CalcSheet.Range("L" & 17)
    Worksheets("Graphical Analysis").Range("AR" & i) = CalcSheet.Range("L" & 17)
    Worksheets("Graphical Analysis").Range("AS" & i) = CalcSheet.Range("Q" & 17) + CalcSheet.Range("L" & 17)
    Worksheets("Graphical Analysis").Range("AT" & i) = UD10.Fields("Number12")
    Worksheets("Graphical Analysis").Range("AU" & i) = CalcSheet.Range("N" & 18) + CalcSheet.Range("L" & 18)
    Worksheets("Graphical Analysis").Range("AV" & i) = CalcSheet.Range("L" & 18)
    Worksheets("Graphical Analysis").Range("AW" & i) = CalcSheet.Range("Q" & 18) + CalcSheet.Range("L" & 18)
    Worksheets("Graphical Analysis").Range("AX" & i) = UD10.Fields("Number13")
    Worksheets("Graphical Analysis").Range("AY" & i) = CalcSheet.Range("N" & 19) + CalcSheet.Range("L" & 19)
    Worksheets("Graphical Analysis").Range("AZ" & i) = CalcSheet.Range("L" & 19)
    Worksheets("Graphical Analysis").Range("BA" & i) = CalcSheet.Range("Q" & 19) + CalcSheet.Range("L" & 19)
    Worksheets("Graphical Analysis").Range("BB" & i) = UD10.Fields("Number14")
    Worksheets("Graphical Analysis").Range("BC" & i) = CalcSheet.Range("N" & 20) + CalcSheet.Range("L" & 20)
    Worksheets("Graphical Analysis").Range("BD" & i) = CalcSheet.Range("L" & 20)
    Worksheets("Graphical Analysis").Range("BE" & i) = CalcSheet.Range("Q" & 20) + CalcSheet.Range("L" & 20)
    Worksheets("Graphical Analysis").Range("BF" & i) = UD10.Fields("Number15")
    Worksheets("Graphical Analysis").Range("BG" & i) = CalcSheet.Range("N" & 21) + CalcSheet.Range("L" & 21)
    Worksheets("Graphical Analysis").Range("BH" & i) = CalcSheet.Range("L" & 21)
    Worksheets("Graphical Analysis").Range("BI" & i) = CalcSheet.Range("Q" & 21) + CalcSheet.Range("L" & 21)
    UD10.MoveNext
    i = i + 1
    Wend
    
    objMyConn.Close
    MsgBox ("Complete")
    'Call Unhide_Sheet("Graphical Analysis")
End Sub
Private Sub Get_Results_Click()
    Call Create_Job_Results("Grid Spiral Inspection", _
    "#,Date    ,Type,Time      ,Employ,Spec       ,Part #,Spiral Hand,I/O Spiral   ,Height(B),C+ADJ,Long Leg Len,Tri Leg Len,Height(D),Width(E),Height(B),F+ADJ,Long Leg Len,Tri Leg Len,Height(D),Width(E),Diam(G),Leg Len,Fabric Width,Ref,Dog Leg,Burrs,Spiral Twist", _
    "Key3,Date01,Key2,ShortChar06,ShortChar02,Character07,Character02,ShortChar04,ShortChar05,Number01,Number02,Number03,Number04,Number05,Number06,Number07,Number08,Number09,Number10,Number11,Number12,Number13,Number14,Number15,ShortChar01,CheckBox02,CheckBox03,CheckBox04")
End Sub
Private Sub Home_Click()
    Unload Spiral_Form
    Start_Screen.Show
    Sheets("Graphical Analysis").Visible = False
End Sub
Private Sub Image1_Click()
    Call Run_Setup
    Grid_Spiral_Setup.Show
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        MsgBox "The X is disabled, please use a button on the form.", vbCritical
    End If
End Sub
Private Sub Inside_Spiral_Click()
    If Spiral_Form.Inside_Spiral = True And CalcSheet.Range("Spirals_Per_Pitch") = 2 Then
        CalcSheet.Range("IO_Spiral") = Radio_Conversion(Spiral_Form.Inside_Spiral.Value & "," & Spiral_Form.Outside_Spiral.Value, "Inside Spiral" & "," & "Outside Spiral")
        Fabric_Width = ""
        Call Run_Setup
    End If
End Sub
Private Sub Outside_Spiral_Click()
    If Spiral_Form.Outside_Spiral = True And CalcSheet.Range("Spirals_Per_Pitch") = 2 Then
        CalcSheet.Range("IO_Spiral") = Radio_Conversion(Spiral_Form.Inside_Spiral.Value & "," & Spiral_Form.Outside_Spiral.Value, "Inside Spiral" & "," & "Outside Spiral")
        Fabric_Width = ""
        Call Run_Setup
    End If
End Sub

Private Sub RH_Spiral_Click()
    If Belt_Type = "ASB" Or Belt_Type = "ASB-W" Then
        Spiral_Form.Inside_Spiral = True
        Spiral_Form.Outside_Spiral = False
        CalcSheet.Range("IO_Spiral") = Radio_Conversion(Spiral_Form.Inside_Spiral.Value & "," & Spiral_Form.Outside_Spiral.Value, "Inside Spiral" & "," & "Outside Spiral")
        Call Run_Setup
    End If
End Sub

Private Sub LH_Spiral_Click()
    If Belt_Type = "ASB" Or Belt_Type = "ASB-W" Then
        Spiral_Form.Inside_Spiral = False
        Spiral_Form.Outside_Spiral = True
        CalcSheet.Range("IO_Spiral") = Radio_Conversion(Spiral_Form.Inside_Spiral.Value & "," & Spiral_Form.Outside_Spiral.Value, "Inside Spiral" & "," & "Outside Spiral")
        Call Run_Setup
    End If
End Sub

Private Sub Gleen_Info()
  On Error GoTo Err
    Dim Job_Comments As String: Job_Comments = CalcSheet.Range("JobComments")
    Dim Op_Comment As String: Op_Comment = CalcSheet.Range("Operation_Comment")
    Dim GoodValues As Integer
    
    Dim regxo As New regexp, regxo_results As MatchCollection, regxo_subs As SubMatches, regxo_units As String
    With regxo
      .MultiLine = True:
      .IgnoreCase = True:
      .Global = True
      .Pattern = Range("RegXBWidth").Value
    End With
    
    If Belt_Width = "" Or Belt_Width = 0 Then
        Set regxo_results = regxo.Execute(Job_Comments)
        Set regxo_subs = regxo_results(0).SubMatches
        Belt_Width = regxo_subs(0)
        regxo_units = regxo_subs(1)
        If InStr(0, regxo_units, "m", vbTextCompare) Then Belt_Width = Belt_Width / 25.4
        GoodValues = GoodValues + 1
        CalcSheet.Range("Belt_Width") = Belt_Width
    End If
    
    If Len(Mesh_Desc) = 0 Then
        regxo.Pattern = Range("RegXMeshDesc").Value
        Set regxo_results = regxo.Execute(Job_Comments)
        Set regxo_subs = regxo_results(0).SubMatches
        Mesh_Desc = regxo_subs(0)
        GoodValues = GoodValues + 1
        CalcSheet.Range("MeshDesc") = Mesh_Desc
    End If
    
    If Center_Link_Location = "" And Spirals_Per_Pitch > 1 Then
        regxo.Pattern = Range("RegXLinkLoc").Value
        Set regxo_results = regxo.Execute(Job_Comments)
        Set regxo_subs = regxo_results(0).SubMatches
        Let regxo_units = regxo_subs(1)
        Center_Link_Location = regxo_subs(0)
        GoodValues = GoodValues + 1
        CalcSheet.Range("Center_Link_Location") = Center_Link_Location
    End If
    
    If (Fabric_Width = "" Or Fabric_Width = 0) And (Spiral_Form.Inside_Spiral Or Spiral_Form.Outside_Spiral) Then
        regxo.Pattern = Range("RegXLinkLoc").Value
        Set regxo_results = regxo.Execute(Job_Comments)
        Set regxo_subs = regxo_results(0).SubMatches
        Let regxo_units = regxo_subs(1)
        Fabric_Width = regxo_subs(0)
        If InStr(0, regxo_units, "m", vbTextCompare) Then Fabric_Width = Fabric_Width / 25.4
        GoodValues = GoodValues + 1
        CalcSheet.Range("Fabric_Width") = Fabric_Width
    End If
    
    Exit Sub
Err:
    Call MsgBox((4 - GoodValues) & " Values Good", vbExclamation, "Error")
    
End Sub

Private Sub Entry_Test()
    'Check To Make Sure Values are Correct Format
    If IsNumeric(Fabric_Width) = False Then
        Fabric_Width = ""
        Call Run_Setup
    End If
    
    If Spirals_Per_Pitch > 1 And IsNumeric(Center_Link_Location) = False Then
        Center_Link_Location = ""
        Call Run_Setup
    End If
    
    If IsNumeric(Belt_Width) = False Then
        Belt_Width = ""
        Call Run_Setup
    End If
    
    If IsNumeric(CalcSheet.Range("FirstCount")) = False Or IsNumeric(CalcSheet.Range("MeshGauge")) = False Then
        Mesh_Desc = ""
        Call Run_Setup
    End If
End Sub
Private Sub Run_Setup()
    While ((Fabric_Width = 0 Or Fabric_Width = "") And (Spiral_Form.Inside_Spiral Or Spiral_Form.Outside_Spiral)) Or Mesh_Desc = "" Or (Center_Link_Location = "" And Spirals_Per_Pitch > 1)
        Call Gleen_Info
    Wend
    
    Call Entry_Test
End Sub
Private Sub Submit_Click()
    On Error GoTo GeneralError
    Dim Inspection_Name As String: Inspection_Name = "Spiral_Inspection"
    
    If Spiral_Form.Machine_No.Value > "" And (Spiral_Form.LH_Spiral Or Spiral_Form.RH_Spiral) Then
    Else
        GoTo BlankForm
    End If
    
    Call Clear_Sheet()
    Call Run_Setup

    'Write Data To Excel Sheet
    CalcSheet.Range("Insp_Plan") = CalcSheet.Range(Inspection_Name & "_Plan")
    CalcSheet.Range("Spec_ID") = CalcSheet.Range(Inspection_Name & "_Spec")
    CalcSheet.Range("Data1") = Number_Conversion(Spiral_Form.F_B.Value)
    CalcSheet.Range("Data2") = Number_Conversion(Spiral_Form.F_C_ADJ.Value)
    CalcSheet.Range("Data3") = Number_Conversion(Spiral_Form.F_Long_Leg.Value)
    CalcSheet.Range("Data4") = Number_Conversion(Spiral_Form.F_Tri_Leg_Len.Value)
    CalcSheet.Range("Data5") = Number_Conversion(Spiral_Form.F_D.Value)
    CalcSheet.Range("Data6") = Number_Conversion(Spiral_Form.F_E.Value)
    CalcSheet.Range("Data7") = Number_Conversion(Spiral_Form.S_B.Value)
    CalcSheet.Range("Data8") = Number_Conversion(Spiral_Form.S_C_ADJ.Value)
    CalcSheet.Range("Data9") = Number_Conversion(Spiral_Form.S_Long_Leg.Value)
    CalcSheet.Range("Data10") = Number_Conversion(Spiral_Form.S_Tri_Leg_Len.Value)
    CalcSheet.Range("Data11") = Number_Conversion(Spiral_Form.S_D.Value)
    CalcSheet.Range("Data12") = Number_Conversion(Spiral_Form.S_E.Value)
    CalcSheet.Range("Data13") = Number_Conversion(Spiral_Form.P_G.Value)
    CalcSheet.Range("Data14") = Number_Conversion(Spiral_Form.P_Leg_Len.Value)
    CalcSheet.Range("Data15") = Number_Conversion(Spiral_Form.O_Fab_Wid.Value)
    CalcSheet.Range("Schar1") = Trim(Spiral_Form.O_Ref.Value)
    CalcSheet.Range("Schar3") = Trim(Spiral_Form.Machine_No.Value)
    
    If Belt_Type = "ASB" Or Belt_Type = "ASB-W" Then
        CalcSheet.Range("Schar4") = Radio_Conversion(Spiral_Form.LH_Spiral.Value & "," & Spiral_Form.RH_Spiral.Value, "Spiral B" & "," & "Spiral A")
    Else
        CalcSheet.Range("Schar4") = Radio_Conversion(Spiral_Form.LH_Spiral.Value & "," & Spiral_Form.RH_Spiral.Value, "LH Spiral" & "," & "RH Spiral")
    End If
    
    CalcSheet.Range("Check2") = Radio_Conversion(Spiral_Form.Dog_Leg.Value & "," & "True", 1 & "," & 0)
    CalcSheet.Range("Check3") = Radio_Conversion(Spiral_Form.Burrs.Value & "," & "True", 1 & "," & 0)
    CalcSheet.Range("Check4") = Radio_Conversion(Spiral_Form.Spiral_Twist.Value & "," & "True", 1 & "," & 0)

    'Results
    If CalcSheet.Range(Inspection_Name & "_Comment") = Empty Then
        CalcSheet.Range("Passed") = 1
        CalcSheet.Range("Value") = ""
        CalcSheet.Range("Failed_Comment") = ""
    ElseIf IsError(CalcSheet.Range(Inspection_Name & "_Comment")) Then
        GoTo MissingData
        MsgBox ("HERE")
    Else
        Temp = Replace(CalcSheet.Range(Inspection_Name & "_Comment"), "?", ".  ")
        CalcSheet.Range("Passed") = 0
        CalcSheet.Range("Failed_Comment") = Temp
        CalcSheet.Range("Value") = "Spiral Rejected"
        Call Rejection_Form(Replace(CalcSheet.Range(Inspection_Name & "_Comment"), "?", vbNewLine))
    End If

    Call Write_Data()

    Spiral_Form.F_B.Value = ""
    Spiral_Form.F_C_ADJ.Value = ""
    Spiral_Form.F_Long_Leg.Value = ""
    Spiral_Form.F_Tri_Leg_Len = ""
    Spiral_Form.F_D.Value = ""
    Spiral_Form.F_E.Value = ""
    Spiral_Form.S_B.Value = ""
    Spiral_Form.S_C_ADJ.Value = ""
    Spiral_Form.S_Long_Leg.Value = ""
    Spiral_Form.S_Tri_Leg_Len.Value = ""
    Spiral_Form.S_D.Value = ""
    Spiral_Form.S_E.Value = ""
    Spiral_Form.P_G.Value = ""
    Spiral_Form.P_Leg_Len.Value = ""
    Spiral_Form.O_Fab_Wid.Value = ""
    Spiral_Form.O_Ref.Value = ""
    Spiral_Form.Dog_Leg.Value = False
    Spiral_Form.Burrs.Value = False
    Spiral_Form.Spiral_Twist = False
    Inspection_Num.Caption = "Inspection Num: " + CStr(CalcSheet.Range("SampleNum") + 1)
    Spiral_Form.F_B.SetFocus
    Exit Sub

GeneralError:
MsgBox ("Errors have been Detected in the Data Entered Please Correct and Resubmit")
Exit Sub

BlankForm:
MsgBox ("The Spiral Hand and Machine Number Must Be Entered.  Please Correct and Resubmit")
Exit Sub

MissingData:
MsgBox ("Data Required For This Inspection is Missing Please Call x1329.")
Exit Sub

End Sub
Private Sub UserForm_Activate()
    Call Resize_Screen(Spiral_Form)
    Call Clear_Calcs()
            
    'Get Specs
    Dim ObjMyConn: Set objMyConn = New ADODB.Connection
    Dim JobOper: Set JobOper = New ADODB.Recordset
    Dim strSQL As String
    Dim Table As String: Table = "erp.JobOper"
    Dim Item As String

    objMyConn.ConnectionString = "Provider=SQLOLEDB;Data Source=esql.ashworth.com;Initial Catalog=Ashworth;User ID=devapp;Password=d3v@PP;" 'Connection String
    objMyConn.Open 'Open Connection
    strSQL = "SELECT * FROM " & Table & " WHERE Company = 200 AND JobNum = '" & JobNum & "' AND OpCode = 'GBDSPR01'"
    
    Set JobOper.ActiveConnection = objMyConn 'Set Connection
    JobOper.Open strSQL 'Open SQL Recordset

    If JobOper.EOF = True Or JobOper.BOF = True Then
        MsgBox ("Error Returning the Operation Comments")
    End If
    
    CalcSheet.Range("Operation_Comment") = Comment_Format(JobOper.Fields("CommentText").Value)
    objMyConn.Close
            
    InspNum = CalcSheet.Range("SampleNum")
    InspNum = InspNum + 1
    Inspection_Num.Caption = "Inspection Num: " + CStr(InspNum)

    'Hide Inside Outside Selection Box if Belt only has one spiral per pitch
    If Spirals_Per_Pitch = 1 Then
        Spiral_Form.Inside_Spiral.Visible = False
        Spiral_Form.Outside_Spiral.Visible = False
        Spiral_Form.Inside_Spiral.Value = True
    Else
        Spiral_Form.Inside_Spiral.Visible = True
        Spiral_Form.Outside_Spiral.Visible = True
        Spiral_Form.Inside_Spiral.Value = False
    End If
    
    'Setup For Stacker Belts
    If Belt_Type = "ASB" Or Belt_Type = "ASB-W" Then
        Spiral_Form.F_Long_Leg.Visible = False
        Spiral_Form.F_Tri_Leg_Len.Visible = False
        Spiral_Form.S_Long_Leg.Visible = False
        Spiral_Form.S_Tri_Leg_Len.Visible = False
        Spiral_Form.P_G.Visible = False
        Spiral_Form.P_Leg_Len.Visible = False
        Spiral_Form.RH_Spiral.Caption = "A Spiral"
        Spiral_Form.LH_Spiral.Caption = "B Spiral"
        Spiral_Form.RH_Spiral = True
    Else
        Spiral_Form.F_Long_Leg.Visible = True
        Spiral_Form.F_Tri_Leg_Len.Visible = True
        Spiral_Form.S_Long_Leg.Visible = True
        Spiral_Form.S_Tri_Leg_Len.Visible = True
        Spiral_Form.P_G.Visible = True
        Spiral_Form.P_Leg_Len.Visible = True
    End If
    
    If Application.UserName = "Black, Norman" Or Application.UserName = "witrinc" Or Application.UserName = "Ryan Taylor" Then
        Spiral_Form.Data_Dump.Visible = True
    Else
        Spiral_Form.Data_Dump.Visible = False
    End If
    
    Spiral_Form.Machine_No = ""

    Call Run_Setup
End Sub
