VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Spiral_Form 
   Caption         =   "ABI Inspection Program"
   ClientHeight    =   10980
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   20415
   OleObjectBlob   =   "Spiral_Form.frx":0000
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
    Worksheets("Graphical Analysis").Range("B1") = Worksheets("Calculations").Range("J" & 7)
    Worksheets("Graphical Analysis").Range("C1") = "Min"
    Worksheets("Graphical Analysis").Range("D1") = "Target"
    Worksheets("Graphical Analysis").Range("E1") = "Max"
    Worksheets("Graphical Analysis").Range("F1") = Worksheets("Calculations").Range("J" & 8)
    Worksheets("Graphical Analysis").Range("G1") = "Min"
    Worksheets("Graphical Analysis").Range("H1") = "Target"
    Worksheets("Graphical Analysis").Range("I1") = "Max"
    Worksheets("Graphical Analysis").Range("J1") = Worksheets("Calculations").Range("J" & 9)
    Worksheets("Graphical Analysis").Range("K1") = "Min"
    Worksheets("Graphical Analysis").Range("L1") = "Target"
    Worksheets("Graphical Analysis").Range("M1") = "Max"
    Worksheets("Graphical Analysis").Range("N1") = Worksheets("Calculations").Range("J" & 10)
    Worksheets("Graphical Analysis").Range("O1") = "Min"
    Worksheets("Graphical Analysis").Range("P1") = "Target"
    Worksheets("Graphical Analysis").Range("Q1") = "Max"
    Worksheets("Graphical Analysis").Range("R1") = Worksheets("Calculations").Range("J" & 11)
    Worksheets("Graphical Analysis").Range("S1") = "Min"
    Worksheets("Graphical Analysis").Range("T1") = "Target"
    Worksheets("Graphical Analysis").Range("U1") = "Max"
    Worksheets("Graphical Analysis").Range("V1") = Worksheets("Calculations").Range("J" & 12)
    Worksheets("Graphical Analysis").Range("W1") = "Min"
    Worksheets("Graphical Analysis").Range("X1") = "Target"
    Worksheets("Graphical Analysis").Range("Y1") = "Max"
    Worksheets("Graphical Analysis").Range("Z1") = Worksheets("Calculations").Range("J" & 13)
    Worksheets("Graphical Analysis").Range("AA1") = "Min"
    Worksheets("Graphical Analysis").Range("AB1") = "Target"
    Worksheets("Graphical Analysis").Range("AC1") = "Max"
    Worksheets("Graphical Analysis").Range("AD1") = Worksheets("Calculations").Range("J" & 14)
    Worksheets("Graphical Analysis").Range("AE1") = "Min"
    Worksheets("Graphical Analysis").Range("AF1") = "Target"
    Worksheets("Graphical Analysis").Range("AG1") = "Max"
    Worksheets("Graphical Analysis").Range("AH1") = Worksheets("Calculations").Range("J" & 15)
    Worksheets("Graphical Analysis").Range("AI1") = "Min"
    Worksheets("Graphical Analysis").Range("AJ1") = "Target"
    Worksheets("Graphical Analysis").Range("AK1") = "Max"
    Worksheets("Graphical Analysis").Range("AL1") = Worksheets("Calculations").Range("J" & 16)
    Worksheets("Graphical Analysis").Range("AM1") = "Min"
    Worksheets("Graphical Analysis").Range("AN1") = "Target"
    Worksheets("Graphical Analysis").Range("AO1") = "Max"
    Worksheets("Graphical Analysis").Range("AP1") = Worksheets("Calculations").Range("J" & 17)
    Worksheets("Graphical Analysis").Range("AQ1") = "Min"
    Worksheets("Graphical Analysis").Range("AR1") = "Target"
    Worksheets("Graphical Analysis").Range("AS1") = "Max"
    Worksheets("Graphical Analysis").Range("AT1") = Worksheets("Calculations").Range("J" & 18)
    Worksheets("Graphical Analysis").Range("AU1") = "Min"
    Worksheets("Graphical Analysis").Range("AV1") = "Target"
    Worksheets("Graphical Analysis").Range("AW1") = "Max"
    Worksheets("Graphical Analysis").Range("AX1") = Worksheets("Calculations").Range("J" & 19)
    Worksheets("Graphical Analysis").Range("AY1") = "Min"
    Worksheets("Graphical Analysis").Range("AZ1") = "Target"
    Worksheets("Graphical Analysis").Range("BA1") = "Max"
    Worksheets("Graphical Analysis").Range("BB1") = Worksheets("Calculations").Range("J" & 20)
    Worksheets("Graphical Analysis").Range("BC1") = "Min"
    Worksheets("Graphical Analysis").Range("BD1") = "Target"
    Worksheets("Graphical Analysis").Range("BE1") = "Max"
    Worksheets("Graphical Analysis").Range("BF1") = Worksheets("Calculations").Range("J" & 21)
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
    Worksheets("Graphical Analysis").Range("C" & i) = Worksheets("Calculations").Range("N" & 7) + Worksheets("Calculations").Range("L" & 7)
    Worksheets("Graphical Analysis").Range("D" & i) = Worksheets("Calculations").Range("L" & 7)
    Worksheets("Graphical Analysis").Range("E" & i) = Worksheets("Calculations").Range("Q" & 7) + Worksheets("Calculations").Range("L" & 7)
    Worksheets("Graphical Analysis").Range("F" & i) = UD10.Fields("Number02")
    Worksheets("Graphical Analysis").Range("G" & i) = Worksheets("Calculations").Range("N" & 8) + Worksheets("Calculations").Range("L" & 8)
    Worksheets("Graphical Analysis").Range("H" & i) = Worksheets("Calculations").Range("L" & 8)
    Worksheets("Graphical Analysis").Range("I" & i) = Worksheets("Calculations").Range("Q" & 8) + Worksheets("Calculations").Range("L" & 8)
    Worksheets("Graphical Analysis").Range("J" & i) = UD10.Fields("Number03")
    Worksheets("Graphical Analysis").Range("K" & i) = Worksheets("Calculations").Range("N" & 9) + Worksheets("Calculations").Range("L" & 9)
    Worksheets("Graphical Analysis").Range("L" & i) = Worksheets("Calculations").Range("L" & 9)
    Worksheets("Graphical Analysis").Range("M" & i) = Worksheets("Calculations").Range("Q" & 9) + Worksheets("Calculations").Range("L" & 9)
    Worksheets("Graphical Analysis").Range("N" & i) = UD10.Fields("Number04")
    Worksheets("Graphical Analysis").Range("O" & i) = Worksheets("Calculations").Range("N" & 10) + Worksheets("Calculations").Range("L" & 10)
    Worksheets("Graphical Analysis").Range("P" & i) = Worksheets("Calculations").Range("L" & 10)
    Worksheets("Graphical Analysis").Range("Q" & i) = Worksheets("Calculations").Range("Q" & 10) + Worksheets("Calculations").Range("L" & 10)
    Worksheets("Graphical Analysis").Range("R" & i) = UD10.Fields("Number05")
    Worksheets("Graphical Analysis").Range("S" & i) = Worksheets("Calculations").Range("N" & 11) + Worksheets("Calculations").Range("L" & 11)
    Worksheets("Graphical Analysis").Range("T" & i) = Worksheets("Calculations").Range("L" & 11)
    Worksheets("Graphical Analysis").Range("U" & i) = Worksheets("Calculations").Range("Q" & 11) + Worksheets("Calculations").Range("L" & 11)
    Worksheets("Graphical Analysis").Range("V" & i) = UD10.Fields("Number06")
    Worksheets("Graphical Analysis").Range("W" & i) = Worksheets("Calculations").Range("N" & 12) + Worksheets("Calculations").Range("L" & 12)
    Worksheets("Graphical Analysis").Range("X" & i) = Worksheets("Calculations").Range("L" & 12)
    Worksheets("Graphical Analysis").Range("Y" & i) = Worksheets("Calculations").Range("Q" & 12) + Worksheets("Calculations").Range("L" & 12)
    Worksheets("Graphical Analysis").Range("Z" & i) = UD10.Fields("Number07")
    Worksheets("Graphical Analysis").Range("AA" & i) = Worksheets("Calculations").Range("N" & 13) + Worksheets("Calculations").Range("L" & 13)
    Worksheets("Graphical Analysis").Range("AB" & i) = Worksheets("Calculations").Range("L" & 13)
    Worksheets("Graphical Analysis").Range("AC" & i) = Worksheets("Calculations").Range("Q" & 13) + Worksheets("Calculations").Range("L" & 13)
    Worksheets("Graphical Analysis").Range("AD" & i) = UD10.Fields("Number08")
    Worksheets("Graphical Analysis").Range("AE" & i) = Worksheets("Calculations").Range("N" & 14) + Worksheets("Calculations").Range("L" & 14)
    Worksheets("Graphical Analysis").Range("AF" & i) = Worksheets("Calculations").Range("L" & 14)
    Worksheets("Graphical Analysis").Range("AG" & i) = Worksheets("Calculations").Range("Q" & 14) + Worksheets("Calculations").Range("L" & 14)
    Worksheets("Graphical Analysis").Range("AH" & i) = UD10.Fields("Number09")
    Worksheets("Graphical Analysis").Range("AI" & i) = Worksheets("Calculations").Range("N" & 15) + Worksheets("Calculations").Range("L" & 15)
    Worksheets("Graphical Analysis").Range("AJ" & i) = Worksheets("Calculations").Range("L" & 15)
    Worksheets("Graphical Analysis").Range("AK" & i) = Worksheets("Calculations").Range("Q" & 15) + Worksheets("Calculations").Range("L" & 15)
    Worksheets("Graphical Analysis").Range("AL" & i) = UD10.Fields("Number10")
    Worksheets("Graphical Analysis").Range("AM" & i) = Worksheets("Calculations").Range("N" & 16) + Worksheets("Calculations").Range("L" & 16)
    Worksheets("Graphical Analysis").Range("AN" & i) = Worksheets("Calculations").Range("L" & 16)
    Worksheets("Graphical Analysis").Range("AO" & i) = Worksheets("Calculations").Range("Q" & 16) + Worksheets("Calculations").Range("L" & 16)
    Worksheets("Graphical Analysis").Range("AP" & i) = UD10.Fields("Number11")
    Worksheets("Graphical Analysis").Range("AQ" & i) = Worksheets("Calculations").Range("N" & 17) + Worksheets("Calculations").Range("L" & 17)
    Worksheets("Graphical Analysis").Range("AR" & i) = Worksheets("Calculations").Range("L" & 17)
    Worksheets("Graphical Analysis").Range("AS" & i) = Worksheets("Calculations").Range("Q" & 17) + Worksheets("Calculations").Range("L" & 17)
    Worksheets("Graphical Analysis").Range("AT" & i) = UD10.Fields("Number12")
    Worksheets("Graphical Analysis").Range("AU" & i) = Worksheets("Calculations").Range("N" & 18) + Worksheets("Calculations").Range("L" & 18)
    Worksheets("Graphical Analysis").Range("AV" & i) = Worksheets("Calculations").Range("L" & 18)
    Worksheets("Graphical Analysis").Range("AW" & i) = Worksheets("Calculations").Range("Q" & 18) + Worksheets("Calculations").Range("L" & 18)
    Worksheets("Graphical Analysis").Range("AX" & i) = UD10.Fields("Number13")
    Worksheets("Graphical Analysis").Range("AY" & i) = Worksheets("Calculations").Range("N" & 19) + Worksheets("Calculations").Range("L" & 19)
    Worksheets("Graphical Analysis").Range("AZ" & i) = Worksheets("Calculations").Range("L" & 19)
    Worksheets("Graphical Analysis").Range("BA" & i) = Worksheets("Calculations").Range("Q" & 19) + Worksheets("Calculations").Range("L" & 19)
    Worksheets("Graphical Analysis").Range("BB" & i) = UD10.Fields("Number14")
    Worksheets("Graphical Analysis").Range("BC" & i) = Worksheets("Calculations").Range("N" & 20) + Worksheets("Calculations").Range("L" & 20)
    Worksheets("Graphical Analysis").Range("BD" & i) = Worksheets("Calculations").Range("L" & 20)
    Worksheets("Graphical Analysis").Range("BE" & i) = Worksheets("Calculations").Range("Q" & 20) + Worksheets("Calculations").Range("L" & 20)
    Worksheets("Graphical Analysis").Range("BF" & i) = UD10.Fields("Number15")
    Worksheets("Graphical Analysis").Range("BG" & i) = Worksheets("Calculations").Range("N" & 21) + Worksheets("Calculations").Range("L" & 21)
    Worksheets("Graphical Analysis").Range("BH" & i) = Worksheets("Calculations").Range("L" & 21)
    Worksheets("Graphical Analysis").Range("BI" & i) = Worksheets("Calculations").Range("Q" & 21) + Worksheets("Calculations").Range("L" & 21)
    UD10.MoveNext
    i = i + 1
    Wend
    
    objMyConn.Close
    MsgBox ("Complete")
    'Run_Function = Unhide_Sheet("Graphical Analysis")
End Sub
Private Sub Get_Results_Click()
    Run_Function = Create_Job_Results("Grid Spiral Inspection", _
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
    If Spiral_Form.Inside_Spiral = True And Worksheets("Calculations").Range("Spirals_Per_Pitch") = 2 Then
        Worksheets("Calculations").Range("IO_Spiral") = Radio_Conversion(Spiral_Form.Inside_Spiral.Value & "," & Spiral_Form.Outside_Spiral.Value, "Inside Spiral" & "," & "Outside Spiral")
        Fabric_Width = ""
        Call Run_Setup
    End If
End Sub
Private Sub Outside_Spiral_Click()
    If Spiral_Form.Outside_Spiral = True And Worksheets("Calculations").Range("Spirals_Per_Pitch") = 2 Then
        Worksheets("Calculations").Range("IO_Spiral") = Radio_Conversion(Spiral_Form.Inside_Spiral.Value & "," & Spiral_Form.Outside_Spiral.Value, "Inside Spiral" & "," & "Outside Spiral")
        Fabric_Width = ""
        Call Run_Setup
    End If
End Sub

Private Sub RH_Spiral_Click()
    If Belt_Type = "ASB" Or Belt_Type = "ASB-W" Then
        Spiral_Form.Inside_Spiral = True
        Spiral_Form.Outside_Spiral = False
        Worksheets("Calculations").Range("IO_Spiral") = Radio_Conversion(Spiral_Form.Inside_Spiral.Value & "," & Spiral_Form.Outside_Spiral.Value, "Inside Spiral" & "," & "Outside Spiral")
        Call Run_Setup
    End If
End Sub

Private Sub LH_Spiral_Click()
    If Belt_Type = "ASB" Or Belt_Type = "ASB-W" Then
        Spiral_Form.Inside_Spiral = False
        Spiral_Form.Outside_Spiral = True
        Worksheets("Calculations").Range("IO_Spiral") = Radio_Conversion(Spiral_Form.Inside_Spiral.Value & "," & Spiral_Form.Outside_Spiral.Value, "Inside Spiral" & "," & "Outside Spiral")
        Call Run_Setup
    End If
End Sub



'Small Radius Omni-Grid 100
'Length: 500 Feet
'Overall Belt Width: 40 Inches
'Rod Material: T304
'Application: Cooker
'Mesh Material: T304M
'Mesh Desc.: B48-12/12-16
'Turn Ratio: 1.10
'Cage Diameter: 88  Inches
'Center Link Location: 20  Inches


Private Sub Gleen_Info()
    Dim Job_Comments As String: Job_Comments = Worksheets("Calculations").Range("JobComments")
    Dim Op_Comment As String: Op_Comment = Worksheets("Calculations").Range("Operation_Comment")
    
    If Belt_Width = "" Then
        Belt_Width = Comment_Search("Width", Job_Comments, "Inches", "in.", "", "")
        If IsNumeric(Belt_Width) And Belt_Width > 0 Then
            Worksheets("Calculations").Range("Belt_Width") = Belt_Width
        Else
            Belt_Width = Number_Conversion(Comment_Box("Belt Width", Job_Comments))
            Worksheets("Calculations").Range("Belt_Width") = Belt_Width
        End If
    End If
    
    If Len(Mesh_Desc) = 0 Then
        Mesh_Desc = Comment_Search("Mesh Description :", Op_Comment, "", "", "", "")
        If Len(Mesh_Desc) = 0 Then
            Mesh_Desc = Comment_Search("Mesh:", Op_Comment, "", "", "", "")
        End If
        If Len(Mesh_Desc) = 0 Then
            Mesh_Desc = Comment_Search("Mesh Desc.:", Job_Comments, "", "", "", "")
        End If
        If Len(Mesh_Desc) > 0 Then
            Worksheets("Calculations").Range("MeshDesc") = Mesh_Desc
        Else
            Mesh_Desc = Replace(Comment_Box("Mesh Description (Example B42-24-12)", Op_Comment), " ", "")
            Worksheets("Calculations").Range("MeshDesc") = Mesh_Desc
        End If
    End If
    
    If Spirals_Per_Pitch > 1 And Center_Link_Location = "" Then
        Center_Link_Location = Comment_Search("Center Link Location:", Job_Comments, "inches", "in.", "in", "")
        If Len(Center_Link_Location) = 0 Then
            Center_Link_Location = Comment_Search("Mesh Desc.:", Job_Comments, "Inches", "in", "in.", "")
        End If
        If IsNumeric(Center_Link_Location) And Center_Link_Location > 0 Then
            Worksheets("Calculations").Range("Center_Link_Location") = Center_Link_Location
        Else
            Center_Link_Location = Number_Conversion(Comment_Box("Center Link Location", Job_Comments))
            Worksheets("Calculations").Range("Center_Link_Location") = Center_Link_Location
        End If
    End If
    
    If (Fabric_Width = "" Or Fabric_Width = 0) And (Spiral_Form.Inside_Spiral = True Or Spiral_Form.Outside_Spiral = True) Then
        If Spirals_Per_Picth = 1 Then
            Fabric_Width = Comment_Search("Fabric Width", Job_Comments, "", "", "", "")
        End If
        If Spirals_Per_Picth = 1 Then
            Fabric_Width = Comment_Search("Overall Belt Width:", Job_Comments, "Inches", "in", "in.", "")
        End If
        If IsNumeric(Fabric_Width) And Fabric_Width > 0 Then
            Worksheets("Calculations").Range("Fabric_Width") = Fabric_Width
        Else
            Fabric_Width = Number_Conversion(Comment_Box("Fabric Width", Op_Comment))
            Worksheets("Calculations").Range("Fabric_Width") = Fabric_Width
        End If
    End If
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
    
    If IsNumeric(Worksheets("Calculations").Range("FirstCount")) = False Or IsNumeric(Worksheets("Calculations").Range("MeshGauge")) = False Then
        Mesh_Desc = ""
        Call Run_Setup
    End If
End Sub
Private Sub Run_Setup()
    While ((Fabric_Width = 0 Or Fabric_Width = "") And (Spiral_Form.Inside_Spiral = True Or Spiral_Form.Outside_Spiral = True)) Or Mesh_Desc = "" Or (Center_Link_Location = "" And Spirals_Per_Pitch > 1)
        Call Gleen_Info
    Wend
    
    Call Entry_Test
End Sub
Private Sub Submit_Click()
    On Error GoTo GeneralError
    Dim Inspection_Name As String: Inspection_Name = "Spiral_Inspection"
    
    If Spiral_Form.Machine_No.Value > "" And (Spiral_Form.LH_Spiral = True Or Spiral_Form.RH_Spiral = True) Then
    Else
        GoTo BlankForm
    End If
    
    Run_Function = Clear_Sheet()
    Call Run_Setup

    'Write Data To Excel Sheet
    Worksheets("Calculations").Range("Insp_Plan") = Worksheets("Calculations").Range(Inspection_Name & "_Plan")
    Worksheets("Calculations").Range("Spec_ID") = Worksheets("Calculations").Range(Inspection_Name & "_Spec")
    Worksheets("Calculations").Range("Data1") = Number_Conversion(Spiral_Form.F_B.Value)
    Worksheets("Calculations").Range("Data2") = Number_Conversion(Spiral_Form.F_C_ADJ.Value)
    Worksheets("Calculations").Range("Data3") = Number_Conversion(Spiral_Form.F_Long_Leg.Value)
    Worksheets("Calculations").Range("Data4") = Number_Conversion(Spiral_Form.F_Tri_Leg_Len.Value)
    Worksheets("Calculations").Range("Data5") = Number_Conversion(Spiral_Form.F_D.Value)
    Worksheets("Calculations").Range("Data6") = Number_Conversion(Spiral_Form.F_E.Value)
    Worksheets("Calculations").Range("Data7") = Number_Conversion(Spiral_Form.S_B.Value)
    Worksheets("Calculations").Range("Data8") = Number_Conversion(Spiral_Form.S_C_ADJ.Value)
    Worksheets("Calculations").Range("Data9") = Number_Conversion(Spiral_Form.S_Long_Leg.Value)
    Worksheets("Calculations").Range("Data10") = Number_Conversion(Spiral_Form.S_Tri_Leg_Len.Value)
    Worksheets("Calculations").Range("Data11") = Number_Conversion(Spiral_Form.S_D.Value)
    Worksheets("Calculations").Range("Data12") = Number_Conversion(Spiral_Form.S_E.Value)
    Worksheets("Calculations").Range("Data13") = Number_Conversion(Spiral_Form.P_G.Value)
    Worksheets("Calculations").Range("Data14") = Number_Conversion(Spiral_Form.P_Leg_Len.Value)
    Worksheets("Calculations").Range("Data15") = Number_Conversion(Spiral_Form.O_Fab_Wid.Value)
    Worksheets("Calculations").Range("Schar1") = Trim(Spiral_Form.O_Ref.Value)
    Worksheets("Calculations").Range("Schar3") = Trim(Spiral_Form.Machine_No.Value)
    
    If Belt_Type = "ASB" Or Belt_Type = "ASB-W" Then
        Worksheets("Calculations").Range("Schar4") = Radio_Conversion(Spiral_Form.LH_Spiral.Value & "," & Spiral_Form.RH_Spiral.Value, "Spiral B" & "," & "Spiral A")
    Else
        Worksheets("Calculations").Range("Schar4") = Radio_Conversion(Spiral_Form.LH_Spiral.Value & "," & Spiral_Form.RH_Spiral.Value, "LH Spiral" & "," & "RH Spiral")
    End If
    
    Worksheets("Calculations").Range("Check2") = Radio_Conversion(Spiral_Form.Dog_Leg.Value & "," & "True", 1 & "," & 0)
    Worksheets("Calculations").Range("Check3") = Radio_Conversion(Spiral_Form.Burrs.Value & "," & "True", 1 & "," & 0)
    Worksheets("Calculations").Range("Check4") = Radio_Conversion(Spiral_Form.Spiral_Twist.Value & "," & "True", 1 & "," & 0)

    'Results
    If Worksheets("Calculations").Range(Inspection_Name & "_Comment") = Empty Then
        Worksheets("Calculations").Range("Passed") = 1
        Worksheets("Calculations").Range("Value") = ""
        Worksheets("Calculations").Range("Failed_Comment") = ""
    ElseIf IsError(Worksheets("Calculations").Range(Inspection_Name & "_Comment")) Then
        GoTo MissingData
        MsgBox ("HERE")
    Else
        Temp = Replace(Worksheets("Calculations").Range(Inspection_Name & "_Comment"), "?", ".  ")
        Worksheets("Calculations").Range("Passed") = 0
        Worksheets("Calculations").Range("Failed_Comment") = Temp
        Worksheets("Calculations").Range("Value") = "Spiral Rejected"
        Run_Function = Rejection_Form(Replace(Worksheets("Calculations").Range(Inspection_Name & "_Comment"), "?", vbNewLine))
    End If

    Run_Function = Write_Data()

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
    Inspection_Num.Caption = "Inspection Num: " + CStr(Worksheets("Calculations").Range("SampleNum") + 1)
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
    Run_Function = Resize_Screen(Spiral_Form)
    Run_Function = Clear_Calcs()
            
    'Get Specs
    Set objMyConn = New ADODB.Connection
    Set JobOper = New ADODB.Recordset
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
    
    Worksheets("Calculations").Range("Operation_Comment") = Comment_Format(JobOper.Fields("CommentText").Value)
    objMyConn.Close
            
    InspNum = Worksheets("Calculations").Range("SampleNum")
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
    
    If Application.UserName = "Black, Norman" Or Application.UserName = "witrinc" Then
        Spiral_Form.Data_Dump.Visible = True
    Else
        Spiral_Form.Data_Dump.Visible = False
    End If
    
    Spiral_Form.Machine_No = ""

    Call Run_Setup
End Sub
