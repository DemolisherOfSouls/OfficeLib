VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Flatwire_Picket_Inspection 
   Caption         =   "ABI Inspection Program"
   ClientHeight    =   10980
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   20415
   OleObjectBlob   =   "Flatwire_Picket_Inspection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Flatwire_Picket_Inspection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1

Private Sub Inspection_Values_Click()

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = 0 Then
    Cancel = True
    MsgBox "The X is disabled, please use a button on the form.", vbCritical
  End If
End Sub

'Buttons
Private Sub Get_Results_Click()
  DisplayResults "Flatwire Picket Inspection", _
    "#,Date    ,Type,Time      ,Employ,Spec       ,Part #,Picket Bow (A),Picket Bow (B),T & F (C),Free Picket Width,Picket Compression,IRP,COA,COLS,MTA,Crooked Loop,Scratches,High Corner,Tool Marks,Burrs", _
    "Key3,Date01,Key2,ShortChar06,ShortChar02,Character07,Character02,Number01,Number02,Number03,Number04,Number05,Checkbox02,Checkbox03,Checkbox04,Checkbox05,Checkbox06,Checkbox07,Checkbox08,Checkbox09,Checkbox10"
End Sub
Private Sub Home_Click()
  Unload Flatwire_Picket_Inspection
  Start_Screen.Show
End Sub

'Radio Boxes
Private Sub Inside_Picket_Change()
  Dim PartNumber As String: PartNumber = CalcSheet.Range("LPartNum")
  
  If PartNumber = "SROFG1" Or PartNumber = "SROFG3" Then
    If Inside_Picket = True Then
      'Inside Picket
      Inside_None.Enabled = True
      Inside_Single_SD.Enabled = True
      Inside_Double_SD.Enabled = True
      Inside_Single_HD.Enabled = True
      Inside_Double_HD.Enabled = True
  
      'Outside Picket
      Outside_None = True
      Outside_None.Enabled = False
      Outside_Single_SD.Enabled = False
      Outside_Double_SD.Enabled = False
      Outside_Single_HD.Enabled = False
      Outside_Double_HD.Enabled = False
  
      Loops = 0
      Run_Setup
    End If
  End If
End Sub

Private Sub Outside_Picket_Change()
  Dim PartNumber As String: PartNumber = CalcSheet.Range("LPartNum")

  If (PartNumber = "SROFG1" Or PartNumber = "SROFG3") And Outside_Picket Then
      'Inside Picket
      Inside_None = True
      Inside_None.Enabled = False
      Inside_Single_SD.Enabled = False
      Inside_Double_SD.Enabled = False
      Inside_Single_HD.Enabled = False
      Inside_Double_HD.Enabled = False
  
      'Outside Picket
      Outside_None.Enabled = True
      Outside_Single_SD.Enabled = True
      Outside_Double_SD.Enabled = True
      Outside_Single_HD.Enabled = True
      Outside_Double_HD.Enabled = True

      Loops = 0
      Run_Setup
  End If
End Sub

Private Sub Dim_B_Calc()

  Dim BYMin, BGMin, BGMax, BYMax
  
  Select Case Belt_Category
    Case "Heavy Duty Flatwire"
      Select Case BeltWidth
        Case Is <= 7
          BYMin = 0: BGMin = 0: BGMax = 0.375: BYMax = 0.5
        Case Is <= 72
          BYMin = 0: BGMin = 0: BGMax = 0.5: BYMax = 1
        Case Is > 72
          BYMin = 0: BGMin = 0: BGMax = 0.75: BYMax = 1.5
      End Select
    Case "Omniflex", "Standard Weight Flatwire Welded", "Standard Weight Flatwire Clinched"
      Select Case BeltWidth
        Case Is <= 6
          BYMin = 0: BGMin = 0: BGMax = 0.25: BYMax = 0.25
        Case Is <= 72
          BYMin = 0: BGMin = 0: BGMax = 0.5: BYMax = 0.5
        Case Is > 72
          BYMin = 0: BGMin = 0: BGMax = 0.75: BYMax = 0.75
      End Select
  End Select
  
  With CalcSheet
    .Range("N88") = BYMin
    .Range("O88") = BGMin
    .Range("P88") = BGMax
    .Range("Q88") = BYMax
  End With
  
End Sub

Private Sub Free_Picket_Width_Calc()

    Dim Inside_Links As String
    Dim Outside_Links As String
    Dim Bar_Link_Adder As Variant: Bar_Link_Adder = 0
    Dim PartNumber As String: PartNumber = UCase(CalcSheet.Range("LPartNum"))
    Dim Sd As Integer
    Dim Hd As String
    Dim BarLinks As String
    Dim MinWidth As Double
    Dim MaxWidth As Double
    Dim Center_Link_Loc As Double: Center_Link_Loc = CDbl(CalcSheet.Range("Center_Link_Location"))
    Dim TempWidth As Double
    Dim FreePicketWidthMin As Double
    Dim FreePicketWidthMax As Double
    Dim TargetWidth As Double
    Dim BeltWidth As Double: BeltWidth = CDbl(CalcSheet.Range("BeltWidth"))
    
    'Determine Bar Links
    Inside_Links = Radio(Inside_None, Inside_Single_SD, Inside_Double_SD, Inside_Single_HD, Inside_Double_HD, "None,Single Standard Duty,Double Standard Duty,Single Heavy Duty,Double Heavy Duty")
    Outside_Links = Radio(Outside_None, Outside_Single_SD, Outside_Double_SD, Outside_Single_HD, Outside_Double_HD, "None,Single Standard Duty,Double Standard Duty,Single Heavy Duty,Double Heavy Duty")
    
    If DiffSpiralCount > 1 And Outside_Picket Then
      Temp = Outside_Links
    Else
      Temp = Inside_Links
    End If

    'Return Values to Spreadsheet
    If PartNumber = "SROFG1" Or PartNumber = "SROFG3" Then
      If Inside_Picket = "True" Then
        Select Case Inside_Links
          Case "None"
            Sd = 0
            Hd = 0
          Case "Single Standard Duty"
            Sd = 1
            Hd = 0
          Case "Double Standard Duty"
            Sd = 2
            Hd = 0
          Case "Single Heavy Duty"
            Sd = 0
            Hd = 1
          Case "Double Heavy Duty"
            Sd = 0
            Hd = 2
          Case Else
            Exit Sub
        End Select

        Hd = Hd + 1

        If Sd > 0 Then
          BarLinks = Sd & "SD"
  
          If Hd > 0 Then
            BarLinks = BarLinks & Hd & "HD"
          End If
        ElseIf Hd > 0 Then
          BarLinks = Hd & "HD"
        Else
          BarLinks = 0
        End If
    
        TempWidth = Center_Link_Loc
        
      Else 'Outside Picket
    
        Sd = 0
        Hd = 1

        Select Case Outside_Links
          Case "None"
          Case "Single Standard Duty"
            Sd = Sd + 1
          Case "Double Standard Duty"
            Sd = Sd + 2
          Case "Single Heavy Duty"
            Hd = Hd + 1
          Case "Double Heavy Duty"
            Hd = Hd + 2
          Case Else
            Exit Sub
        End Select

        If Sd > 0 Then
          BarLinks = Sd & "SD"
  
          If Hd > 0 Then
            BarLinks = BarLinks & Hd & "HD"
          End If
        ElseIf Hd > 0 Then
          BarLinks = Hd & "HD"
        Else
          BarLinks = 0
        End If
    
        TempWidth = BeltWidth - Center_Link_Loc
      End If
    Else
      Select Case Inside_Links
        Case "None"
          Sd = 0
          Hd = 0
        Case "Single Standard Duty"
          Sd = 1
          Hd = 0
        Case "Double Standard Duty"
          Sd = 2
          Hd = 0
        Case "Single Heavy Duty"
          Sd = 0
          Hd = 1
        Case "Double Heavy Duty"
          Sd = 0
          Hd = 2
        Case Else
          Exit Sub
      End Select
  
      Select Case Outside_Links
        Case "None"
        Case "Single Standard Duty"
          Sd = Sd + 1
        Case "Double Standard Duty"
          Sd = Sd + 2
        Case "Single Heavy Duty"
          Hd = Hd + 1
        Case "Double Heavy Duty"
          Hd = Hd + 2
        Case Else
          Exit Sub
      End Select
  
      If Sd > 0 Then
        BarLinks = Sd & "SD"
    
        If Hd > 0 Then
          BarLinks = BarLinks & Hd & "HD"
        End If
      ElseIf Hd > 0 Then
        BarLinks = Hd & "HD"
      Else
        BarLinks = 0
      End If
  
      TempWidth = BeltWidth
    End If

  If IsNumeric(TempWidth) Then
    MinWidth = FwFreePicketWidth(PartNumber, BeltWidth, TempWidth, BarLinks, True)
    MaxWidth = FwFreePicketWidth(PartNumber, BeltWidth, TempWidth, BarLinks, False)
    FreePicketWidthMin = RoundToFraction(TempWidth + MinWidth, 32, False, True)
    FreePicketWidthMax = RoundToFraction(TempWidth + MaxWidth, 32, False, True)
    TargetWidth = (FreePicketWidthMin + FreePicketWidthMax) / 2
    
    CalcSheet.Range("L90") = TargetWidth
    CalcSheet.Range("N90") = FreePicketWidthMin - TargetWidth
    CalcSheet.Range("O90") = FreePicketWidthMin - TargetWidth
    CalcSheet.Range("P90") = FreePicketWidthMax - TargetWidth
    CalcSheet.Range("Q90") = FreePicketWidthMax - TargetWidth
  End If
End Sub

Private Sub Picket_Compression_Calc()

  Dim CPW_YMin As Variant
  Dim CPW_GMin As Variant
  Dim CPW_Target As Variant
  Dim CPW_GMax As Variant
  Dim CPW_YMax As Variant

    Select Case Belt_Category
        Case "Standard Weight Flatwire Welded"
            Select Case Loops
                Case Is < 13
                  CPW_YMin = -0.0625: CPW_GMin = -0.0625: CPW_Target = 0.5: CPW_GMax = 0.0625: CPW_YMax = 0.0625
                Case Is < 19
                  CPW_YMin = -0.09375: CPW_GMin = -0.09375: CPW_Target = 0.65625: CPW_GMax = 0.09375: CPW_YMax = 0.09375
                Case Is < 25
                  CPW_YMin = -0.09375: CPW_GMin = -0.09375: CPW_Target = 0.71875: CPW_GMax = 0.09375: CPW_YMax = 0.09375
                Case Is < 31
                  CPW_YMin = -0.125: CPW_GMin = -0.125: CPW_Target = 0.8125: CPW_GMax = 0.125: CPW_YMax = 0.125
                Case Is < 37
                  CPW_YMin = -0.125: CPW_GMin = -0.125: CPW_Target = 0.875: CPW_GMax = 0.125: CPW_YMax = 0.125
                Case Is < 43
                  CPW_YMin = -0.15625: CPW_GMin = -0.15625: CPW_Target = 1.03125: CPW_GMax = 0.15625: CPW_YMax = 0.15625
                Case Is < 49
                  CPW_YMin = -0.1875: CPW_GMin = -0.1875: CPW_Target = 1.125: CPW_GMax = 0.1875: CPW_YMax = 0.1875
                Case Is < 55
                  CPW_YMin = -0.21875: CPW_GMin = -0.21875: CPW_Target = 1.28125: CPW_GMax = 0.21875: CPW_YMax = 0.21875
                Case Is < 61
                  CPW_YMin = -0.25: CPW_GMin = -0.25: CPW_Target = 1.375: CPW_GMax = 0.25: CPW_YMax = 0.25
                Case Is < 67
                  CPW_YMin = -0.25: CPW_GMin = -0.25: CPW_Target = 1.4375: CPW_GMax = 0.25: CPW_YMax = 0.25
                Case Is < 73
                  CPW_YMin = -0.28125: CPW_GMin = -0.28125: CPW_Target = 1.59375: CPW_GMax = 0.28125: CPW_YMax = 0.28125
                Case Is < 79
                  CPW_YMin = -0.3125: CPW_GMin = -0.3125: CPW_Target = 1.6875: CPW_GMax = 0.3125: CPW_YMax = 0.3125
                Case Is < 85
                  CPW_YMin = -0.34375: CPW_GMin = -0.34375: CPW_Target = 1.84375: CPW_GMax = 0.34375: CPW_YMax = 0.34375
                Case Is < 91
                  CPW_YMin = -0.34375: CPW_GMin = -0.34375: CPW_Target = 1.90625: CPW_GMax = 0.34375: CPW_YMax = 0.34375
                Case Is < 97
                  CPW_YMin = -0.375: CPW_GMin = -0.375: CPW_Target = 2: CPW_GMax = 0.375: CPW_YMax = 0.375
                Case Is < 103
                  CPW_YMin = -0.40625: CPW_GMin = -0.40625: CPW_Target = 2.15625: CPW_GMax = 0.40625: CPW_YMax = 0.40625
                Case Is < 109
                  CPW_YMin = -0.721875: CPW_GMin = -0.721875: CPW_Target = 1.903125: CPW_GMax = 0.721875: CPW_YMax = 0.721875
                Case Is < 115
                  CPW_YMin = -0.4375: CPW_GMin = -0.4375: CPW_Target = 2.1875: CPW_GMax = 0.4375: CPW_YMax = 0.4375
                Case Is < 121
                  CPW_YMin = -0.46875: CPW_GMin = -0.46875: CPW_Target = 2.46875: CPW_GMax = 0.46875: CPW_YMax = 0.46875
                Case Is < 127
                  CPW_YMin = -0.5: CPW_GMin = -0.5: CPW_Target = 2.5: CPW_GMax = 0.5: CPW_YMax = 0.5
                Case Is < 133
                  CPW_YMin = -0.5: CPW_GMin = -0.5: CPW_Target = 2.625: CPW_GMax = 0.5: CPW_YMax = 0.5
                Case Is < 139
                  CPW_YMin = -0.53125: CPW_GMin = -0.53125: CPW_Target = 2.78125: CPW_GMax = 0.53125: CPW_YMax = 0.53125
                Case Is < 145
                  CPW_YMin = -0.5625: CPW_GMin = -0.5625: CPW_Target = 2.875: CPW_GMax = 0.5625: CPW_YMax = 0.5625
                Case Is < 151
                  CPW_YMin = -0.59375: CPW_GMin = -0.59375: CPW_Target = 3.03125: CPW_GMax = 0.59375: CPW_YMax = 0.59375
                Case Is < 157
                  CPW_YMin = -0.625: CPW_GMin = -0.625: CPW_Target = 3.125: CPW_GMax = 0.625: CPW_YMax = 0.625
            End Select
        Case "Standard Weight Flatwire Clinched"
            Select Case Loops
                Case Is < 13
                  CPW_YMin = -0.1875: CPW_GMin = -0.1875: CPW_Target = 0.8125: CPW_GMax = 0.1875: CPW_YMax = 0.1875
                Case Is < 19
                  CPW_YMin = -0.21875: CPW_GMin = -0.21875: CPW_Target = 0.84375: CPW_GMax = 0.21875: CPW_YMax = 0.21875
                Case Is < 25
                  CPW_YMin = -0.21875: CPW_GMin = -0.21875: CPW_Target = 0.96875: CPW_GMax = 0.21875: CPW_YMax = 0.21875
                Case Is < 31
                  CPW_YMin = -0.25: CPW_GMin = -0.25: CPW_Target = 1.0625: CPW_GMax = 0.25: CPW_YMax = 0.25
                Case Is < 37
                  CPW_YMin = -0.25: CPW_GMin = -0.25: CPW_Target = 1.125: CPW_GMax = 0.25: CPW_YMax = 0.25
                Case Is < 43
                  CPW_YMin = -0.28125: CPW_GMin = -0.28125: CPW_Target = 1.21875: CPW_GMax = 0.28125: CPW_YMax = 0.28125
                Case Is < 49
                  CPW_YMin = -0.3125: CPW_GMin = -0.3125: CPW_Target = 1.375: CPW_GMax = 0.3125: CPW_YMax = 0.3125
                Case Is < 55
                  CPW_YMin = -0.40625: CPW_GMin = -0.40625: CPW_Target = 1.59375: CPW_GMax = 0.40625: CPW_YMax = 0.40625
                Case Is < 61
                  CPW_YMin = -0.375: CPW_GMin = -0.375: CPW_Target = 1.625: CPW_GMax = 0.375: CPW_YMax = 0.375
                Case Is < 67
                  CPW_YMin = -0.375: CPW_GMin = -0.375: CPW_Target = 1.6875: CPW_GMax = 0.375: CPW_YMax = 0.375
                Case Is < 73
                  CPW_YMin = -0.375: CPW_GMin = -0.375: CPW_Target = 1.8125: CPW_GMax = 0.375: CPW_YMax = 0.375
                Case Is < 79
                  CPW_YMin = -0.4375: CPW_GMin = -0.4375: CPW_Target = 1.9375: CPW_GMax = 0.4375: CPW_YMax = 0.4375
                Case Is < 85
                  CPW_YMin = -0.46875: CPW_GMin = -0.46875: CPW_Target = 2.09375: CPW_GMax = 0.46875: CPW_YMax = 0.46875
                Case Is < 91
                    CPW_YMin = -0.46875: CPW_GMin = -0.46875: CPW_Target = 2.15625: CPW_GMax = 0.46875: CPW_YMax = 0.46875
                Case Is < 97
                    CPW_YMin = -0.5: CPW_GMin = -0.5: CPW_Target = 2.25: CPW_GMax = 0.5: CPW_YMax = 0.5
                Case Is < 103
                    CPW_YMin = -0.53125: CPW_GMin = -0.53125: CPW_Target = 2.40625: CPW_GMax = 0.53125: CPW_YMax = 0.53125
                Case Is < 109
                    CPW_YMin = -0.53125: CPW_GMin = -0.53125: CPW_Target = 2.46875: CPW_GMax = 0.53125: CPW_YMax = 0.53125
                Case Is < 115
                    CPW_YMin = -0.5625: CPW_GMin = -0.5625: CPW_Target = 2.5625: CPW_GMax = 0.5625: CPW_YMax = 0.5625
                Case Is < 121
                    CPW_YMin = -0.59375: CPW_GMin = -0.59375: CPW_Target = 2.71875: CPW_GMax = 0.59375: CPW_YMax = 0.59375
                Case Is < 127
                    CPW_YMin = -0.65625: CPW_GMin = -0.65625: CPW_Target = 2.84375: CPW_GMax = 0.65625: CPW_YMax = 0.65625
                Case Is < 133
                    CPW_YMin = -0.625: CPW_GMin = -0.625: CPW_Target = 2.875: CPW_GMax = 0.625: CPW_YMax = 0.625
                Case Is < 139
                    CPW_YMin = -0.65625: CPW_GMin = -0.65625: CPW_Target = 3.03125: CPW_GMax = 0.65625: CPW_YMax = 0.65625
                Case Is < 145
                    CPW_YMin = -0.6875: CPW_GMin = -0.6875: CPW_Target = 3.125: CPW_GMax = 0.6875: CPW_YMax = 0.6875
                Case Is < 151
                    CPW_YMin = -0.71875: CPW_GMin = -0.71875: CPW_Target = 3.28125: CPW_GMax = 0.71875: CPW_YMax = 0.71875
                Case Is < 157
                    CPW_YMin = -0.75: CPW_GMin = -0.75: CPW_Target = 3.375: CPW_GMax = 0.75: CPW_YMax = 0.75
            End Select
        Case "Heavy Duty Flatwire"
            Select Case Loops
                Case Is < 19
                  CPW_YMin = -0.09375: CPW_GMin = -0.09375: CPW_Target = 0.65625: CPW_GMax = 0.09375: CPW_YMax = 0.09375
                Case Is < 25
                  CPW_YMin = -0.09375: CPW_GMin = -0.09375: CPW_Target = 0.71875: CPW_GMax = 0.09375: CPW_YMax = 0.09375
                Case Is < 31
                  CPW_YMin = -0.125: CPW_GMin = -0.125: CPW_Target = 0.8125: CPW_GMax = 0.125: CPW_YMax = 0.125
                Case Is < 37
                  CPW_YMin = -0.125: CPW_GMin = -0.125: CPW_Target = 0.875: CPW_GMax = 0.125: CPW_YMax = 0.125
                Case Is < 43
                  CPW_YMin = -0.15625: CPW_GMin = -0.15625: CPW_Target = 0.96875: CPW_GMax = 0.15625: CPW_YMax = 0.15625
                Case Is < 49
                  CPW_YMin = -0.1875: CPW_GMin = -0.1875: CPW_Target = 1.0625: CPW_GMax = 0.1875: CPW_YMax = 0.1875
                Case Is < 55
                  CPW_YMin = -0.125: CPW_GMin = -0.125: CPW_Target = 1.0625: CPW_GMax = 0.125: CPW_YMax = 0.125
                Case Is < 61
                  CPW_YMin = -0.25: CPW_GMin = -0.25: CPW_Target = 1.25: CPW_GMax = 0.25: CPW_YMax = 0.25
                Case Is < 67
                  CPW_YMin = -0.25: CPW_GMin = -0.25: CPW_Target = 1.3125: CPW_GMax = 0.25: CPW_YMax = 0.25
                Case Is < 73
                  CPW_YMin = -0.28125: CPW_GMin = -0.28125: CPW_Target = 1.40625: CPW_GMax = 0.28125: CPW_YMax = 0.28125
                Case Is < 79
                  CPW_YMin = -0.3125: CPW_GMin = -0.3125: CPW_Target = 1.5: CPW_GMax = 0.3125: CPW_YMax = 0.3125
                Case Is < 85
                  CPW_YMin = -0.34375: CPW_GMin = -0.34375: CPW_Target = 1.59375: CPW_GMax = 0.34375: CPW_YMax = 0.34375
                Case Is < 91
                  CPW_YMin = -0.34375: CPW_GMin = -0.34375: CPW_Target = 1.65625: CPW_GMax = 0.34375: CPW_YMax = 0.34375
                Case Is < 97
                  CPW_YMin = -0.375: CPW_GMin = -0.375: CPW_Target = 1.75: CPW_GMax = 0.375: CPW_YMax = 0.375
                Case Is < 103
                  CPW_YMin = -0.40625: CPW_GMin = -0.40625: CPW_Target = 1.84375: CPW_GMax = 0.40625: CPW_YMax = 0.40625
                Case Is < 109
                  CPW_YMin = -0.40625: CPW_GMin = -0.40625: CPW_Target = 1.90625: CPW_GMax = 0.40625: CPW_YMax = 0.40625
                Case Is < 115
                  CPW_YMin = -0.4375: CPW_GMin = -0.4375: CPW_Target = 2: CPW_GMax = 0.4375: CPW_YMax = 0.4375
                Case Is < 121
                  CPW_YMin = -0.46875: CPW_GMin = -0.46875: CPW_Target = 2.09375: CPW_GMax = 0.46875: CPW_YMax = 0.46875
                Case Is < 127
                  CPW_YMin = -0.46875: CPW_GMin = -0.46875: CPW_Target = 2.15625: CPW_GMax = 0.46875: CPW_YMax = 0.46875
                Case Is < 133
                  CPW_YMin = -0.5: CPW_GMin = -0.5: CPW_Target = 2.25: CPW_GMax = 0.5: CPW_YMax = 0.5
                Case Is < 139
                  CPW_YMin = -0.53125: CPW_GMin = -0.53125: CPW_Target = 2.34375: CPW_GMax = 0.53125: CPW_YMax = 0.53125
                Case Is < 145
                  CPW_YMin = -0.5625: CPW_GMin = -0.5625: CPW_Target = 2.4375: CPW_GMax = 0.5625: CPW_YMax = 0.5625
                Case Is < 151
                  CPW_YMin = -0.59375: CPW_GMin = -0.59375: CPW_Target = 2.53125: CPW_GMax = 0.59375: CPW_YMax = 0.59375
                Case Is > 157
                  CPW_YMin = -0.625: CPW_GMin = -0.625: CPW_Target = 2.625: CPW_GMax = 0.625: CPW_YMax = 0.625
            End Select
        Case "Omniflex"
          CPW_YMin = 0: CPW_GMin = 0: CPW_Target = 0: CPW_GMax = 0: CPW_YMax = 0
    End Select

  With CalcSheet
    .Range("L91") = CPW_Target
    .Range("N91") = CPW_YMin
    .Range("O91") = CPW_GMin
    .Range("P91") = CPW_GMax
    .Range("Q91") = CPW_YMax
  End With
  
End Sub

Private Sub Gleen_Info()

  Dim Job_Comments As String: Job_Comments = CalcSheet.Range("JobComments")
  Dim Op_Comment As String: Op_Comment = CalcSheet.Range("Operation_Comment")
  Dim Temp As Variant
  
  If IsBad(BeltWidth) Then
    BeltWidth = Comment_Search("Width", Job_Comments, "Inches", "in.", "", "")
    If IsNumeric(BeltWidth) And BeltWidth > 0 Then
      CalcSheet.Range("BeltWidth") = BeltWidth
    Else
      BeltWidth = TryParseFraction(ShowCommentBox("Belt Width", Job_Comments))
      If IsNumeric(BeltWidth) = True And BeltWidth > 0 Then
        CalcSheet.Range("BeltWidth") = BeltWidth
      End If
    End If
  End If
  
  If DiffSpiralCount > 1 And IsBad(Center_Link_Location) Then
    Center_Link_Location = Comment_Search("Center Link", Job_Comments, "inches", "in.", "in", "")
    If IsNumeric(Center_Link_Location) And Center_Link_Location > 0 Then
      CalcSheet.Range("Center_Link_Location") = Center_Link_Location
    Else
      Center_Link_Location = TryParseFraction(ShowCommentBox("Center Link Location", Op_Comment))
      CalcSheet.Range("Center_Link_Location") = Center_Link_Location
    End If
  End If
  
  If IsBad(Loops) Then
    If IsNumeric(Comment_Search("Loops", Op_Comment, "", "", "", "")) = True Then
      Loops = Comment_Search("Loops", Op_Comment, "", "", "", "")
      CalcSheet.Range("Loop_Count") = Loops
    Else
      Temp = TryParseFraction(ShowCommentBox("Number of Loops", Op_Comment))
      If IsNumeric(Temp) Then
        Loops = Temp
        CalcSheet.Range("Loop_Count") = Loops
      Else
        Loops = 0
      End If
    End If
  End If
End Sub

Private Sub Validator()

  While IsBad(BeltWidth) Or IsBad(Loops) Or (IsBad(Center_Link_Location) And DiffSpiralCount > 1)
    Gleen_Info
  Wend

  Free_Picket_Width_Calc
  Picket_Compression_Calc
  Dim_B_Calc
End Sub

Private Sub Image1_Click()
  Dim Inside_Links As String, Outside_Links As String, IO_Picket As String

    Inside_Links = Radio(Inside_None, Inside_Single_SD, Inside_Double_SD, Inside_Single_HD, Inside_Double_HD, "None,Single Standard Duty,Double Standard Duty,Single Heavy Duty,Double Heavy Duty")
    Outside_Links = Radio(Outside_None, Outside_Single_SD, Outside_Double_SD, Outside_Single_HD, Outside_Double_HD, "None,Single Standard Duty,Double Standard Duty,Single Heavy Duty,Double Heavy Duty")
    IO_Picket = Radio(Inside_Picket, Outside_Picket, "Inside Picket", "Outside Picket")
  
  If Inside_Links > "" And Outside_Links > "" And IO_Picket > "" Then
    Validator
    Flatwire_Picket_Setup.Show
  Else
    MsgBox "The Bar Links and Inside / Outside Picket Must Be Completed to Continue."
  End If
End Sub

Private Sub Submit_Click()
  On Error GoTo GeneralError
  
  Dim InspName As String: InspName = "Flatwire_Picket_Inspection"
  
  'Setup and Verify Parameters
  ClearCalcSheet True
  Validator
      
  CalcSheet.Range("Insp_Plan") = CalcSheet.Range(Inspection_Name & "_Plan")
  CalcSheet.Range("Spec_ID") = CalcSheet.Range(Inspection_Name & "_Spec")
  CalcSheet.Range("Schar3") = Machine_No
  CalcSheet.Range("Schar7") = Radio(Inside_None, Inside_Single_SD, Inside_Double_SD, Inside_Single_HD, Inside_Double_HD, "None,Single Standard Duty,Double Standard Duty,Single Heavy Duty,Double Heavy Duty")
  CalcSheet.Range("Schar8") = Radio(Outside_None, Outside_Single_SD, Outside_Double_SD, Outside_Single_HD, Outside_Double_HD, "None,Single Standard Duty,Double Standard Duty,Single Heavy Duty,Double Heavy Duty")
  CalcSheet.Range("Data1") = TryParseFraction(A)
  CalcSheet.Range("Data2") = TryParseFraction(B)
  CalcSheet.Range("Data3") = TryParseFraction(C)
  CalcSheet.Range("Data4") = TryParseFraction(Free_Picket_Width)
  CalcSheet.Range("Data5") = TryParseFraction(Compresed_Picket_Width)
  CalcSheet.Range("Check2") = Radio(IRP_Pass, IRP_Fail)
  CalcSheet.Range("Check3") = Radio(COA_Pass, COA_Fail)
  CalcSheet.Range("Check4") = Radio(COLS_Pass, COLS_Fail)
  CalcSheet.Range("Check5") = Radio(MTA_Pass, MTA_Fail)
  CalcSheet.Range("Check6") = Radio(Crooked_Loop)
  CalcSheet.Range("Check7") = Radio(Scratches)
  CalcSheet.Range("Check8") = Radio(High_Corner)
  CalcSheet.Range("Check9") = Radio(Tool_Marks)
  CalcSheet.Range("Check10") = Radio(Burrs)
  CalcSheet.Range("IO_Spiral") = Radio(Inside_Picket, Outside_Picket, "Inside Picket", "Outside Picket")

  If IsEmpty(CalcSheet.Range("Schar3")) Or _
    IsEmpty(CalcSheet.Range("Schar7")) Or _
    IsEmpty(CalcSheet.Range("Schar8")) Or _
    IsEmpty(CalcSheet.Range("Data1")) Or _
    IsEmpty(CalcSheet.Range("Data2")) Or _
    IsEmpty(CalcSheet.Range("Data3")) Or _
    IsEmpty(CalcSheet.Range("Data4")) Or _
    (IsEmpty(CalcSheet.Range("Data5")) And Belt_Category <> "Omniflex") Or _
    IsEmpty(CalcSheet.Range("Check2")) Or _
    IsEmpty(CalcSheet.Range("Check3")) Or _
    IsEmpty(CalcSheet.Range("Check4")) Or _
    IsEmpty(CalcSheet.Range("Check5")) Or _
    IsEmpty(CalcSheet.Range("IO_Spiral")) Then
    GoTo BlankForm
  End If
  
  'Calculate Results
  If CalcSheet.Range(InspName & "_Comment") = Empty Then
    CalcSheet.Range("Passed") = 1
    CalcSheet.Range("Value") = ""
    CalcSheet.Range("Failed_Comment") = ""
  Else
    Temp = Replace(CalcSheet.Range(InspName & "_Comment"), "?", ".  ")
    CalcSheet.Range("Passed") = 0
    CalcSheet.Range("Value") = "Picket Rejected"
    CalcSheet.Range("Failed_Comment") = Temp
    RejectForm Replace(CalcSheet.Range(InspName & "_Comment"), "?", vbNewLine)
  End If

  'Send Data To SQL
  WriteToSQL "ice.UD10"
  
  A = ""
  B = ""
  C = ""
  Free_Picket_Width = ""
  Compresed_Picket_Width = ""
  IRP_Pass = False
  IRP_Fail = False
  COA_Pass = False
  COA_Fail = False
  COLS_Pass = False
  COLS_Fail = False
  MTA_Pass = False
  MTA_Fail = False
  Crooked_Loop = False
  Scratches = False
  High_Corner = False
  Tool_Marks = False
  Burrs = False
  A.SetFocus
  Inspection_Num = "Inspection Num: " + CStr(CalcSheet.Range("SampleNum") + 1)
Exit Sub

GeneralError:
  MsgBox "Errors have been Detected in the Data Entered Please Correct and Resubmit"
Exit Sub

BlankForm:
  MsgBox "Please Completely Fill out the Form and Resubmit."
End Sub

Private Sub UserForm_Activate()

  ClearCalcSheet False
  Data_Dump.Visible = False
  
  Dim JobOper As New ADODB.Recordset

  DBEpicor.Open
  With JobOper
    Set .ActiveConnection = DBEpicor
    .Open "SELECT * FROM " & "erp.JobOper" & " WHERE Company = " & Company & " AND JobNum = '" & JobNum & "' AND OpCode = 'FWMUL01'"
    If .EOF Or .BOF Then MsgBox "Error Returning the Operation Comments"
  End With
  
  CalcSheet.Range("Operation_Comment") = Compact(JobOper.Fields("CommentText"))
  DBEpicor.Close
  
  Select Case BeltType
    Case "FWA1", "FWA3", "FWB1", "FWB3"
      Belt_Category = "Standard Weight Flatwire Clinched"
      Flatwire_Picket_Inspection.Inside_None = True
      Flatwire_Picket_Inspection.Outside_None = True
    Case "FWA2", "FWA4", "FWA5", "FWA5S", "FWA6", "FWB2", "FWB4", "FWB5", "FWB6"
      Belt_Category = "Standard Weight Flatwire Welded"
      Flatwire_Picket_Inspection.Inside_None = True
      Flatwire_Picket_Inspection.Outside_None = True
    Case "FWC1", "FWC2", "FWC6", "FWC6SB", "FWH3"
      Belt_Category = "Heavy Duty Flatwire"
    Case "OFE1", "OFE2", "OFE3", "SROFG1", "SROFG3"
      Belt_Category = "Omniflex"
  End Select
  
  Inspection_Num = "Inspection Num: " + CStr(SampleNum)
  Machine_No = ""
  
  If DiffSpiralCount > 1 Then
    Inside_Picket.Enabled = True
    Outside_Picket.Enabled = True
  Else
    Inside_Picket.Enabled = False
    Outside_Picket.Enabled = False
    Inside_Picket = True
  End If
  
  If Belt_Category = "Omniflex" Then
    Compresed_Picket_Width.Enabled = False
    Lbl_CPW.Enabled = False
  Else
    Compresed_Picket_Width.Enabled = True
    Lbl_CPW.Enabled = True
  End If

  Inside_Picket = True
End Sub

