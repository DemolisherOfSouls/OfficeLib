Attribute VB_Name = "Common"
Option Explicit
Option Base 1
Option Compare Text

Public Type BarLinkCounts
  Sd As Byte
  Hd As Byte
End Type

Public Sub ProgramInit()
  DBEpicor.ConnectionString = CSEpicor
  DBEng.ConnectionString = CSEngineer
  If Company = 0 Then Company = 200
End Sub

Public Function IsText(ByVal Value)
  IsText = Not IsNumeric(Value) And Not IsEmpty(Value) And Not IsDate(Value)
End Function

Public Function IsBad(ByRef Var As Variant, Optional ByVal CheckNum As Boolean = True) As Boolean
  If (Not IsNumeric(Var) And CheckNum) Or Var = 0 Or Var = "" Then
    IsBad = True
    Var = ""
  Else
    IsBad = False
  End If
End Function

Public Function Radio(ByVal Option1 As Boolean, Optional ByVal Option2 As Boolean, Optional ByVal Result1 As Variant = 1, Optional ByVal Result2 As Variant = 0) As Variant
  If Option1 Then
    Radio = Result1
  ElseIf Option2 Then
    Radio = Result2
  Else
    Radio = Empty
  End If
End Function

Public Sub ClearCalcSheet(Optional ByVal Written As Boolean = False)
  With CalcSheet
    If Written Then
      .Range("C21:C78") = ""
    Else
      .Range("BeltWidth") = ""
      .Range("Center_Link_Location") = ""
      .Range("Operation_Comment") = ""
      .Range("Spiral_Size") = ""
      .Range("Loop_Count") = ""
      .Range("CrimpDepth") = ""
      .Range("Fabric_Width") = ""
      .Range("Free_Picket_Width") = ""
    End If
  End With
End Sub

Public Function ShowCommentBox(ByVal Error As String, ByVal Comments As String) As Variant
  With CommentBoxForm
    .Error_Comment = Error
    .Comment_Box = Comments
    .Show True
  End With
  ShowCommentBox = CommentBoxReturn
End Function

Public Sub ShowRejectionForm(ByVal Reason As String)
  With RejectForm
    .Reject_Desc = Reason
    .Show True
  End With
End Sub

Public Function Atize(ByVal Num As Integer)
  Dim ITemp As Integer: ITemp = 0
  
  While Num > 26
    ITemp = ITemp + 1
    Num = Num - 26
  Wend
  
  If ITemp = 0 Then
    Atize = Chr(Num + 64)
  Else
    Atize = (Chr(ITemp + 64) & Chr(Num + 64))
  End If
End Function

Public Function RoundToFraction(Number As Double, Denom As Integer, ByVal RoundUp As Boolean) As Double

  Const Whole As Integer = Int(Number)
  Dim Fract As Integer: Fract = Int((Number - Whole) * Denom * 2)

  If RoundUp Or Fract Mod 2 = 1 Then Fract = Fract + 1
  RoundToFraction = CDbl(Whole) + CDbl(Fract) / CDbl(Denom)
    
End Function

Function Required_Field(In1, In2, In3)
  If In1 > "" Then
    Required_Field = In1
  ElseIf In2 > "" Then
    Required_Field = In2
  ElseIf In3 > "" Then
    Required_Field = In3
  Else
    MsgBox "Please Completely Fill out the Form and Resubmit."
    Required_Field = "Exit Sub"
  End If
End Function

'Setters for Global Variables and CalcSheet

Public Sub SetBeltType(ByVal PN As String)
  If InStr(PN, "STK") > 0 Then
    BeltTypeScreen.Show True
  Else
    CalcSheet.Range("LPartNum") = PN
    BeltType = PN
  End If
End Sub

Public Sub SetOperation(Typ, Run, Setup)

  Inspection = Typ
  Operation = Radio(Run, Setup, "Run", "Setup")
  
  CalcSheet.Range("Insp_Type") = Inspection & " " & Operation
  CalcSheet.Range("Setup_Run") = Operation

End Sub

Public Sub SetNextSampleNum()

  Dim UD10 As New ADODB.Recordset
  
  DBEpicor.Open
  Set UD10.ActiveConnection = DBEpicor
  UD10.Open "SELECT * FROM " & "ice.UD10" & " WHERE Company = " & Company & " AND Key2 = '" & Inspection & " " & Operation & "' AND Key1 = '" & JobNum & "'"
  
  If UD10.EOF Then
    SampleNum = 0
  Else
    UD10.MoveFirst

    While Not UD10.EOF
      If SampleNum < UD10.Fields("Key3") Then
        SampleNum = UD10.Fields("Key3")
      End If
      UD10.MoveNext
    Wend
  End If
  
  DBEpicor.Close
  
  SampleNum = SampleNum + 1

  CalcSheet.Range("SampleNum") = SampleNum
  
End Sub

Public Function Max(ByVal X, ByVal y)
  Max = IIf(X > y, X, y)
End Function

