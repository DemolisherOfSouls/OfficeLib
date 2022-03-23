Attribute VB_Name = "Radio"
Function Radio(T1, Answer)

  Dim Working_Input As String: Working_Input = T1
  Dim Working_Answer As String: Working_Answer = Answer
  Dim Temp_Input As String
  Dim Temp_Answer As String
  
  If InStr(1, Working_Input, ",", vbTextCompare) > 0 Then
    Temp_Input = Left(Working_Input, InStr(1, Working_Input, ",", vbTextCompare) - 1)
  Else
      MsgBox ("Only One Input")
      Radio = ""
      Exit Function
  End If
  
  If InStr(1, Working_Answer, ",", vbTextCompare) > 0 Then
      Temp_Answer = Left(Working_Answer, InStr(1, Working_Answer, ",", vbTextCompare) - 1)
  Else
      MsgBox ("Only One Answer")
      Radio = ""
      Exit Function
  End If
  
  While Temp_Input <> True
      If InStr(1, Working_Input, ",", vbTextCompare) > 0 Then
          Temp_Input = Left(Working_Input, InStr(1, Working_Input, ",", vbTextCompare) - 1)
          Temp_Answer = Left(Working_Answer, InStr(1, Working_Answer, ",", vbTextCompare) - 1)
      Else
          Temp_Input = Working_Input
          Temp_Answer = Working_Answer
      End If
      
      If Temp_Input <> True Then
          If Len(Working_Input) - Len(Temp_Input) > 1 Then
              Working_Input = Right(Working_Input, Len(Working_Input) - Len(Temp_Input) - 1)
          Else
              Radio = ""
              Exit Function
          End If
      
          If Len(Working_Answer) - Len(Temp_Answer) > 1 Then
              Working_Answer = Right(Working_Answer, Len(Working_Answer) - Len(Temp_Answer) - 1)
          Else
              Radio = ""
              Exit Function
          End If
      End If
  Wend
      
  If InStr(1, Temp_Answer, ",") > 0 Then
      Radio = Left(Temp_Answer, InStr(1, Temp_Answer, ",") - 1)
  Else
      Radio = Temp_Answer
  End If
    
End Function
