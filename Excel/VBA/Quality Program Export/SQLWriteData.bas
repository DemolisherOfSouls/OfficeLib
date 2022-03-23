Attribute VB_Name = "SQLWriteData"
Option Explicit
Option Compare Text
Option Base 1

Private Const WriteRowStart = 6
Private Const WriteRowEnd = 77
Private Const WriteCol = "C"
Private Const FieldCol = "A"

Public Sub WriteToSQL(ByVal Table As String)

  Dim Col As String: Col = ""
  Dim Val As String: Val = ""
  Dim Row As Integer

  SetNextSampleNum

  DBEpicor.Open

  For Row = WriteRowStart To WriteRowEnd
  
    With CalcSheet
      Const V As String = .Range(WriteCol & Row)
      Const C As String = .Range(FieldCol & i)
      Const S As String = IIf(Row = WriteRowStart, "", ",")
    End With
    
    If V <> Empty Then
      Val = Val & S & "'" & V & "'"
      Col = Col & S & C
    End If
    
  Next

  DBEpicor.Execute ("INSERT INTO " & Table & " (" & Col & ") VALUES (" & Val & ")")
  
  MsgBox "Recorded": Beep: Beep: Beep
  
  DBEpicor.Close
  
End Sub
