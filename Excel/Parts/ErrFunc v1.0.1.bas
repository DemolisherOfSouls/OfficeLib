Attribute VB_Name = "ErrFunc"
Option Explicit
Option Compare Text
Option Base 1

'Error Handling Function Library
'Version 1.0.1

'Current

Public Function ErrorMsg() As VbMsgBoxResult
  ErrorMsg = MsgBox("Error: " & Err(), vbExclamation, "Error")
End Function
