Attribute VB_Name = "ErrFunc"
Option Explicit
Option Compare Text
Option Base 1

'`Error Handling Function Library
'Version 1.0.0

Public Function ErrorMsg()
  ErrorMsg = MsgBox("Error: " & Err(), vbOKOnly, "Error")
End Function
