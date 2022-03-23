Attribute VB_Name = "CodeFunc1"
Option Explicit
Option Compare Text
Option Base 1

'`Formula Builder Function Library
'Version 1.0.3

'Imports
'Microsoft VBScript Regular Expressions 5.5

'History
' 1.0.2 - Condensed Functions
' 1.0.3 - Syntax Corrections

'Current

Private Const RXL_AllFormula As String = "=(.*)"
Private Const RXR_IfError As String = "=IFERROR($1, '')"
Private Const RXR_Let As String = "=LET(val, $1, IFERROR(val, ''))"

Private Sub SurroundBlock(ByVal Func As String, ByVal Template As String)
  On Error GoTo Invalid
  
  Dim Cell: For Each Cell In Selection.Cells
    If Not CheckFormulaFunction(Cell.Formula2, Func) Then
      Cell.Formula2 = RegExReplace(Cell.Formula2, RXL_AllFormula, Template)
    End If
  Next Cell
Exit Sub
Invalid:
  ErrorMsg
End Sub


Public Sub SurroundIfErrorBlock()
  SurroundBlock "LET|IFERROR", RXR_IfError
End Sub

Public Sub SurroundLetBlock()
  SurroundBlock "LET", RXR_Let
End Sub

