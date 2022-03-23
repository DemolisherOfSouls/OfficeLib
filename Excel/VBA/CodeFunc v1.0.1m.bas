Attribute VB_Name = "CodeFunc"
Option Explicit
Option Compare Text
Option Base 1

'Formula Builder Function Library
'Version 1.0.1m

Private Const EQ As String = "="

Private Const FormRepIfError As String = "%%%"
Private Const FormRepLet As String = "%$%"

Private Const FormIfError As String = "IfError( " & FormRepIfError & ", """" )"
Private Const FormLet As String = "Let( val, " & FormRepLet & ", " & FormIfError & " )"

Private Const EqCheck As String = "^\s*=\s*"
Private Const IfErrorCheck As String = "^\s*=\s*(IFERROR)"
Private Const LetCheck As String = "^\s*=\s*(LET)"

Public Sub SurroundIfErrorBlock()
  On Error GoTo Invalid

  Dim Cell As Range, Frm As String
  
  For Each Cell In Selection.Cells
  
    If RegExTest(Cell.Formula2, IfErrorCheck) Then GoTo Skip
    If RegExTest(Cell.Formula2, LetCheck) Then GoTo Skip
    If Not RegExTest(Cell.Formula2, EqCheck) Then GoTo Skip
    
    Frm = RegExReplace(Cell.Formula2, EqCheck, "")
    
    Cell.Formula2 = EQ & Replace(FormIfError, FormRep, Frm)
    
Skip:
  Next Cell
Exit Sub

Invalid:
  ErrorMsg
End Sub


Public Sub SurroundLetBlock()
  On Error GoTo Invalid

  Dim Cell As Range, Frm As String, Valid As Boolean, Ln As Integer
  
  For Each Cell In Selection.Cells
  
    If RegExTest(Cell.Formula2, LetCheck) Then GoTo Skip
    If Not RegExTest(Cell.Formula2, EqCheck) Then GoTo Skip
    
    Frm = RegExReplace(Cell.Formula2, EqCheck, "")

    Cell.Formula2 = EQ & ReplaceMany(FormLet, FormRepLet, Frm, FormRepIfError, "val")
    
Skip:
  Next Cell
Exit Sub
  
Invalid:
  ErrorMsg
End Sub

