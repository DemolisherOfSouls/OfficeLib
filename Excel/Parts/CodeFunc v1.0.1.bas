Attribute VB_Name = "CodeFunc"
Option Explicit
Option Compare Text
Option Base 1

'Formula Builder Function Library
'Version 1.0.1

'Imports
'Microsoft VBScript Regular Expressions 5.5

Private Const EQ As String = "="

Private Const BLK_IfError As String = "IFERROR("
Private Const LEN_IfError As Integer = 8
Private Const ELN_IfError As Integer = 5
Private Const END_IfError As String = ", """")"

Private Const BLK_Let As String = "Let("
Private Const VAL_Let As String = "val"
Private Const MID_Let As String = ", "
Private Const LEN_Let As Integer = 4
Private Const END_Let As String = ")"

Public Sub SurroundIfErrorBlock()
  On Error GoTo Invalid
  Dim Cell As Range, Frm As String, Valid As Boolean
 
  For Each Cell In Selection.Cells
  
    Frm = Cell.Formula2
    Valid = Not RegExTest(Frm, "=\s*(IFERROR|LET)\s*\(", 0, 1)
    
    If Not Valid Then Goto Skip
    
    Cell.Formula2 = RegExReplace(Frm, "=(.*)", "=IFERROR($1, '')")
Skip:
  Next Cell
Exit Sub
Invalid:
  Call ErrorMsg
End Sub


Public Sub SurroundLetBlock()
  On Error GoTo Invalid

  Dim Cell As Range, Frm As String, Valid As Boolean, Ln As Integer
  
  For Each Cell In Selection.Cells
  
    Let Frm = Cell.Formula2
    Valid = Not RegExTest(Frm, "=\s*(LET)\s*\(", 0, 1)

    If Not Valid Then GoTo Skip
    
    Cell.Formula2 = RegExReplace(Frm, "=(.*)", "=LET(val, $1, IFERROR(val, ''))")
Skip:
  Next Cell
Exit Sub
Invalid:
  ErrorMsg
End Sub

