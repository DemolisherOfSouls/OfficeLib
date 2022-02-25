Attribute VB_Name = "CodeFunc"
Option Explicit
Option Compare Text
Option Base 1

'Formula Builder Function Library
'Version 1.0.0

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

  Dim Cell As Range, Frm As String, Valid As Boolean, Ln As Integer
  
  For Each Cell In Selection.Cells
  
    Let Frm = Cell.Formula2
    Let Ln = Len(Frm) - 1
    
    If StartsWith(Frm, EQ) And Not StartsWith(Frm, BLK_Let) Then
      Frm = Trim(Right(Frm, Ln))
      Valid = True
    End If
    
    If Not Valid Or StartsWith(Frm, BLK_IfError) Or StartsWith(Frm, BLK_Let) Then GoTo Skip
    
    Frm = (EQ & BLK_IfError & Right(Frm, Len(Frm) - 1) & END_IfError)
    Cell.Formula2 = Frm
    
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
    Let Ln = Len(Frm) - 1
    
    If StartsWith(Frm, EQ) And Not StartsWith(Frm, BLK_Let) Then
      Frm = Trim(Right(Frm, Ln))
      Valid = True
    End If
    
    If StartsWith(Frm, BLK_IfError) Then
      Frm = Right(Frm, Len(Frm) - LEN_IfError)
      Frm = Left(Frm, Len(Frm) - ELN_IfError)
    End If

    If Not Valid Or StartsWith(Frm, BLK_Let) Then GoTo Skip
    
    'Update Formula
    Frm = (EQ & BLK_Let & VAL_Let & MID_Let & Frm & MID_Let & BLK_IfError & VAL_Let & END_IfError & END_Let)
    Cell.Formula2 = Frm
    
Skip:
  Next Cell
Exit Sub
  
Invalid:
  ErrorMsg
End Sub

