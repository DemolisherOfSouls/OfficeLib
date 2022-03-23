Attribute VB_Name = "ExtFunc"
Option Explicit
Option Compare Text
Option Base 1

'Lookup Function Library

Public Const EQ As String = "="

Public Const BLK_IfError As String = "IFERROR("
Public Const LEN_IfError As Integer = 8
Public Const ELN_IfError As Integer = 5
Public Const END_IfError As String = ", """")"

Public Const BLK_Let As String = "Let("
Public Const VAL_Let As String = "val"
Public Const MID_Let As String = ", "
Public Const LEN_Let As Integer = 4
Public Const END_Let As String = ")"

Public Function XLIntersect(ByVal col As Variant, ByVal row As Variant)
 
 XLIntersect = Intersect(col, row)

End Function

Public Function GLookup(ByRef table, ByVal rval, ByVal row, ByVal cval, ByVal col) As Variant
  On Error GoTo Invalid

 Dim r As Range: Set r = Range.Find(What:=rval, LookIn:=row).EntireColumn
 
 Dim c As Range: Set c = Range.Find(What:=cval, LookIn:=col).EntireRow
 
 GLookup = Intersect(r, c).Value
 
Exit Sub
Invalid:
  Call ErrorMsg
 
End Function

Public Sub SurroundIfErrorBlock()
  On Error GoTo Invalid

  Dim Cell As Range, Frm As String, Valid As Boolean, Ln As Integer
  
  For Each Cell In Selection.Cells
  
    Let Frm = Cell.Formula
    Let Ln = Len(Frm) - 1
    
    If StartsWith(Frm, EQ) And Not StartsWith(Frm, BLK_Let) Then
      Frm = Trim(Right(Frm, Ln))
      Valid = True
    End If
    
    If Not Valid Or StartsWith(Frm, BLK_IfError) Or StartsWith(Frm, BLK_Let) Then GoTo Skip
    
    Frm = (EQ & BLK_IfError & Right(Frm, Len(Frm) - 1) & END_IfError)
    Cell.Formula = Frm
    
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
  
    Let Frm = Cell.Formula
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
    Cell.Formula = Frm
    
Skip:
  Next Cell
Exit Sub
  
Invalid:
  Call ErrorMsg
  
End Sub

Public Function ErrorMsg()
  ErrorMsg = MsgBox("Error: " & Err(), vbOKOnly, "Error", Nothing)
End Function

