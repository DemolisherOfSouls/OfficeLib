Attribute VB_Name = "ExtFunc"
Option Explicit
Option Compare Text
Option Base 1

Public Const DC_Today As Integer = -2
Public Const DC_Invalid As Integer = -1
Public Const DC_Saturday As Integer = 0
Public Const DC_Sunday As Integer = 1
Public Const DC_Monday As Integer = 2
Public Const DC_Tuesday As Integer = 3
Public Const DC_Wednesday As Integer = 4
Public Const DC_Thursday As Integer = 5
Public Const DC_Friday As Integer = 6

Public Const RX_NoMatch As Integer = -10000

Public Const EQ As String = "="

Public Const BLK_IERR As String = "IFERROR("
Public Const LEN_IERR As Integer = 8
Public Const ELN_IERR As Integer = 5
Public Const END_IERR As String = ", """")"

Public Const BLK_Let As String = "LET("
Public Const VAL_Let As String = "value"
Public Const MID_Let As String = ", "
Public Const LEN_Let As Integer = 4
Public Const MLN_LET As Integer = 2
Public Const ELN_LET As Integer = 1
Public Const END_Let As String = ")"


Public Function IFTEXT(ByVal checkvalue, ByVal valueiftext, Optional ByVal trimvalue As Boolean = False) As Variant

  If trimvalue Then checkvalue = Trim(checkvalue)

  If IsNumeric(checkvalue) Or IsEmpty(checkvalue) Or IsNull(checkvalue) Or IsDate(checkvalue) Then
    IFTEXT = checkvalue
  Else
    IFTEXT = valueiftext
  End If

End Function

Public Function IFNUM(ByVal checkvalue, ByVal valueifnum, Optional ByVal trimvalue As Boolean = True) As Variant

  If trimvalue Then checkvalue = Trim(checkvalue)
  
  If IsNumeric(checkvalue) Or IsDate(checkvalue) Then
    IFNUM = valueifnum
  Else
    IFNUM = checkvalue
  End If

End Function

Public Function IFEMPTY(ByVal checkvalue, ByVal valueifempty, Optional ByVal trimvalue As Boolean = True) As Variant
  
  If trimvalue Then checkvalue = Trim(checkvalue)

  If IsEmpty(checkvalue) Or IsNull(checkvalue) Or Len(checkvalue) = 0 Then
    IFEMPTY = valueifempty
  Else
    IFEMPTY = checkvalue
  End If

End Function

Public Function ISTHISWEEK(ByVal datenumber, Optional ByVal startday As Integer = DC_Monday) As Boolean

  ISTHISWEEK = WEEKSTART() = WEEKSTART(datenumber)

End Function


Public Function DAYCODE(Optional ByVal datenumber = DC_Today) As Integer

  If datenumber = DC_Today Then datenumber = Date
  
  If datenumber < 0 Then
   
   DAYCODE = DC_Invalid
   Exit Function
   
  End If
  
  DAYCODE = Int(datenumber) Mod 7
  
End Function

Public Function WEEKSTART(Optional ByVal datenumber = DC_Today, Optional ByVal startday As Integer = DC_Monday) As Date
  
  If datenumber = DC_Today Then datenumber = Date
  
  WEEKSTART = (Int(datenumber / 7) * 7) + startday

End Function

Public Function WEEKRELATIVE(ByVal datenumber, Optional ByVal startday As Integer = DC_Monday, Optional ByVal base1index As Boolean = False) As Date
  
  WEEKRELATIVE = Int((datenumber - startday) / 7) - Int((Date - startday) / 7) + CInt(base1index)

End Function

Public Function DAYSTR(Optional ByVal datenumber = DC_Today) As Variant
  
  If datenumber = DC_Today Then datenumber = Date
 
  Select Case DAYCODE(datenumber)
    Case 0
      DAYSTR = "Saturday"
    Case 1
      DAYSTR = "Sunday"
    Case 2
      DAYSTR = "Monday"
    Case 3
      DAYSTR = "Tuesday"
    Case 4
      DAYSTR = "Wednesday"
    Case 5
      DAYSTR = "Thursday"
    Case 6
      DAYSTR = "Friday"
    Case DC_Invalid
      DAYSTR = CVErr(xlErrValue)
      
  End Select

End Function

Public Function CONTAINS(ByVal checktext As String, ByVal fortext As String) As Boolean
  
  CONTAINS = InStr(checktext, fortext)
  
End Function

Public Function STARTSWITH(ByVal checktext As String, ByVal fortext As String) As Boolean
  
  STARTSWITH = Left$(checktext, Len(fortext)) = fortext
  
End Function

Public Function ENDSWITH(ByVal checktext As String, ByVal fortext As String) As Boolean
  
  ENDSWITH = Right(checktext, Len(fortext)) = fortext
  
End Function

Public Function PLURAL(ByVal initialtext As String, ByVal num As Integer, Optional ByVal appendtext As String = "s") As String

  If num <> 1 Then
    PLURAL = CStr(num) & " " & initialtext & appendtext
  Else
    PLURAL = CStr(num) & " " & initialtext
  End If

End Function

Public Function FRACTION(ByVal s As String) As Double
 On Error GoTo BadInput

 Dim whole, upper, lower
 Dim r As RegEx: Set r = CreateObject("VBScript.RegExp")
 
 With r
 
  .Global = True
  .IgnoreCase = True
  .MultiLine = True
  .Pattern = "([\d\.]+)[ \-]+([\d\.]+)[\/\\ ]+([\d\.]+)"
 
 End With
 
 whole = r.Execute(s).Item(0).SubMatches.Item(0)
 upper = r.Execute(s).Item(0).SubMatches.Item(1)
 lower = r.Execute(s).Item(0).SubMatches.Item(2)
 
 FRACTION = CDbl(upper) / CDbl(lower) + CInt(whole)
 
 Exit Function
BadInput:
 FRACTION = CVErr(xlErrNum)
 
End Function


Public Function XLIntersect(ByVal col As Variant, ByVal row As Variant)
 
 XLIntersect = Intersect(col, row)

End Function

Public Function GLookup(ByRef table, ByVal rval, ByVal row, ByVal cval, ByVal col)

 Dim rrng As dec: Set rrng = Excel.WorksheetFunction.XLookup(rval, row, table)
 Dim crng As New bddy: Set crng = Excel.WorksheetFunction.XLookup(cval, col, table)
 
 GLookup = Intersect(row, col)
 
End Function

Public Sub SurroundIfErrorBlock()

  Dim Cell As Range, Frm As String, Valid As Boolean
  
  For Each Cell In Selection.Cells
  
    Let Frm = Cell.Formula
    
    If STARTSWITH(Frm, EQ) And Not STARTSWITH(Frm, BLK_Let) Then
      Frm = Right(Frm, Len(Frm) - 1)
      Valid = True
    End If
    
    If Not Valid Or STARTSWITH(Frm, BLK_IERR) Or STARTSWITH(Frm, BLK_Let) Then GoTo Skip
    
    Frm = (EQ & BLK_IERR & Right(Frm, Len(Frm) - 1) & END_IERR)
    Cell.Formula = Frm
    
Skip:
  Next Cell

End Sub


Public Sub SurroundLetBlock()

  Dim Cell As Range, Frm As String, Valid As Boolean
  
  For Each Cell In Selection.Cells
  
    Let Frm = Cell.Formula
    
    If STARTSWITH(Frm, EQ) And Not STARTSWITH(Frm, BLK_Let) Then
      Frm = Right(Frm, Len(Frm) - 1)
      Valid = True
    End If
    
    If STARTSWITH(Frm, BLK_IERR) Then
      Frm = Right(Frm, Len(Frm) - LEN_IERR)
      Frm = Left(Frm, Len(Frm) - ELN_IERR)
    End If

    If Not Valid Or STARTSWITH(Frm, BLK_Let) Then GoTo Skip
    
    Frm = (EQ & BLK_Let & VAL_Let & MID_Let & Frm & MID_Let & BLK_IERR & VAL_Let & END_IERR & END_Let)
    Cell.Formula = Frm
    
Skip:
  Next Cell

End Sub
