Attribute VB_Name = "ExtFunc"
Option Explicit
Option Compare Text
Option Base 1

Const DC_Today As Integer = -2
Const DC_Invalid As Integer = -1
Const DC_Saturday As Integer = 0
Const DC_Sunday As Integer = 1
Const DC_Monday As Integer = 2
Const DC_Tuesday As Integer = 3
Const DC_Wednesday As Integer = 4
Const DC_Thursday As Integer = 5
Const DC_Friday As Integer = 6

Const RX_NoMatch As Integer = -10000

'Attribute IFTEXT.VB_Description = "SomeProcedure does something cool."
Public Function IfText(ByVal checkvalue, ByVal valueiftext, Optional ByVal trimvalue As Boolean = False) As Variant

    If trimvalue Then checkvalue = Trim(checkvalue)

    If IsNumeric(checkvalue) Or IsEmpty(checkvalue) Or IsNull(checkvalue) Or IsDate(checkvalue) Then
        IfText = checkvalue
    Else
        IfText = valueiftext
    End If

End Function

Public Function IfNum(ByVal checkvalue, ByVal valueifnum, Optional ByVal trimvalue As Boolean = True) As Variant

    If trimvalue Then checkvalue = Trim(checkvalue)
    
    If IsNumeric(checkvalue) Or IsDate(checkvalue) Then
        IfNum = valueifnum
    Else
        IfNum = checkvalue
    End If

End Function

Public Function IfEmpty(ByVal checkvalue, ByVal valueifempty, Optional ByVal trimvalue As Boolean = True) As Variant
    
    If trimvalue Then checkvalue = Trim(checkvalue)

    If IsEmpty(checkvalue) Or IsNull(checkvalue) Or Len(checkvalue) = 0 Then
        IfEmpty = valueifempty
    Else
        IfEmpty = checkvalue
    End If

End Function

Public Function IsThisWeek(ByVal datenumber, Optional ByVal startday As Integer = DC_Monday) As Boolean

    IsThisWeek = WeekStart() = WeekStart(datenumber)

End Function


Public Function DayCode(Optional ByVal datenumber = DC_Today) As Integer

    If datenumber = DC_Today Then datenumber = Date
    
    If datenumber < 0 Then
      
      DayCode = DC_Invalid
      Exit Function
      
    End If
    
    DayCode = Int(datenumber) Mod 7
    
End Function

Public Function WeekStart(Optional ByVal datenumber = DC_Today, Optional ByVal startday As Integer = DC_Monday) As Date
    
    If datenumber = DC_Today Then datenumber = Date
    
    WeekStart = (Int(datenumber / 7) * 7) + startday

End Function

Public Function WEEKRELATIVE(ByVal datenumber, Optional ByVal startday As Integer = DC_Monday, Optional ByVal base1index As Boolean = False) As Date
    
    WEEKRELATIVE = Int((datenumber - startday) / 7) - Int((Date - startday) / 7) + CInt(base1index)

End Function

Public Function DayStr(Optional ByVal datenumber = DC_Today) As Variant
    
    If datenumber = DC_Today Then datenumber = Date
  
    Select Case DayCode(datenumber)
        Case 0
            DayStr = "Saturday"
        Case 1
            DayStr = "Sunday"
        Case 2
            DayStr = "Monday"
        Case 3
            DayStr = "Tuesday"
        Case 4
            DayStr = "Wednesday"
        Case 5
            DayStr = "Thursday"
        Case 6
            DayStr = "Friday"
        Case DC_Invalid
            DayStr = CVErr(xlErrValue)
            
    End Select

End Function

Public Function Contains(ByVal checktext As String, ByVal fortext As String) As Boolean
    
    Contains = InStr(checktext, fortext)
    
End Function

Public Function StartsWith(ByVal checktext As String, ByVal fortext As String) As Boolean
    
    StartsWith = Left$(checktext, Len(fortext)) Is fortext
    
End Function

Public Function EndsWith(ByVal checktext As String, ByVal fortext As String) As Boolean
    
    EndsWith = Right(checktext, Len(fortext)) Is fortext
    
End Function

Public Function Plural(ByVal initialtext As String, ByVal num As Integer, Optional ByVal appendtext As String = "s") As String

    If num <> 1 Then
        Plural = CStr(num) & " " & initialtext & appendtext
    Else
        Plural = CStr(num) & " " & initialtext
    End If

End Function

Public Function FRACTION(ByVal s As String) As Double
  On Error GoTo BadInput

  Dim whole, upper, lower
  'Dim r As RegExp
  Dim r As Object: Set r = CreateObject("VBScript.RegExp")
  
  With r
  
    .Global = True
    .IgnoreCase = True
    .MultiLine = True
    .Pattern = Range("RegXFraction").Text
  
  End With
  
  whole = r.Execute(s).Item(0).SubMatches.Item(0)
  upper = r.Execute(s).Item(0).SubMatches.Item(1)
  lower = r.Execute(s).Item(0).SubMatches.Item(2)
  
  FRACTION = CDbl(upper) / CDbl(lower) + CInt(whole)
  Exit Function
  
BadInput:
  FRACTION = CVErr(xlErrNum)
End Function


Public Function XLIntersect(col As Variant, row As Variant)
  
  XLIntersect = Intersect(col, row)

End Function

Public Function GLookup(table, rval, row, cval, col)

  Dim rrng: Set rrng = Excel.WorksheetFunction.XLookup(rval, row, table)
  Dim crng: Set crng = Excel.WorksheetFunction.XLookup(cval, col, table)
  
  GLookup = Intersect(row, col)
  
End Function
