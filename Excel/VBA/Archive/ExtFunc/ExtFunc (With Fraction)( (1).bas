Attribute VB_Name = "ExtFunc"
Option Explicit
Option Compare Text
Option Base 1

Public Enum DayCodeValue

  DC_Today = -2
  DC_Invalid = -1
  DC_Saturday = 0
  DC_Sunday = 1
  DC_Monday = 2
  DC_Tuesday = 3
  DC_Wednesday = 4
  DC_Thursday = 5
  DC_Friday = 6

End Enum

Public Function IFTEXT(ByVal checkvalue, ByVal valueiftext, Optional ByVal trimvalue As Boolean = True) As Variant

  If trimvalue Then checkvalue = Trim(checkvalue)

  If IsNumeric(checkvalue) Or IsEmpty(checkvalue) Or _
    IsNull(checkvalue) Or IsDate(checkvalue) Then
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

Public Function ISTHISWEEK(Optional ByVal datenumber As Variant = DC_Today, Optional ByVal startday = DC_Monday) As Boolean

  If datenumber = DC_Today Then datenumber = Date

  ISTHISWEEK = WEEKSTART() = WEEKSTART(datenumber)

End Function

Public Function DAYCODE(Optional ByVal datenumber As Variant = DC_Today) As Variant

  If datenumber = DC_Today Then datenumber = Date

  DAYCODE = Int(IFTEXT(datenumber, DC_Invalid)) Mod 7

End Function

Public Function WEEKSTART(Optional ByVal datenumber As Variant = DC_Today, Optional ByVal startday As DayCodeValue = DC_Monday) As Date
Foot
  If datenumber = DC_Today Then datenumber = Date

  WEEKSTART = (IFTEXT(datenumber, 0) \ 7 * 7) + startday

End Function

Public Function WEEKRELATIVE(ByVal datenumber As Variant, Optional ByVal startday As DayCodeValue = DC_Monday, Optional ByVal base1index As Boolean = False) As Variant

  WEEKRELATIVE = Int((datenumber - startday) / 7) - Int((Date - startday) / 7) + CInt(base1index)

End Function

Public Function DAYSTR(Optional ByVal datenumber As Variant = DC_Today) As Variant

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
    Case Else
      DAYSTR = Error(xlErrNum)
  End Select

End Function

Public Function CONTAINS(ByVal checktext As Variant, ByVal fortext) As Boolean

  CONTAINS = InStr(checktext, fortext)

End Function

Public Function STARTSWITH(ByVal checktext, ByVal fortext) As Boolean

  STARTSWITH = Left(checktext, Len(fortext)) = fortext

End Function

Public Function ENDSWITH(ByVal checktext, ByVal fortext) As Boolean

  ENDSWITH = Right(checktext, Len(fortext)) = fortext

End Function

Public Function PLURAL(ByVal initialtext, ByVal num, Optional ByVal appendtext As String = "s") As String

  If num <> 1 Then
    PLURAL = CStr(num) & " " & initialtext & appendtext
  Else
    PLURAL = CStr(num) & " " & initialtext
  End If

End Function

Public Function FRACTION(ByVal s As String) As Double

  Dim whole, upper, lower
  'Dim r As RegExp
  Dim r As Object: Set r = CreateObject("VBScript.RegExp")
  
  
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
Err:
  FRACTION = -10000

End Function
