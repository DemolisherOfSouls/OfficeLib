Attribute VB_Name = "ExtFunc"
Option Explicit
Option Compare Text
Option Base 1

  Public Const DC_Today = -2
  Public Const DC_Invalid = -1
  Public Const DC_Saturday = 0
  Public Const DC_Sunday = 1
  Public Const DC_Monday = 2
  Public Const DC_Tuesday = 3
  Public Const DC_Wednesday = 4
  Public Const DC_Thursday = 5
  Public Const DC_Friday = 6

Public Function IFTEXT(ByVal checkvalue, ByVal valueiftext, Optional ByVal trimvalue As Boolean = False) As Variant

    If trimvalue Then checkvalue = Trim(checkvalue)

    If IsNumeric(checkvalue) Or IsEmpty(checkvalue) Or _
        IsNull(checkvalue) Or IsDate(checkvalue) Then
        IFTEXT = checkvalue
    Else
        IFTEXT = valueiftext
    End If

End Function

Public Function IFNUM(ByVal checkvalue, ByVal valueifnum, Optional ByVal trimvalue As Boolean = False) As Variant

    If trimvalue Then checkvalue = Trim(checkvalue)
    
    If IsNumeric(checkvalue) Or IsDate(checkvalue) Then
        IFNUM = valueifnum
    Else
        IFNUM = checkvalue
    End If

End Function

Public Function IFEMPTY(ByVal checkvalue, ByVal valueifempty, Optional ByVal trimvalue As Boolean = True) As Variant
    
    If trimvalue Then checkvalue = Trim(checkvalue)

    If IsEmpty(checkvalue) Or IsNull(checkvalue) Or Len(checkvalue) = 0  or Then
        IFEMPTY = valueifempty
    Else
        IFEMPTY = checkvalue
    End If

End Function

Public Function ISTHISWEEK(ByVal datenumber, Optional ByVal startday = DC_Monday) As Boolean

    ISTHISWEEK = WEEKSTART(Date) = WEEKSTART(datenumber)

End Function

Public Function DAYCODE(ByVal datenumber) As Integer

    DAYCODE = Int(datenumber) Mod 7
    
End Function

Public Function WEEKSTART(ByVal datenumber, Optional ByVal startday As Integer = DC_Monday) As Date

    WEEKSTART = (Int(datenumber / 7) * 7) + startday

End Function

Public Function WEEKRELATIVE(ByVal datenumber, Optional ByVal startday = DC_Monday, Optional ByVal base1index As Boolean = False) As Integer

    WEEKRELATIVE = Int((datenumber - startday) / 7) - Int((Date - startday) / 7) + CInt(base1index)

End Function

Public Function DAYSTR(ByVal datenumber As Integer) As String
    
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

Public Function CONTAINS(ByVal checktext, ByVal fortext) As Boolean
    
    CONTAINS = InStr(checktext, fortext)
    
End Function

Public Function STARTSWITH(ByVal checktext, ByVal fortext) As Boolean
    
    STARTSWITH = Left(checktext, Len(fortext)) = fortext
    
End Function

Public Function ENDSWITH(ByVal checktext, ByVal fortext) As Boolean
    
    ENDSWITH = Right(checktext, Len(fortext)) = fortext
    
End Function

Public Function PLURAL(ByVal initialtext, ByVal num, Optional ByVal appendtext = "s") As String

    If num <> 1 Then
        PLURAL = CStr(num) & " " & initialtext & appendtext
    Else
        PLURAL = CStr(num) & " " & initialtext
    End If

End Function



