Attribute VB_Name = "ExtFunc"
Option Explicit
Option Compare Text
Option Base 1

Const DC_Invalid As Integer = -1
Const DC_Saturday As Integer = 0
Const DC_Sunday As Integer = 1
Const DC_Monday As Integer = 2
Const DC_Tuesday As Integer = 3
Const DC_Wednesday As Integer = 4
Const DC_Thursday As Integer = 5
Const DC_Friday As Integer = 6
  
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

Public Function ISTHISWEEK(ByVal datenumber, Optional ByVal startday As Integer = DC_Monday) As Boolean

    ISTHISWEEK = WEEKSTART(Date) = WEEKSTART(datenumber)

End Function


Public Function DAYCODE(ByVal datenumber) As Integer

    DAYCODE = Int(datenumber) Mod 7
    
End Function

Public Function WEEKSTART(ByVal datenumber, Optional ByVal startday = DC_Monday) As Variant
    
    WEEKSTART = (Int(Int(datenumber) / 7) * 7) + startday

End Function

Public Function WEEKRELATIVE(ByVal datenumber As Variant, Optional ByVal startday As Integer = DC_Monday, Optional ByVal base1index As Boolean = False) As Variant
    
    If IsError(datenumber) Then WEEKRELATIVE = ""

    WEEKRELATIVE = Int((datenumber - startday) / 7) - Int((Date - startday) / 7) + CInt(base1index)

End Function

Public Function DAYSTR(ByVal datenumber) As String
    
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
            DAYSTR = Error(xlErrNum)
    End Select
        
End Function

Public Function CONTAINS(ByVal checktext, ByVal fortext) As Boolean
    
    CONTAINS = InStr(checktext, fortext)
    
End Function

Public Function STARTSWITH(ByVal checktext, ByVal fortext) As Boolean
    
    STARTSWITH = Left(checktext, Len(fortext)) Is fortext
    
End Function

Public Function ENDSWITH(ByVal checktext, ByVal fortext) As Boolean
    
    ENDSWITH = Right(checktext, Len(fortext)) Is fortext
    
End Function

Public Function PLURAL(ByVal initialtext, ByVal num, Optional appendtext = "s") As String

    If num <> 1 Then
        PLURAL = CStr(num) & " " & initialtext & appendtext
    Else
        PLURAL = CStr(num) & " " & initialtext
    End If

End Function



