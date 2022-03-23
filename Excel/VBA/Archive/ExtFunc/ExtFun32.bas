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
  
Public Function IFTEXT(ByVal checkvalue, ByVal valueiftext, Optional ByVal trimvalue As Boolean = False) As Variant

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

    ISTHISWEEK = WEEKSTART() = WEEKSTART(datenumber)

End Function


Public Function DAYCODE(Optional ByVal datenumber = DC_Today) As Integer

    If datenumber = DC_Today Then datenumber = Date
    
    DAYCODE = Int(datenumber) Mod 7
    
End Function

Public Function WEEKSTART(Optional ByVal datenumber = DC_Today, Optional ByVal startday As Integer = DC_Monday) As Date
    
    If datenumber = DC_Today Then datenumber = Date
    
    WEEKSTART = (Int(Int(datenumber) / 7) * 7) + startday

End Function

Public Function WEEKRELATIVE(ByVal datenumber, Optional ByVal startday As Integer = DC_Monday, Optional ByVal base1index As Boolean = False) As Date
    
    WEEKRELATIVE = Int((datenumber - startday) / 7) - Int((Date - startday) / 7) + CInt(base1index)

End Function

Public Function DAYSTR(Optional ByVal datenumber = DC_Today) As String
    
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
            DAYSTR = "#VALUE!"
    End Select
        
End Function

Public Function CONTAINS(ByVal checktext As String, ByVal fortext As String) As Boolean
    
    CONTAINS = InStr(checktext, fortext)
    
End Function

Public Function STARTSWITH(ByVal checktext As String, ByVal fortext As String) As Boolean
    
    STARTSWITH = Left$(checktext, Len(fortext)) Is fortext
    
End Function

Public Function ENDSWITH(ByVal checktext As String, ByVal fortext As String) As Boolean
    
    ENDSWITH = Right(checktext, Len(fortext)) Is fortext
    
End Function

Public Function PLURAL(ByVal initialtext As String, ByVal num As Integer, Optional ByVal appendtext As String = "s") As String

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
