Attribute VB_Name = "StrFunc"
Option Explicit
Option Compare Text
Option Base 1

'String Function Library
'Version 1.0.1

Public Const NoParam = "@@@@@"

Public Function Contains(ByVal checktext As String, ByVal fortext As String) As Boolean
  Contains = InStr(checktext, fortext)
End Function

Public Function ContainsAny(ByVal checktext As String, ByVal fortext1 As String, Optional ByVal fortext2 As String = NoParam, Optional ByVal fortext3 As String = NoParam) As Boolean
  ContainsAny = InStr(checktext, fortext1) Or InStr(checktext, fortext2) Or InStr(checktext, fortext3)
End Function

Public Function StartsWith(ByVal checktext As String, ByVal fortext As String) As Boolean
  StartsWith = Left$(CStr(checktext), Len(fortext)) = CStr(fortext)
End Function

Public Function StartsWithAny(ByVal checktext As String, ByVal fortext1 As String, Optional ByVal fortext2 As String = NoParam, Optional ByVal fortext3 As String = NoParam) As Boolean
  StartsWithAny = Left$(CStr(checktext), Len(fortext1)) = CStr(fortext1) Or Left$(CStr(checktext), Len(fortext2)) = CStr(fortext2) Or Left$(CStr(checktext), Len(fortext3)) = CStr(fortext3)
End Function

Public Function EndsWith(ByVal checktext As String, ByVal fortext As String) As Boolean
  EndsWith = Right$(CStr(checktext), Len(fortext)) = CStr(fortext)
End Function

Public Function Plural(ByVal initialtext As String, ByVal num As Integer, Optional ByVal appendtext As String = "s") As String
  Plural = CStr(num) & " " & initialtext & IIf(num <> 1, appendtext, "")
End Function


Public Function IfText(ByVal checkvalue, ByVal valueiftext, Optional ByVal trimvalue As Boolean = False) As Variant
  If trimvalue Then checkvalue = Trim(checkvalue)
  
  IfText = IIf(IsNum(checkvalue) Or IsEmpty2(checkvalue), checkvalue, valueiftext)
  
End Function

Public Function IfNum(ByVal checkvalue, ByVal valueifnum, Optional ByVal trimvalue As Boolean = True) As Variant

  If trimvalue Then checkvalue = Trim(checkvalue)
  
  IfNum = IIf(IsNum(checkvalue) Or IsEmpty2(checkvalue), valueifnum, checkvalue)

End Function

Public Function IsEmpty2(ByVal checkvalue) As Boolean

  IsEmpty2 = IsEmpty(checkvalue) Or IsNull(checkvalue) Or Len(checkvalue) = 0

End Function

Public Function IsNum(ByVal checkvalue) As Boolean

  IsNum = IsNumeric(checkvalue) Or IsDate(checkvalue)

End Function

Public Function IfEmpty(ByVal checkvalue, ByVal valueifempty, Optional ByVal trimvalue As Boolean = True) As Variant
  
  If trimvalue Then checkvalue = Trim(checkvalue)

  IfEmpty = IIf(IsEmpty2(checkvalue), valueifempty, checkvalue)

End Function
