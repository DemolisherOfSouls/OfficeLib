Attribute VB_Name = "StrFunc"
Option Explicit
Option Compare Text
Option Base 1

'String Function Library
'Version 1.0.3

'Current

Private Const NoParam = "@@@@@"

'Text Analysis

Public Function Contains(ByVal checktext As String, ByVal fortext As String) As Boolean
  Contains = InStr(checktext, fortext) = 0
End Function

Public Function ContainsAny(ByVal checktext As String, ByVal fortext1 As String, Optional ByVal fortext2 As String = NoParam, Optional ByVal fortext3 As String = NoParam) As Boolean
  ContainsAny = Contains(checktext, fortext1) Or Contains(checktext, fortext2) Or Contains(checktext, fortext3)
End Function

Public Function StartsWith(ByVal checktext As String, ByVal fortext As String) As Boolean
  StartsWith = InStr(checktext, fortext) = 1
End Function

Public Function StartsWithAny(ByVal checktext As String, ByVal fortext1 As String, Optional ByVal fortext2 As String = NoParam, Optional ByVal fortext3 As String = NoParam) As Boolean
  StartsWithAny = StartsWith(checktext, fortext1) Or StartsWith(checktext, fortext2) Or StartsWith(checktext, fortext3)
End Function

Public Function EndsWith(ByVal checktext As String, ByVal fortext As String) As Boolean
  EndsWith = InStr(checktext, fortext) + Len(fortext) = Len(checktext)
End Function

Public Function EndsWithAny(ByVal checktext As String, ByVal fortext1 As String, Optional ByVal fortext2 As String = NoParam, Optional ByVal fortext3 As String = NoParam) As Boolean
  EndsWithAny = EndsWith(checktext, fortext1) Or EndsWith(checktext, fortext2) Or EndsWith(checktext, fortext3)
End Function

'IsNumeric Variants

Public Function IsEmpty2(ByVal checkvalue) As Boolean
  IsEmpty2 = IsEmpty(checkvalue) Or IsNull(checkvalue) Or Len(checkvalue) = 0
End Function

Public Function IsNum(ByVal checkvalue) As Boolean
  IsNum = IsNumeric(checkvalue) Or IsDate(checkvalue)
End Function

Public Function IsText(ByVal checkvalue) As Boolean
  IsText = Not IsNumeric(checkvalue) And Not IsEmpty2(checkvalue) And Not IsDate(checkvalue)
End Function

'IfError Variants

Public Function IfText(ByVal checkvalue, ByVal valueiftext, Optional ByVal trimvalue As Boolean = False) As Variant
  If trimvalue Then checkvalue = Trim(checkvalue)
  
  IfText = IIf(IsText(checkvalue), checkvalue, valueiftext)
End Function

Public Function IfNum(ByVal checkvalue, ByVal valueifnum, Optional ByVal trimvalue As Boolean = True) As Variant
  If trimvalue Then checkvalue = Trim(checkvalue)
  
  IfNum = IIf(IsNum(checkvalue) Or IsEmpty2(checkvalue), valueifnum, checkvalue)
End Function

Public Function IfEmpty(ByVal checkvalue, ByVal valueifempty, Optional ByVal trimvalue As Boolean = True) As Variant
  If trimvalue Then checkvalue = Trim(checkvalue)

  IfEmpty = IIf(IsEmpty2(checkvalue), valueifempty, checkvalue)
End Function

'Make Plural if needed

Public Function Plural(ByVal initialtext As String, ByVal num As Integer, Optional ByVal appendtext As String = "s") As String
  Plural = CStr(num) & " " & initialtext & IIf(num <> 1, appendtext, "")
End Function
