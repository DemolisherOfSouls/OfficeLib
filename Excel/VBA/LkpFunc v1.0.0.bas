Attribute VB_Name = "LkpFunc"
Option Explicit
Option Compare Text
Option Base 1

'Lookup Function Library
'Version 1.0.0

Public Function XLIntersect(ByVal col As Variant, ByVal row As Variant)
 
  XLIntersect = Intersect(col, row)

End Function

Public Function GLookup(ByRef table as Variant, ByVal rval, ByVal row, ByVal cval, ByVal col) As Variant
  On Error GoTo Invalid

  Dim r As Range: Set r = Range.Find(What:=rval, LookIn:=row).EntireColumn
  Dim c As Range: Set c = Range.Find(What:=cval, LookIn:=col).EntireRow
  
  GLookup = Intersect(r, c).Value
 
Exit Sub
Invalid:
  ErrorMsg
End Function
