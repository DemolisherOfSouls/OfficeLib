Attribute VB_Name = "LkpFunc"
Option Explicit
Option Compare Text
Option Base 1

'Lookup Function Library
'Version 1.1.0

'Current

Public Function XLIntersect(ByVal r1, ByVal r2)
  XLIntersect = Intersect(r1, r2)
End Function

Public Function XLEntireRow(ByVal Cell)
  XLEntireRow = Cell.EntireRow
End Function

Public Function XLEntireColumn(ByVal Cell)
  XLEntireColumn = Cell.EntireColumn
End Function

Public Function GLookup(ByRef table, ByVal rval, ByVal row, ByVal cval, ByVal col)
  On Error GoTo Invalid

  Dim r As Range: Set r = Range.Find(rval, LookIn:=row).EntireColumn
  Dim C As Range: Set C = Range.Find(cval, LookIn:=col).EntireRow
  
  GLookup = Intersect(r, C)
 
Exit Function
Invalid:
  ErrorMsg
End Function
