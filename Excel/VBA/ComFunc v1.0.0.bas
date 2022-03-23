Attribute VB_Name = "ComFunc"
Option Explicit
Option Compare Text
Option Base 1

'Common Function Library
'Version 1.0.0

Public Function Max(ByVal X, ByVal y)
  Max = IIf(X > y, X, y)
End Function

Public Function Min(ByVal X, ByVal y)
  Min = IIf(X < y, X, y)
End Function
