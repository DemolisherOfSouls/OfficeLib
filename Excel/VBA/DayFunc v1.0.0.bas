Attribute VB_Name = "DayFunc"
Option Explicit
Option Compare Text
Option Base 1

'Date Function Library
'Version 1.0.0

Public Const DC_Today As Integer = -2
Public Const DC_Invalid As Integer = -1
Public Const DC_Saturday As Integer = 0
Public Const DC_Sunday As Integer = 1
Public Const DC_Monday As Integer = 2
Public Const DC_Tuesday As Integer = 3
Public Const DC_Wednesday As Integer = 4
Public Const DC_Thursday As Integer = 5
Public Const DC_Friday As Integer = 6

Public Function IsThisWeek(ByVal datenumber, Optional ByVal startday As Integer = DC_Monday) As Boolean

  IsThisWeek = WeekStart() = WeekStart(datenumber)

End Function


Public Function DayCode(Optional ByVal datenumber = DC_Today) As Integer

  If datenumber = DC_Today Then datenumber = Date
  If datenumber < 0 Then GoTo Invalid
  
  DayCode = Int(datenumber) Mod 7
Exit Function
  
Invalid:
  DayCode = DC_Invalid
End Function

Public Function WeekStart(Optional ByVal datenumber = DC_Today, Optional ByVal startday As Integer = DC_Monday) As Date
  
  If datenumber = DC_Today Then datenumber = Date
  If datenumber < 0 Then GoTo Invalid
  
  WeekStart = (Int(datenumber / 7) * 7) + startday
Exit Function
  
Invalid:
  WeekStart = CVErr(xlErrValue)
End Function

Public Function WeekFrom(ByVal datenumber, Optional ByVal todatenumber = DC_Today, Optional ByVal startday = DC_Monday, Optional ByVal base1index As Boolean = False) As Integer

  If todatenumber = DC_Today Then todatenumber = Date
  
  WeekFrom = Int((datenumber - startday) / 7) - Int((todatenumber - startday) / 7) - CInt(base1index)

End Function

Public Function DayStr(Optional ByVal datenumber = DC_Today) As String
  
  If datenumber = DC_Today Then datenumber = Date
  If datenumber < 0 Then GoTo Invalid
 
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
      GoTo Invalid
  End Select
Exit Function
  
Invalid:
  DayStr = CVErr(xlErrValue)
End Function

Public Function YearStart(Optional ByVal datenumber = DC_Today) As Date

  If datenumber = DC_Today Then datenumber = Date
  
  YearStart = DateSerial(Year(datenumber), 1, 1)

End Function
