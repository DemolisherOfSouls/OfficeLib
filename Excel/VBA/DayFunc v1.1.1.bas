Attribute VB_Name = "DayFunc"
Option Explicit
Option Compare Text
Option Base 1

'Date Function Library
'Version 1.1.1

Public Const DTToday As Integer = -2
Public Const vbSaturday2 As Integer = 0

Public Function IsThisWeek(ByVal datenumber, Optional ByVal startday As Integer = vbMonday) As Boolean

  If datenumber < 0 Then GoTo Invalid

  IsThisWeek = WeekStart() = WeekStart(datenumber)
Exit Function
  
Invalid:
  IsThisWeek = CVErr(xlErrValue)
End Function


Public Function DayCode(Optional ByVal datenumber = DTToday, Optional ByVal Base7 As Boolean = False) As Integer

  If datenumber = DTToday Then datenumber = Date
  If datenumber < 0 Then GoTo Invalid
  
  DayCode = IIf(Base7 And Int(datenumber) Mod 7 = vbSaturday2, vbSaturday, Int(datenumber) Mod 7)
Exit Function
  
Invalid:
  DayCode = CVErr(xlErrValue)
End Function

Public Function WeekStart(Optional ByVal datenumber = DTToday, Optional ByVal startday As Integer = vbMonday) As Date
  
  If datenumber = DTToday Then datenumber = Date
  If datenumber < 0 Then GoTo Invalid
  
  WeekStart = (Int(datenumber / 7) * 7) + startday
Exit Function
  
Invalid:
  WeekStart = CVErr(xlErrValue)
End Function

Public Function WeekFrom(ByVal datenumber, Optional ByVal todatenumber = DTToday, Optional ByVal startday = vbMonday, Optional ByVal base1index As Boolean = False) As Integer

  If todatenumber = DTToday Then todatenumber = Date
  If todatenumber < 0 Then GoTo Invalid
  
  WeekFrom = Int((datenumber - startday) / 7) - Int((todatenumber - startday) / 7) - CInt(base1index)
Exit Function
  
Invalid:
  WeekFrom = CVErr(xlErrValue)
End Function

Public Function DayStr(Optional ByVal datenumber = DTToday) As String
  
  If datenumber = DTToday Then datenumber = Date
  If datenumber < 0 Then GoTo Invalid
 
  DayStr = WeekdayName(DayCode(datenumber))
Exit Function
  
Invalid:
  DayStr = CVErr(xlErrValue)
End Function

Public Function YearStart(Optional ByVal datenumber = DTToday) As Date

  If datenumber = DTToday Then datenumber = Date
  If datenumber < 0 Then GoTo Invalid
  
  YearStart = DateSerial(Year(datenumber), 1, 1)
Exit Function
  
Invalid:
  YearStart = CVErr(xlErrValue)
End Function

