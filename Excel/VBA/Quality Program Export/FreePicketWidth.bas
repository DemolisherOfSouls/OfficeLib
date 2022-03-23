Attribute VB_Name = "FreePicketWidth"
Option Explicit
Option Compare Text
Option Base 1

Function FwFreePicketWidth(ByRef PartNum As String, BeltWidth As Double, Width As Double, BarLinks As String, Minimum As Boolean) As Double
  Dim rs As New ADODB.Recordset
  Dim BarLinkThickness As Double
  
  Const Sd As Double = 0.062
  Const Hd As Double = 0.09
  
  Select Case PartNum
    Case "OFE1", "OFE2", "SROFG1", "SROFG3"
      'Calculate Bar Link Thickness
      Dim C As BarLinkCounts: Set C = ParseBarLinks(BarLinks)
      BarLinkThickness = C.Sd * Sd + C.Hd * Hd
      Dim MinSize As Double: MinSize = IIf(Minimum, 0.188, 0.062)
      
      'Calculate Free Picket Width
      If PartNum = "SROFG1" Or PartNum = "SROFG3" Then 'Small Radius Belt
        FwFreePicketWidth = (Max(BeltWidth - 24, 0) * 0.002) - BarLinkThickness * 2 - MinSize
      Else 'Regular Belt
        FwFreePicketWidth = (Max(Width - 24, 0) * 0.004) - BarLinkThickness * 2 - MinSize
      End If
    Case Else
      Set rs = GetSQLData("FWFreePicketWidth", "", "", "PartNum='" & PartNum & "' AND BarLinks ='" & BarLinks & "' AND MinWidth<" & Width & " AND MaxWidth>=" & Width, "", CSEngineer)

      FwFreePicketWidth = CDbl(rs.Fields(IIf(Minimum, "MinPicketWidth", "MaxPicketWidth")))
   End Select
End Function

Function GetSQLData(ByVal SQLTable As String, ByVal Selects As String, ByVal Joins As String, ByVal Constraints As String, ByVal Order As String, ByVal Connection As String) As ADODB.Recordset
  
  Dim rs As New ADODB.Recordset
  
  Selects = "SELECT " & IIf(Len(Selects) > 0, Selects, "*")
  If Len(Constraints) > 0 Then Constraints = "WHERE " & Constraints
  If Len(Order) > 0 Then Order = "ORDER BY " & Order

  DBEpicor.Open
  
  Set rs.ActiveConnection = DBEpicor
  rs.Open Selects & " FROM " & SQLTable & " " & Joins & " " & Constraints & " " & Order & ";"
  If rs.EOF Or rs.BOF Then Exit Function
  
  Set GetSQLData = rs
  
End Function

