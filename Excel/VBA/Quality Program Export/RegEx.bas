Attribute VB_Name = "RegEx"
Option Explicit
Option Compare Text
Option Base 1

'Quality Program RegEx Storage

Private Const LineReduce As String = "\n\s*\n(\s*\n)?"
Private Const LinkReader As String = "(?:(\d)(?:SD))?(?:(\d)(?:HD))?"
Private Const MeshDesc As String = "(?:CB\d )?([A-Z]{0,3}[\d.]+-[\d./]+-[\d.]+F?|No Mesh)\b"
Private Const BeltWidth As String = "(?:Max Fabric Width:[\s\S]*?)?(?:(?:(?:Overall)?Belt)?O.A""""?|width| x|OVERALL \(in\))[\s.:]*([\d./-]+)""""?[ \t]*(in\.?(?:ches)?|m(?:trs|m)?|feet|$)"
Private Const FabricWidth As String = "fab[\w]*[\W]+?(?:([\d.]+)[ \t]*([\w]*)|([\w\/-]+))"
Private Const CenterLinkLoc As String = "loc[\w]*[\W]+?(?:([\d.]+)[ \t]*([\w]*)|([\w\/-]+))"

'Send the expressions to excel

Public Function GleenMeshDesc(ByVal src As Variant) As String
  GleenMeshDesc = RegExQuick(CStr(src), MeshDesc)
End Function

Public Function GleenBeltWidth(ByVal src As Variant) As Double
  Dim Width As Double: Width = CDbl(RegExExecute(CStr(src), BeltWidth, 0, 0))
  Dim Units As String: Units = RegExExecute(CStr(src), BeltWidth, 0, 1)
  GleenBeltWidth = IIf(InStr(0, Units, "m"), Width / 25.4, Width)
End Function

Public Function GleenFabricWidth(ByVal src As Variant) As Double
  Dim Width As Double: Width = CDbl(RegExExecute(CStr(src), FabricWidth, 0, 0))
  Dim Units As String: Units = RegExExecute(CStr(src), FabricWidth, 0, 1)
  GleenFabricWidth = IIf(InStr(0, Units, "m"), Width / 25.4, Width)
End Function

Public Function GleenCenterLinkLoc(ByVal src As Variant) As Double
  Dim CLLoc As Double: CLLoc = CDbl(RegExExecute(CStr(src), CenterLinkLoc, 0, 0))
  Dim Units As String: Units = RegExExecute(CStr(src), CenterLinkLoc, 0, 1)
  GleenCenterLinkLoc = IIf(InStr(0, Units, "m"), CLLoc / 25.4, CLLoc)
End Function

Public Function Compact(ByVal Comment As String) As String
  Compact = RegExReplace(Comment, RegEx.LineReduce, vbNewLine)
End Function

Public Function ParseBarLinks(ByVal S As String) As BarLinkCounts
  On Error GoTo Invalid
  Dim Ret As New BarLinkCounts
  With RegExO
    .Global = True
    .IgnoreCase = True
    .MultiLine = True
    .Pattern = RegEx.LinkReader
    With .Execute(S).Item(0).SubMatches
      Ret.Sd = .Item(0)
      Ret.Hd = .Item(1)
    End With
  End With
  ParseBarLinks = Ret
End Function
