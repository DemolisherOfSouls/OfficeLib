Attribute VB_Name = "LinkFunc"
Option Explicit
Option Compare Text

Sub LinkFromContent()
Attribute LinkFromContent.VB_ProcData.VB_Invoke_Func = "L\n14"

    Dim cs As Range, c As Range
    Set cs = Selection.Cells
    
    If (cs Is Nothing) Then Exit Sub
    
    For Each c In cs
    
        Dim add As String
    
        If Not InStr(c.Value, "://") Then
            add = "https://" & c.Value
        Else
            add = c.Value
        End If
        
        If c.Text <> "" Then
            ActiveSheet.Hyperlinks.add Anchor:=c, Address:=add, TextToDisplay:=c.Text
        End If
        
    Next c

End Sub
