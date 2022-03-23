Attribute VB_Name = "ScreenResize"
Option Explicit
Option Compare Text
Option Base 1

Private Declare PtrSafe Function GetSystemMetrics Lib "User32" (ByVal whichMetric As Long) As Long

Private Const SM_CXSCREEN As Long = 0
Private Const SM_CYSCREEN As Long = 1

Sub Resize_Screen(Screen As Object)

  Dim Swidth As Long: Swidth = GetSystemMetrics(SM_CXSCREEN)
  Dim Sheight As Long: Sheight = GetSystemMetrics(SM_CYSCREEN)
  Dim Wd, Hd, SZoom
  Dim cCont As Control
  
  Wd = Swidth - 1366
  Hd = Sheight - 768
  
  If Wd < 0 Or Hd < 0 Then
    If Hd < Wd Then
      SZoom = Sheight / 768
    Else
      SZoom = Swidth / 1366
    End If
  ElseIf Wd > 0 Or Hd > 0 Then
    If Hd > Wd Then
      SZoom = 1 'Sheight / 768
    Else
      SZoom = 1 'Swidth / 1366
    End If
  Else
      SZoom = 1
  End If
          
  Screen.Zoom = SZoom * 100
  Screen.Width = SZoom * 1025.25
  Screen.Height = SZoom * 570
  
  If SZoom < 1 Then
    For Each cCont In Screen.Controls
      If TypeOf cCont Is MSForms.Label Then
        cCont.Font.Size = 20 * SZoom * 1.3
        'cCont.AutoSize = True
      End If
      If TypeOf cCont Is MSForms.OptionButton Then
        cCont.Font.Size = 20 * SZoom * 1.3
        'cCont.Width = 120
      End If
      If TypeOf cCont Is MSForms.CheckBox Then
        cCont.Font.Size = 20 * SZoom * 1.3
        'cCont.Width = 120
      End If
      If TypeOf cCont Is MSForms.TextBox Then
        cCont.Font.Size = 20 * SZoom * 1.3
      End If
    Next
    Screen.Lbl_Main.Font.Size = 28 * SZoom * 1.2
  End If
  
End Sub
