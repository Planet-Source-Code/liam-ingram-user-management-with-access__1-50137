Attribute VB_Name = "effects"
'----------------------------------------------------
'/ Created by Teh Ming Han                          /
'/ E-mail: teh_minghan@hotmail.com                  /
'/ Singapore                                        /
'/ 8 September 2001                                 /
'/                                                  /
'/ REMEMBER TO RATE THIS AS-->EXCELLENT             /
'/                                                  /
'/ www.planet-source-code.com/vb/                   /
'/ Have you voted?                                  /
'----------------------------------------------------


Option Explicit
'------------------------------------------
' define type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

' declare api calls
Private Declare Function GetWindowRect Lib "User32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetDC Lib "User32" (ByVal hwnd As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
'-------------------Explode-----------------^^^^^^^^

Public Function Explode(newform As Form, Increment1 As Integer)
'for LOAD & UNLOAD
Dim increment As Integer
increment = Increment1 * 1000
Dim Size As RECT     ' setup form as rect type
GetWindowRect newform.hwnd, Size

Dim FormWidth, FormHeight As Integer ' establish dimension variables
FormWidth = (Size.Right - Size.Left)
FormHeight = (Size.Bottom - Size.Top)

Dim TempDC
TempDC = GetDC(ByVal 0&)  ' obtain memory dc for resizing

Dim Count, LeftPoint, TopPoint, nWidth, nHeight As Integer                                      ' establish resizing variables
For Count = 1 To increment    ' loop to new sizes
    nWidth = FormWidth * (Count / increment)
    nHeight = FormHeight * (Count / increment)
    LeftPoint = Size.Left + (FormWidth - nWidth) / 2
    TopPoint = Size.Top + (FormHeight - nHeight) / 2
    Rectangle TempDC, LeftPoint, TopPoint, LeftPoint + nWidth, TopPoint + nHeight     ' draw rectangles to build form
DoEvents
Next Count

DeleteDC (TempDC) ' release  memory resource
End Function

Public Function rollup(newform As Form, Increase As Integer)
'for UNLOAD ONLY
If newform.ScaleMode <> 1 Then Exit Function
Do
newform.Height = newform.Height - Increase
DoEvents
Loop Until newform.Height < 410
End Function

Public Function closein(frm As Form)
'for UNLOAD ONLY
Dim GotoVal
Dim Gointo

    GotoVal = frm.Height / 2
    For Gointo = 1 To GotoVal
        DoEvents
        frm.Height = frm.Height - 100
        frm.Top = (Screen.Height - frm.Height) \ 2
        If frm.Height <= 500 Then Exit For
    Next Gointo
horiz:
    frm.Height = 30
    GotoVal = frm.Width / 2
    For Gointo = 1 To GotoVal
        DoEvents
        frm.Width = frm.Width - 100
        frm.Left = (Screen.Width - frm.Width) \ 2
        If frm.Width <= 2000 Then Exit For
    Next Gointo

End Function

Public Function slideup(frm As Form)
'for UNLOAD ONLY
Do
frm.Top = frm.Top - 1
frm.Left = frm.Left - 1
 If frm.Top < 2 Then Exit Function
 If frm.Left < 2 Then Exit Function
 If frm.Height > 500 Then frm.Height = frm.Height - 5
DoEvents
Loop

End Function

Public Function rush(newform As Form)
'for UNLOAD ONLY
If newform.ScaleMode <> 1 Then Exit Function
Do
If newform.Height > 410 Then
 newform.Height = newform.Height - 1
 newform.Top = (Screen.Height - newform.Height) / 2
 DoEvents
End If
newform.Left = newform.Left + 2
DoEvents
Loop Until newform.Left >= Screen.Width - 1

End Function

Public Function spiral(newform As Form)
'prepare to get dizzy
'for unload only
On Error GoTo err

 Do
  newform.Top = newform.Top - 5
  DoEvents
 Loop Until newform.Top < 2

 Do
  newform.Left = newform.Left + 5
  DoEvents
 Loop Until newform.Left > Screen.Width - newform.Width

 Do
  newform.Top = newform.Top + 5
  DoEvents
 Loop Until newform.Top > Screen.Height - newform.Height
 
 Do
  newform.Left = newform.Left - 5
  DoEvents
 Loop Until newform.Left < 2
Exit Function

err:
Exit Function

End Function

Public Function openall(frm As Form)
'for UNLOAD ONLY
Do
frm.Height = frm.Height + 1
frm.Width = frm.Width + 1
frm.Top = (Screen.Height - frm.Height) / 2
frm.Left = (Screen.Width - frm.Width) / 2
DoEvents
If frm.Height >= Screen.Height Then Exit Function
If frm.Width >= Screen.Width Then Exit Function
Loop

End Function

Public Function pressed(frm As Form)
'FOR UNLOAD ONLY
Do
 frm.Height = frm.Height - 1
 frm.Width = frm.Width + 1
 frm.Top = (Screen.Height - frm.Height) / 2
 frm.Left = (Screen.Width - frm.Width) / 2
 DoEvents
Loop Until frm.Width >= Screen.Width
 
End Function

Public Function funnyshape(frm As Form)
'FOR LOAD & UNLOAD
'
'FOR LOAD, YOU HAVE TO PUT THIS
'IN A CONTROL'S GOTFOCUS EVENT TO WORK

Dim h, w, t, l As Integer
h = frm.Height
w = frm.Width
t = frm.Top
l = frm.Left
Do
 frm.Height = frm.Height - 1
 frm.Width = frm.Width + 1
 frm.Top = (Screen.Height - frm.Height) / 2
 frm.Left = (Screen.Width - frm.Width) / 2
 DoEvents
Loop Until frm.Width >= Screen.Width

Do
 frm.Height = frm.Height + 1
 frm.Width = frm.Width - 1
 frm.Top = (Screen.Height - frm.Height) / 2
 frm.Left = (Screen.Width - frm.Width) / 2
 DoEvents
Loop Until frm.Width <= w
frm.Height = h
frm.Width = w
frm.Top = t
frm.Left = l

End Function

Public Function bounce_go(frm As Form)
'FOR UNLOAD ONLY
Dim num As Integer
frm.Top = (Screen.Height - frm.Height) / 10
frm.Left = 10

For num = 1 To 2
 Do
  frm.Left = frm.Left + 1
  frm.Top = frm.Top + 1
  DoEvents
 Loop Until (frm.Top + frm.Height) >= Screen.Height

 Do
  frm.Left = frm.Left + 1
  frm.Top = frm.Top - 1
  DoEvents
 Loop Until frm.Top <= (Screen.Height - frm.Height) / 2
Next num
End Function

Public Function bounce_updown(frm As Form)
'FOR UNLOAD ONLY
Dim num, d As Integer
frm.Top = (Screen.Height - frm.Height) / 10
d = 2
For num = 1 To 3
 Do
  frm.Top = frm.Top + 0.6
  DoEvents
 Loop Until (frm.Top + frm.Height) >= Screen.Height

 Do
  frm.Top = frm.Top - 0.6
  DoEvents
 Loop Until frm.Top <= (Screen.Height - frm.Height) / d
Next num
d = d + 3
End Function
