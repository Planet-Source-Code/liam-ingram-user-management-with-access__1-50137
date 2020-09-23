VERSION 5.00
Begin VB.Form frmXPMenu 
   BackColor       =   &H00F7F8F9&
   BorderStyle     =   0  'None
   ClientHeight    =   3270
   ClientLeft      =   4410
   ClientTop       =   5955
   ClientWidth     =   2340
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   218
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   156
   ShowInTaskbar   =   0   'False
   Tag             =   "XPMenu"
   Begin VB.Timer tmrActive 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   855
      Top             =   1815
   End
   Begin VB.PictureBox picMenuBuffer 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3165
      Left            =   0
      ScaleHeight     =   211
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   156
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2340
      Begin VB.Timer tmrHover 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   870
         Top             =   930
      End
      Begin VB.PictureBox picPopup 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   9.75
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -435
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   2
         Top             =   1740
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.PictureBox picIcon 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   1
         Top             =   690
         Visible         =   0   'False
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmXPMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public XPMenuClass As clsXPMenu
Private Declare Function GetActiveWindow Lib "User32" () As Long
Private Declare Function WindowFromPoint Lib "User32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Type POINTAPI
        x As Long
        y As Long
End Type
Private Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long

Public upY As Single

Private Sub Form_Click()
    Dim selectedItem As Long
    selectedItem = XPMenuClass.GetHilightedItem(upY)
    
    If XPMenuClass.IsTextItem(CInt(selectedItem)) Then
        XPMenuClass.KillAllMenus
        
        HandleClick XPMenuClass.GetMenuName(), CInt(selectedItem), XPMenuClass.GetItemText(CInt(selectedItem))
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim getHilight As Long
    getHilight = XPMenuClass.GetHilightedItem(y)
    
    If getHilight = XPMenuClass.GetHilightNum Then Exit Sub
    XPMenuClass.setHilightedItem CInt(getHilight)

End Sub


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    upY = y
End Sub


Private Sub tmrActive_Timer()
    Dim frm As Form
    
    For Each frm In Forms
        If frm.Tag = "XPMenu" And GetActiveWindow() = frm.hwnd Then Exit Sub
    Next frm
    
    XPMenuClass.KillPopupMenus
    XPMenuClass.UnloadMenu
End Sub


Private Sub tmrHover_Timer()
    Dim pt As POINTAPI
    GetCursorPos pt
    
    Dim hw As Long
    hw = WindowFromPoint(pt.x, pt.y)
    
    If hw <> Me.hwnd Then
        If XPMenuClass.PopupShown() = False Then
            XPMenuClass.setHilightedItem -1
            'XPMenuClass.DrawMenu
        End If
    End If
End Sub


