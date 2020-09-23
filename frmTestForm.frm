VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTestForm 
   Caption         =   "Test XP Menu"
   ClientHeight    =   1290
   ClientLeft      =   5715
   ClientTop       =   2445
   ClientWidth     =   2520
   Icon            =   "frmTestForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1290
   ScaleWidth      =   2520
   Begin VB.CommandButton Command1 
      Caption         =   "See Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   570
      TabIndex        =   0
      Top             =   420
      Width           =   1305
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1485
      Top             =   570
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestForm.frx":2AFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestForm.frx":2E94
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestForm.frx":322E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTestForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public XPMenu As New clsXPMenu
Public XPMenu2 As New clsXPMenu
Public XPM_EFNet As New clsXPMenu
Public XPM_DALNet As New clsXPMenu

Private Type POINTAPI
        x As Long
        y As Long
End Type
Private Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long

Private Sub Command1_Click()
    
    'Set XPM_DALNet = New clsXPMenu
    XPM_DALNet.Init "DALNet"
    XPM_DALNet.AddItem 0, "Server1 (blah blah blah)", False, False
    XPM_DALNet.AddItem 0, "Server2 (asdfasdf:5636)", False, False
    XPM_DALNet.AddItem 0, "Server3 (dalnet)", False, False
    XPM_DALNet.AddItem 0, "", False, True
    XPM_DALNet.AddItem 0, "Random DALNet Server", False, False
    
    'Set XPM_EFNet = New clsXPMenu
    XPM_EFNet.Init "EFNet"
    XPM_EFNet.AddItem 0, "Prison (irc.prison.net)", False, False
    XPM_EFNet.AddItem 0, "Lagged (irc.lagged.org)", False, False
    XPM_EFNet.AddItem 0, "Another one.... (unknown)", False, False
    XPM_EFNet.AddItem 0, "", False, True
    XPM_EFNet.AddItem 0, "Random EFNet Server", False, False
       
    'Set XPMenu2 = New clsXPMenu
    XPMenu2.Init "Servers"
    XPMenu2.AddItem 0, "DALNet", True, False, XPM_DALNet
    XPMenu2.AddItem 0, "EFNet", True, False, XPM_EFNet
    
    'Set XPMenu = New clsXPMenu
    XPMenu.Init "Connect", ImageList1
    XPMenu.AddItem 0, "New Server", True, False, XPMenu2
    XPMenu.AddItem 0, "", False, True
    XPMenu.AddItem 1, "Connect", False, False
    XPMenu.AddItem 2, "Disconnect", False, False
    XPMenu.AddItem 0, "", False, True
    XPMenu.AddItem 3, "Change Profile", False, False
    XPMenu.AddItem 0, "", False, True
    XPMenu.AddItem 0, "Exit", False, False
    
    Dim pos As POINTAPI
    GetCursorPos pos
        
    XPMenu.ShowMenu pos.x, pos.y
    
End Sub


Private Sub Command2_Click()

End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'ss
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim frm As Form
    For Each frm In Forms
        Unload frm
    Next
    
    End
End Sub


