VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Login"
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   2100
      TabIndex        =   5
      Tag             =   "Cancel"
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   495
      TabIndex        =   4
      Tag             =   "OK"
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1305
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   1305
      TabIndex        =   2
      Top             =   135
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   248
      Index           =   1
      Left            =   105
      TabIndex        =   1
      Tag             =   "&Password:"
      Top             =   540
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   248
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Tag             =   "&User Name:"
      Top             =   150
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long


Public OK As Boolean
Private Sub cmdCancel_Click()
    OK = False
    Me.Hide
End Sub


Private Sub cmdOK_Click()
Dim oldUs As String
    'ToDo: create test for correct password
    'check for correct password
    With DEnv1.rstblUsers
        If .State = 0 Then
        .Open
        End If
        .MoveFirst
        On Error Resume Next
        Do Until .Fields("Username") = txtUserName.Text
            On Error Resume Next
            .MoveNext
            If .EOF = True Then
                .MovePrevious
                oldUs = txtUserName.Text
                txtUserName.Text = .Fields("username")
            End If
        Loop
    
        If .Fields("password") = txtPassword.Text Then
            If .Fields("password") = "jubill" Then
                pwJubill = True
                MsgBox "Sorry for the inconvenience, but your password needs to be updated", , "Update Password."
            End If
            
            g_strUser = txtUserName.Text
            If .Fields("status") = 0 Then
                MsgBox "Your account has been deactivated.", , "Sorry"
                OK = False
                Unload Me
            Else
                
                With DEnv1.rstblUserLog
                    .Open
                    .addnew
                    .Fields("Username") = g_strUser
                    .Fields("LoggedIn") = "Yes"
                    .Fields("LoginDT") = Now
                    .Update
                    .Close
                End With
                OK = True
                Me.Hide
            End If
            .Close
        Else
            MsgBox "Invalid Password!", , "Login"
            
            If txtUserName.Text = "" Then
                txtUserName.SetFocus
            Else
                txtPassword.SetFocus
            End If
            txtPassword.Text = ""
            .Close
        End If
        If oldUs <> "" Then
            txtUserName.Text = oldUs
        End If
    End With
End Sub

