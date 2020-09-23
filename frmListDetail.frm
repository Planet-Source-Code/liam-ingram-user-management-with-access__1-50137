VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "FLASH.OCX"
Begin VB.Form frmListDetail 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FrmListDetail"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10530
   Icon            =   "frmListDetail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   548
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   702
   ShowInTaskbar   =   0   'False
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   1320
      Left            =   120
      TabIndex        =   60
      Top             =   15
      Width           =   10335
      _cx             =   18230
      _cy             =   2328
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Users"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin TabDlg.SSTab TabMain 
      Height          =   6765
      Left            =   5400
      TabIndex        =   14
      Top             =   1335
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   11933
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   441
      BackColor       =   16777215
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tab1"
      TabPicture(0)   =   "frmListDetail.frx":2AFA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "pctPrepare"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "pctUser"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Tab2"
      TabPicture(1)   =   "frmListDetail.frx":2B16
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "pctUser2NoPerm"
      Tab(1).Control(1)=   "pctUser2"
      Tab(1).Control(2)=   "Label2"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Tab3"
      TabPicture(2)   =   "frmListDetail.frx":2B32
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "PctUser3NoPerm"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Tab4"
      TabPicture(3)   =   "frmListDetail.frx":2B4E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "pctUser4NoPerm"
      Tab(3).Control(1)=   "pctUser4"
      Tab(3).Control(2)=   "Label4"
      Tab(3).ControlCount=   3
      Begin VB.PictureBox pctUser2NoPerm 
         BackColor       =   &H00FFFFFF&
         Height          =   6255
         Left            =   -74880
         ScaleHeight     =   6195
         ScaleWidth      =   4755
         TabIndex        =   67
         Top             =   360
         Width           =   4815
         Begin VB.CommandButton pctUser2cmdCancel2 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   2520
            TabIndex        =   73
            Top             =   5640
            Width           =   975
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Update"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3600
            TabIndex        =   72
            Top             =   5640
            Width           =   975
         End
         Begin VB.Label pctUser2LblNoPerm 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            Caption         =   "Sorry, you do not have the appropriate permission to view these details."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   1215
            Left            =   120
            TabIndex        =   71
            Top             =   2520
            Width           =   4575
         End
      End
      Begin VB.PictureBox pctUser4NoPerm 
         BackColor       =   &H80000009&
         Height          =   6255
         Left            =   -74880
         ScaleHeight     =   6195
         ScaleWidth      =   4755
         TabIndex        =   66
         Top             =   360
         Width           =   4815
         Begin VB.CommandButton Command5 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   2520
            TabIndex        =   70
            Top             =   5640
            Width           =   975
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Update"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3600
            TabIndex        =   69
            Top             =   5640
            Width           =   975
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            Caption         =   "Sorry, you do not have the appropriate permission to view these details."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   1215
            Left            =   120
            TabIndex        =   68
            Top             =   2520
            Width           =   4575
         End
      End
      Begin VB.PictureBox PctUser3NoPerm 
         BackColor       =   &H80000009&
         Height          =   6255
         Left            =   -74880
         ScaleHeight     =   6195
         ScaleWidth      =   4755
         TabIndex        =   62
         Top             =   360
         Width           =   4815
         Begin VB.CommandButton pctUser3cmdUpdate 
            Caption         =   "Update"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3600
            TabIndex        =   65
            Top             =   5640
            Width           =   975
         End
         Begin VB.CommandButton pctUser3cmdCancel 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   2520
            TabIndex        =   64
            Top             =   5640
            Width           =   975
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            Caption         =   "Sorry, you do not have the appropriate permission to view these details."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   1215
            Left            =   120
            TabIndex        =   63
            Top             =   2520
            Width           =   4575
         End
      End
      Begin VB.PictureBox pctUser4 
         BackColor       =   &H00FFFFFF&
         Height          =   6255
         Left            =   -74880
         ScaleHeight     =   6195
         ScaleWidth      =   4755
         TabIndex        =   52
         Top             =   360
         Width           =   4815
         Begin VB.CommandButton pctUser4CmdUpdate 
            Caption         =   "Update"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3600
            TabIndex        =   59
            Top             =   5640
            Width           =   975
         End
         Begin VB.CommandButton pctUser4CmdCancel 
            Cancel          =   -1  'True
            Caption         =   "Cancel"
            Height          =   255
            Left            =   2520
            TabIndex        =   58
            Top             =   5640
            Width           =   975
         End
         Begin VB.Frame pctUser4FraActions 
            BackColor       =   &H00FFFFFF&
            Caption         =   "User Actions"
            Height          =   1215
            Left            =   120
            TabIndex        =   54
            Top             =   1080
            Width           =   4455
            Begin VB.CommandButton pctUser4CmdAddNew 
               Caption         =   "Add new User"
               Height          =   615
               Left            =   120
               TabIndex        =   55
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label pctUser4lblAddnew 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "This action allows you to create a new user, please ensure you have all nessacary details before beginning"
               Height          =   615
               Left            =   1560
               TabIndex        =   56
               Top             =   360
               Width           =   2775
            End
         End
         Begin VB.Label pctUser4lblActions 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "pctUser4lblActions"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   600
            TabIndex        =   53
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.PictureBox pctUser2 
         BackColor       =   &H00FFFFFF&
         Height          =   6255
         Left            =   -74880
         ScaleHeight     =   6195
         ScaleWidth      =   4755
         TabIndex        =   47
         Top             =   360
         Width           =   4815
         Begin VB.CommandButton pctUser2CmdPerm 
            Caption         =   "Apply"
            Height          =   255
            Left            =   3600
            TabIndex        =   61
            Top             =   5640
            Width           =   975
         End
         Begin VB.CommandButton pctUser2cmdCancel 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   2520
            TabIndex        =   57
            Top             =   5640
            Width           =   975
         End
         Begin MSFlexGridLib.MSFlexGrid pctUser2Flexgrid 
            Height          =   4575
            Left            =   240
            TabIndex        =   49
            Top             =   960
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   8070
            _Version        =   393216
            Cols            =   3
            FixedCols       =   2
            TextStyleFixed  =   1
            HighLight       =   2
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin VB.Label pctUser2LblUserPerms 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "pctUser2LblUserPerms"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   495
            Left            =   0
            TabIndex        =   48
            Top             =   240
            Width           =   4815
         End
      End
      Begin VB.PictureBox pctUser 
         BackColor       =   &H00FFFFFF&
         Height          =   6255
         Left            =   120
         ScaleHeight     =   6195
         ScaleWidth      =   4755
         TabIndex        =   18
         Top             =   360
         Width           =   4815
         Begin VB.CommandButton pctUserCmdChangePassword 
            Caption         =   "Change"
            Height          =   255
            Left            =   3600
            TabIndex        =   46
            Top             =   1680
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton pctUserCmdCancel 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   2520
            TabIndex        =   45
            Top             =   5640
            Width           =   975
         End
         Begin VB.CommandButton pctUserCmdUpdate 
            Caption         =   "Update"
            Height          =   255
            Left            =   3600
            TabIndex        =   44
            Top             =   5640
            Width           =   975
         End
         Begin VB.TextBox pctUserTxtMobile 
            Height          =   285
            Left            =   1920
            TabIndex        =   43
            Top             =   5280
            Width           =   2655
         End
         Begin VB.TextBox pctUserTxtWorkEmail 
            Height          =   285
            Left            =   1920
            TabIndex        =   41
            Top             =   4920
            Width           =   2655
         End
         Begin VB.TextBox pctUserTxtWorkFax 
            Height          =   285
            Left            =   1920
            TabIndex        =   39
            Top             =   4560
            Width           =   2655
         End
         Begin VB.TextBox pctUserTxtWorkTel 
            Height          =   285
            Left            =   1920
            TabIndex        =   37
            Top             =   4200
            Width           =   2655
         End
         Begin VB.TextBox pctUserTxtHomeEmail 
            Height          =   285
            Left            =   1920
            TabIndex        =   35
            Top             =   3840
            Width           =   2655
         End
         Begin VB.TextBox pctUserTxtHomeTel 
            Height          =   285
            Left            =   1920
            TabIndex        =   33
            Top             =   3480
            Width           =   2655
         End
         Begin VB.TextBox pctUserTxtJobTitle 
            Height          =   285
            Left            =   1920
            TabIndex        =   31
            Top             =   3120
            Width           =   2655
         End
         Begin VB.TextBox pctUserTxtFullName 
            Height          =   285
            Left            =   1920
            TabIndex        =   29
            Top             =   2760
            Width           =   2655
         End
         Begin VB.CommandButton pctUserCmdActive 
            Caption         =   "Deactivate"
            Height          =   255
            Left            =   3600
            TabIndex        =   27
            Top             =   2040
            Width           =   975
         End
         Begin VB.TextBox pctUserTxtPassword 
            Enabled         =   0   'False
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1920
            PasswordChar    =   "*"
            TabIndex        =   24
            Top             =   1680
            Width           =   2655
         End
         Begin VB.TextBox pctUserTxtUsername 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1920
            TabIndex        =   22
            Top             =   1320
            Width           =   2655
         End
         Begin VB.Label pctUserLblMobile 
            BackStyle       =   0  'Transparent
            Caption         =   "Mobile:"
            Height          =   255
            Left            =   360
            TabIndex        =   42
            Top             =   5280
            Width           =   975
         End
         Begin VB.Label pctUserLblWorkEmail 
            BackStyle       =   0  'Transparent
            Caption         =   "Work Email:"
            Height          =   255
            Left            =   360
            TabIndex        =   40
            Top             =   4920
            Width           =   1095
         End
         Begin VB.Label pctUserLblWorkFax 
            BackStyle       =   0  'Transparent
            Caption         =   "Work Fax:"
            Height          =   255
            Left            =   360
            TabIndex        =   38
            Top             =   4560
            Width           =   1095
         End
         Begin VB.Label pctUserLblWorkTel 
            BackStyle       =   0  'Transparent
            Caption         =   "Work Tel:"
            Height          =   255
            Left            =   360
            TabIndex        =   36
            Top             =   4200
            Width           =   1095
         End
         Begin VB.Label pctUserLblHomeEmail 
            BackStyle       =   0  'Transparent
            Caption         =   "Home Email:"
            Height          =   255
            Left            =   360
            TabIndex        =   34
            Top             =   3840
            Width           =   975
         End
         Begin VB.Label pctUserLblHomeTel 
            BackStyle       =   0  'Transparent
            Caption         =   "Home Tel:"
            Height          =   255
            Left            =   360
            TabIndex        =   32
            Top             =   3480
            Width           =   975
         End
         Begin VB.Label pctUserLblJobTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Job Title:"
            Height          =   255
            Left            =   360
            TabIndex        =   30
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Label pctUserLblFullName 
            BackStyle       =   0  'Transparent
            Caption         =   "Full Name:"
            Height          =   255
            Left            =   360
            TabIndex        =   28
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label pctUserlblEmployeeInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Info"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label pctUserLblActive 
            BackStyle       =   0  'Transparent
            Caption         =   "pctUserLblActive"
            Height          =   255
            Left            =   360
            TabIndex        =   25
            Top             =   2040
            Width           =   3255
         End
         Begin VB.Label pctUserLblPassword 
            BackStyle       =   0  'Transparent
            Caption         =   "Password:"
            Height          =   255
            Left            =   360
            TabIndex        =   23
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label pctUserLblUsername 
            BackStyle       =   0  'Transparent
            Caption         =   "Username:"
            Height          =   255
            Left            =   360
            TabIndex        =   21
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label pctUserLblUserAccount 
            BackStyle       =   0  'Transparent
            Caption         =   "User Account"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label pctUserLblUserDetails 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "pctUserLblUserDetails"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   615
            Left            =   240
            TabIndex        =   19
            Top             =   240
            Width           =   4335
         End
      End
      Begin VB.PictureBox pctPrepare 
         BackColor       =   &H00FFFFFF&
         Height          =   6255
         Left            =   120
         ScaleHeight     =   6195
         ScaleWidth      =   4755
         TabIndex        =   15
         Top             =   360
         Width           =   4815
         Begin VB.Label lblPrepare 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Please double click a user on the left to view and edit details."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   735
            Left            =   120
            TabIndex        =   16
            Top             =   2760
            Width           =   4575
         End
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "You do not have permission to view these details!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   735
         Left            =   -74760
         TabIndex        =   51
         Top             =   3120
         Width           =   4695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "You do not have permission to view these details!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   735
         Left            =   -74760
         TabIndex        =   50
         Top             =   3120
         Width           =   4695
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salesmen"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Customers"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame fraSort 
      BackColor       =   &H00FFFFFF&
      Height          =   1280
      Left            =   120
      TabIndex        =   1
      Top             =   6840
      Width           =   5175
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sorting"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   985
         Left            =   120
         TabIndex        =   5
         Top             =   175
         Width           =   2820
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFFFFF&
            Height          =   420
            Left            =   120
            TabIndex        =   8
            Top             =   444
            Width           =   2615
            Begin VB.OptionButton Option4 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Descending"
               Height          =   255
               Left            =   1320
               TabIndex        =   10
               Top             =   120
               Width           =   1215
            End
            Begin VB.OptionButton Option3 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Ascending"
               Height          =   255
               Left            =   80
               TabIndex        =   9
               Top             =   120
               Width           =   1095
            End
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Option1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   200
            TabIndex        =   7
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Option2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            TabIndex        =   6
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.TextBox txtsearch 
         Height          =   285
         Left            =   3120
         TabIndex        =   2
         Top             =   420
         Width           =   1815
      End
      Begin VB.Label lblRecords 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "There are currently 000 records"
         Height          =   450
         Left            =   3360
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblSearch 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   3
         Top             =   195
         Width           =   2055
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Flexgrid 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   9763
      _Version        =   393216
      FixedCols       =   0
      BackColorSel    =   -2147483632
      BackColorBkg    =   16777215
      GridColor       =   8454143
      TextStyleFixed  =   1
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   3240
      TabIndex        =   13
      Top             =   240
      Width           =   6975
   End
End
Attribute VB_Name = "frmListDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oldPw, newPw, newPw2 As String
Dim oldPwEnt, newPwEnt, newPw2Ent, justLoaded, recAltered, passincorrect, passupdated, userPermOK As Boolean
Dim intClick, irPermCount As Integer
Public addnew As Boolean

Private Sub Command1_Click()
    
    PopulateFrm "Customers", "option1", True, ""
    populateFlex
    Label1.Caption = "Customers"

End Sub

Private Sub Command2_Click()
    
    PopulateFrm "Salesmen", "option1", True, ""
    populateFlex
    Label1.Caption = "Salesmen"
    
End Sub

Private Sub Command3_Click()
    
    PopulateFrm "Users", "option1", True, ""
    populateFlex
    Label1.Caption = "User Management"
    
End Sub

Private Sub Command5_Click()
pctUserCmdCancel_Click
End Sub



Private Sub Flexgrid_DblClick()
    populatetabs True, latestentity, Flexgrid.Text
End Sub

Private Sub Form_Activate()
    
    If EntityMnu <> "" Then
        PopulateFrm "Users", "option1", True, ""
        populateFlex
        Label1.Caption = "User Management"
    End If
    
    If pwJubill = True Then
        PopulateFrm "Users", "option1", True, ""
        populateFlex
        Label1.Caption = "User Management"
        Flexgrid.Row = 0
        Do Until Flexgrid.Text = g_strUser
            Flexgrid.Row = Flexgrid.Row + 1
        Loop
        populatetabs True, latestentity, Flexgrid.Text
        intClick = 1
        pctUserCmdChangePassword_Click
    End If
    
    checkPermission 1
    If permOk = True Then
        userPermOK = True
    Else
        userPermOK = False
    End If
    TabMain.Tab = 0
    
End Sub

Private Sub Form_Load()
    SetMenu Me.hWnd, , , lv_MDIchildForm_NoMenus
    Me.Height = 8535
    Me.Width = 10665
    Me.Top = 250
    Me.Left = 250
    justLoaded = True
    Option1.Value = True
    Option3.Value = True
    populatetabs False, "", ""
    ShockwaveFlash1.Play
    
End Sub

Private Function populateFlex()

    With Flexgrid
        .Visible = False
        .Clear
        .Rows = iRcount + 1
        .Cols = 2
        .ColWidth(0) = 1200
        .ColWidth(1) = 3350
        .TextMatrix(0, 0) = option1Tag
        .TextMatrix(0, 1) = option2Tag
        On Error Resume Next
        .Row = 1
        .Col = 0
        .RowSel = iRcount
        .ColSel = 1
        
        Select Case latestentity
            
            Case "Customers"
                With DEnv1.rsClientTable
                    On Error Resume Next
                    .Open
                    On Error Resume Next
                    .MoveFirst
                End With
                .Clip = DEnv1.rsClientTable.GetString
            
            Case "Salesmen"
                With DEnv1.rstblAgents
                    On Error Resume Next
                    .Open
                    On Error Resume Next
                    .MoveFirst
                End With
                .Clip = DEnv1.rstblAgents.GetString
        
            Case "Users"
                With DEnv1.rstblUsers
                    On Error Resume Next
                    .Open
                    On Error Resume Next
                    .MoveFirst
                End With
                ShockwaveFlash1.Movie = App.Path & "\images\usermanagement2.swf"
                .Clip = DEnv1.rstblUsers.GetString
               ' frmListDetail.Caption = "User Management - Welcome " & g_strUser
                
        End Select
        .Row = 1
        .Col = 0
        .ColSel = 0
        .ColSel = 1
        
    End With

    Option1.Caption = option1Tag
    Option2.Caption = option2Tag
    
    lblRecords.Caption = "There are currently " & iRcount & " records."
    Flexgrid.Visible = True
    Flexgrid.ColAlignment(0) = 1
    Flexgrid.ColAlignment(1) = 1

    If Option1.Value = True Then
        lblSearch.Caption = "Search by " & option1Tag
        If Option3.Value = True Then
            sortOnlyNeeded "option1", True
        Else
            sortOnlyNeeded "option1", False
        End If
    Else
        lblSearch.Caption = "Search by " & option2Tag
        If Option3.Value = True Then
            sortOnlyNeeded "option2", True
        Else
            sortOnlyNeeded "option2", False
        End If
    End If
    

'ShockwaveFlash1.Play
End Function

Private Sub Form_Unload(Cancel As Integer)

If pwJubill = True Then
    MsgBox "Sorry please complete updating your password", , "Update password"
    Cancel = -1
End If

End Sub



Private Sub Option1_Click()
    lblSearch.Caption = "Search by " & option1Tag
    If justLoaded = False Then
        If Option3.Value = True Then
            sortOnlyNeeded "option1", True
        Else
            sortOnlyNeeded "option1", False
        End If
    
    End If

End Sub

Private Sub Option2_Click()
    lblSearch.Caption = "Search by " & option2Tag
    If Option3.Value = True Then
        sortOnlyNeeded "option2", True
    Else
        sortOnlyNeeded "option2", False
    End If
    
End Sub

Private Sub Option3_Click()
    
    If justLoaded = False Then
        If Option1.Value = True Then
            sortOnlyNeeded "option1", True
        Else
            sortOnlyNeeded "option2", True
        End If
    
    End If
    justLoaded = False
End Sub

Private Sub Option4_Click()
    
    If Option1.Value = True Then
        sortOnlyNeeded "option1", False
    Else
        sortOnlyNeeded "option2", False
    End If
    
End Sub


Public Function sortOnlyNeeded(sortBy As String, Ascending As Boolean)

    If sortBy = "option1" Then
        Flexgrid.Col = 0
        If Ascending = True Then
            Flexgrid.Sort = flexSortStringAscending
        Else
            Flexgrid.Sort = flexSortStringDescending
        End If
    Else
        Flexgrid.Col = 1
        If Ascending = True Then
            Flexgrid.Sort = flexSortStringAscending
        Else
            Flexgrid.Sort = flexSortStringDescending
        End If
    End If
   
End Function


Private Sub pctUser2cmdCancel_Click()
    If recAltered = False Then
        populatetabs False, "", ""
    Else
        recAltered = False
        TabMain.Tab = 0
        Select Case MsgBox("Do you wish to save these changes?", vbYesNoCancel + vbQuestion, "Save Changes?")

        Case vbYes

            With DEnv1.rstblUsers
                .Fields("Fullname") = pctUserTxtFullName.Text
                .Fields("jobdescription") = pctUserTxtJobTitle.Text
                .Fields("hometel") = pctUserTxtHomeTel.Text
                .Fields("homeemail") = pctUserTxtHomeEmail.Text
                .Fields("Worktel") = pctUserTxtWorkTel.Text
                .Fields("workfax") = pctUserTxtWorkFax.Text
                .Fields("workemail") = pctUserTxtWorkEmail.Text
                .Fields("mobile") = pctUserTxtMobile.Text
                .Update
            End With
            populatetabs False, "", ""

        Case vbNo
            populatetabs False, "", ""

        Case vbCancel
            recAltered = True

        End Select
    End If
End Sub

Private Sub pctUser2cmdCancel2_Click()
pctUserCmdCancel_Click
End Sub

Private Sub pctUser2CmdPerm_Click()
Dim mbResult As String
Dim currRow
    With pctUser2Flexgrid
        currRow = .Row
        If .Text = "Yes" Then
            mbResult = MsgBox("Are you sure you wish to remove this permission from this account?", vbYesNo + vbQuestion, "Remove permission?")
            If mbResult = vbYes Then
                .Text = "No"
            End If
        Else
            mbResult = MsgBox("Are you sure you wish to apply this permission for this account?", vbYesNo + vbQuestion, "Apply permission?")
            If mbResult = vbYes Then
                .Text = "Yes"
            End If
        End If
    End With
    pctUser2Flexgrid_Click
    'pctUser2CmdUpdate.enabled = True

    Dim foundRec As Boolean
        
        With DEnv1.rstblUserPermissions
            .MoveFirst
            Do Until foundRec = True
            If .Fields("username") = Flexgrid.Text Then
                pctUser2Flexgrid.Col = 0
                pctUser2Flexgrid.Row = currRow
                If .Fields("permissionid") = pctUser2Flexgrid.Text Then
                    foundRec = True
                Else
                    .MoveNext
                End If
            Else
                .MoveNext
            End If
                
            Loop
            
            foundRec = False
            pctUser2Flexgrid.Col = 2
            .Fields("Value") = pctUser2Flexgrid.Text
           ' On Error Resume Next
            .Update
        End With
    

End Sub

'Private Sub pctUser2CmdUpdate_Click()


'End Sub

Private Sub pctUser2Flexgrid_Click()
If pctUser2Flexgrid.Text <> "Yes" Then
    pctUser2CmdPerm.Caption = "Apply"
Else
    pctUser2CmdPerm.Caption = "Remove"
End If

End Sub

Private Sub pctUser3cmdCancel_Click()
pctUserCmdCancel_Click
End Sub

Private Sub pctUser4CmdAddNew_Click()
    
    'Dim usName As String
    TabMain.Tab = 0
    
    pctUserCmdChangePassword.Visible = False
    pctUserLblPassword.FontBold = False
    pctUserTxtPassword.PasswordChar = ""
    pctUserCmdActive.enabled = False
    pctUserCmdChangePassword.enabled = False
    pctUserLblUserDetails.Caption = "new user details"
    pctUserTxtUsername.enabled = True
    pctUserTxtPassword.enabled = True
    pctUserTxtFullName.enabled = True
    pctUserTxtJobTitle.enabled = True
    pctUserTxtHomeTel.enabled = True
    pctUserTxtHomeEmail.enabled = True
    pctUserTxtMobile.enabled = True
    pctUserTxtWorkTel.enabled = True
    pctUserTxtWorkFax.enabled = True
    pctUserTxtWorkEmail.enabled = True
    
    pctUserTxtUsername.Text = ""
    pctUserTxtPassword.Text = ""
    pctUserTxtFullName.Text = ""
    pctUserTxtJobTitle.Text = ""
    pctUserTxtHomeTel.Text = ""
    pctUserTxtHomeEmail.Text = ""
    pctUserTxtMobile.Text = ""
    pctUserTxtWorkTel.Text = ""
    pctUserTxtWorkFax.Text = ""
    pctUserTxtWorkEmail.Text = ""


    pctUserLblFullName.enabled = True
    pctUserLblJobTitle.enabled = True
    pctUserLblHomeTel.enabled = True
    pctUserLblHomeEmail.enabled = True
    pctUserLblMobile.enabled = True
    pctUserLblWorkTel.enabled = True
    pctUserLblWorkFax.enabled = True
    pctUserLblWorkEmail.enabled = True
    pctUserLblUsername.enabled = True
    pctUserLblActive.enabled = True
        
    pctUserCmdActive.enabled = False
    pctUserCmdCancel.enabled = True
    pctUserCmdUpdate.enabled = True
    pctUserCmdChangePassword.enabled = False
    pctUserCmdChangePassword.Default = False
    pctUserLblActive.enabled = False
    pctUserLblPassword.Caption = "Password:"
    pctUserLblActive.Caption = "Account does not exist"
    addnew = True
    passchange = True
End Sub

Private Sub pctUser4CmdCancel_Click()
    If recAltered = False Then
        populatetabs False, "", ""
    Else
        recAltered = False
        TabMain.Tab = 0
        Select Case MsgBox("Do you wish to save these changes?", vbYesNoCancel + vbQuestion, "Save Changes?")

        Case vbYes

            With DEnv1.rstblUsers
                .Fields("Fullname") = pctUserTxtFullName.Text
                .Fields("jobdescription") = pctUserTxtJobTitle.Text
                .Fields("hometel") = pctUserTxtHomeTel.Text
                .Fields("homeemail") = pctUserTxtHomeEmail.Text
                .Fields("Worktel") = pctUserTxtWorkTel.Text
                .Fields("workfax") = pctUserTxtWorkFax.Text
                .Fields("workemail") = pctUserTxtWorkEmail.Text
                .Fields("mobile") = pctUserTxtMobile.Text
                .Update
            End With
            populatetabs False, "", ""

        Case vbNo
            populatetabs False, "", ""

        Case vbCancel
            recAltered = True

        End Select
    End If
End Sub



Private Sub pctUserCmdActive_Click()
    
    If DEnv1.rstblUsers.Fields("Status") = 1 Then
        If MsgBox("Are you sure you wish to deactivate this account?", vbYesNo + vbQuestion, "Deactivate?") = vbYes Then
            DEnv1.rstblUsers.Fields("Status") = 0
            DEnv1.rstblUsers.Update
            populatetabs True, latestentity, Flexgrid.Text
        End If
    Else
        If MsgBox("Are you sure you wish to activate this account?", vbYesNo + vbQuestion, "Activate?") = vbYes Then
            DEnv1.rstblUsers.Fields("Status") = 1
            DEnv1.rstblUsers.Update
            populatetabs True, latestentity, Flexgrid.Text
        End If
    End If
    
End Sub

Private Sub pctUserCmdCancel_Click()
    If recAltered = False Then
        populatetabs False, "", ""
    Else
        recAltered = False
        Select Case MsgBox("Do you wish to save these changes?", vbYesNoCancel + vbQuestion, "Save Changes?")

        Case vbYes

            With DEnv1.rstblUsers
                If addnew = True Then
                    .addnew
                    .Fields("username") = pctUserTxtUsername.Text
                    .Fields("password") = pctUserTxtPassword.Text
                    .Fields("Status") = 1
                End If
                .Fields("Fullname") = pctUserTxtFullName.Text
                .Fields("jobdescription") = pctUserTxtJobTitle.Text
                .Fields("hometel") = pctUserTxtHomeTel.Text
                .Fields("homeemail") = pctUserTxtHomeEmail.Text
                .Fields("Worktel") = pctUserTxtWorkTel.Text
                .Fields("workfax") = pctUserTxtWorkFax.Text
                .Fields("workemail") = pctUserTxtWorkEmail.Text
                .Fields("mobile") = pctUserTxtMobile.Text
                .Update
                addnew = False
                passchange = False
            End With
            populatetabs False, "", ""

        Case vbNo
            populatetabs False, "", ""

        Case vbCancel
            recAltered = True

        End Select
    End If
End Sub

Private Sub pctUserCmdChangePassword_Click()
    
    If userPermOK = True Then
        If intClick <> 4 Then
            If intClick = 3 Then
                intClick = 3
            Else
                If intClick = 2 Then
                    intClick = 2
                Else
                    If intClick = 0 Then
                        intClick = 1
                    End If
                End If
            End If
        End If
    End If
                    
                
        
    
    
    
    If intClick = 0 Then
        intClick = intClick + 1
        pctUserCmdChangePassword.Default = True
    Else
        If pctUserTxtPassword.Text <> "" Then
            intClick = intClick + 1
        Else
            MsgBox "Please input password", , "Change Password."
            pctUserTxtPassword.SetFocus
        End If
    End If
        
    Select Case intClick
    
        Case 1
            pctUserTxtPassword.enabled = True
            pctUserCmdChangePassword.Caption = "Done"
            pctUserTxtPassword.SetFocus
            pctUserLblPassword.Caption = "Input Current:"
            pctUserLblPassword.FontBold = True
            pctUserTxtPassword.Text = ""
        
            pctUserTxtFullName.enabled = False
            pctUserTxtJobTitle.enabled = False
            pctUserTxtHomeTel.enabled = False
            pctUserTxtHomeEmail.enabled = False
            pctUserTxtMobile.enabled = False
            pctUserTxtWorkTel.enabled = False
            pctUserTxtWorkFax.enabled = False
            pctUserTxtWorkEmail.enabled = False
        
            pctUserLblFullName.enabled = False
            pctUserLblJobTitle.enabled = False
            pctUserLblHomeTel.enabled = False
            pctUserLblHomeEmail.enabled = False
            pctUserLblMobile.enabled = False
            pctUserLblWorkTel.enabled = False
            pctUserLblWorkFax.enabled = False
            pctUserLblWorkEmail.enabled = False
            pctUserLblUsername.enabled = False
            pctUserLblActive.enabled = False
            
            pctUserCmdActive.enabled = False
            pctUserCmdCancel.enabled = False
            pctUserCmdUpdate.enabled = False
            
            
        Case 2
            If userPermOK = True Then
                pctUserTxtPassword.enabled = True
                pctUserCmdChangePassword.Caption = "Done"
                pctUserTxtPassword.SetFocus
                pctUserLblPassword.Caption = "Input Current:"
                pctUserLblPassword.FontBold = True
                pctUserTxtPassword.Text = ""
        
                pctUserTxtFullName.enabled = False
                pctUserTxtJobTitle.enabled = False
                pctUserTxtHomeTel.enabled = False
                pctUserTxtHomeEmail.enabled = False
                pctUserTxtMobile.enabled = False
                pctUserTxtWorkTel.enabled = False
                pctUserTxtWorkFax.enabled = False
                pctUserTxtWorkEmail.enabled = False
        
                pctUserLblFullName.enabled = False
                pctUserLblJobTitle.enabled = False
                pctUserLblHomeTel.enabled = False
                pctUserLblHomeEmail.enabled = False
                pctUserLblMobile.enabled = False
                pctUserLblWorkTel.enabled = False
                pctUserLblWorkFax.enabled = False
                pctUserLblWorkEmail.enabled = False
                pctUserLblUsername.enabled = False
                pctUserLblActive.enabled = False
            
                pctUserCmdActive.enabled = False
                pctUserCmdCancel.enabled = False
                pctUserCmdUpdate.enabled = False
                pctUserCmdChangePassword.Default = True
            End If
            If pwJubill = True Then
                pctUserTxtPassword.enabled = True
                pctUserCmdChangePassword.Caption = "Done"
                pctUserTxtPassword.SetFocus
                pctUserLblPassword.Caption = "Input Current:"
                pctUserLblPassword.FontBold = True
                pctUserTxtPassword.Text = ""
        
                pctUserTxtFullName.enabled = False
                pctUserTxtJobTitle.enabled = False
                pctUserTxtHomeTel.enabled = False
                pctUserTxtHomeEmail.enabled = False
                pctUserTxtMobile.enabled = False
                pctUserTxtWorkTel.enabled = False
                pctUserTxtWorkFax.enabled = False
                pctUserTxtWorkEmail.enabled = False
        
                pctUserLblFullName.enabled = False
                pctUserLblJobTitle.enabled = False
                pctUserLblHomeTel.enabled = False
                pctUserLblHomeEmail.enabled = False
                pctUserLblMobile.enabled = False
                pctUserLblWorkTel.enabled = False
                pctUserLblWorkFax.enabled = False
                pctUserLblWorkEmail.enabled = False
                pctUserLblUsername.enabled = False
                pctUserLblActive.enabled = False
            
                pctUserCmdActive.enabled = False
                pctUserCmdCancel.enabled = False
                pctUserCmdUpdate.enabled = False
                pctUserCmdChangePassword.Default = True
            End If
                oldPw = pctUserTxtPassword.Text
                pctUserTxtPassword.SetFocus
                pctUserLblPassword.Caption = "Input new:"
                pctUserTxtPassword.Text = ""
    
        Case 3
            newPw = pctUserTxtPassword.Text
            pctUserTxtPassword.SetFocus
            pctUserLblPassword.Caption = "Again:"
            pctUserTxtPassword.Text = ""
            
        Case 4
            newPw2 = pctUserTxtPassword.Text
            If pwJubill = True Then
                oldPw = "jubill"
            End If
            If userPermOK = True Then
                oldPw = DEnv1.rstblUsers.Fields("password")
            End If
            If DEnv1.rstblUsers.Fields("password") <> oldPw Then
                MsgBox "Password change unsuccessfull!", , "Incorrect Password"
                passincorrect = True
            Else
                If newPw <> newPw2 Then
                    MsgBox "Password change unsuccessfull!", , "Incorrect Password"
                    passincorrect = True
                    
                Else
                    If pctUserTxtPassword.Text = g_strUser Then
                        MsgBox "Your password cannot be the same as your username!", , "Sorry!"
                        passincorrect = True
                    Else
                        If MsgBox("Are you sure you wish to change the password?", vbYesNo + vbQuestion, "Change Password?") = vbYes Then
                            DEnv1.rstblUsers.Fields("Password") = pctUserTxtPassword.Text
                            DEnv1.rstblUsers.Update
        
                            pctUserTxtPassword.enabled = False
                            pctUserCmdChangePassword.Caption = "Change"
                            pctUserLblPassword.FontBold = False
                            pctUserTxtFullName.enabled = True
                            pctUserTxtJobTitle.enabled = True
                            pctUserTxtHomeTel.enabled = True
                            pctUserTxtHomeEmail.enabled = True
                            pctUserTxtMobile.enabled = True
                            pctUserTxtWorkTel.enabled = True
                            pctUserTxtWorkFax.enabled = True
                            pctUserTxtWorkEmail.enabled = True
                            pctUserCmdActive.enabled = True
                            pctUserCmdCancel.enabled = True
                            pctUserCmdUpdate.enabled = True
                            pctUserCmdChangePassword.Default = False
                            
                            pctUserLblFullName.enabled = True
                            pctUserLblJobTitle.enabled = True
                            pctUserLblHomeTel.enabled = True
                            pctUserLblHomeEmail.enabled = True
                            pctUserLblMobile.enabled = True
                            pctUserLblWorkTel.enabled = True
                            pctUserLblWorkFax.enabled = True
                            pctUserLblWorkEmail.enabled = True
                            pctUserLblUsername.enabled = True
                            pctUserLblActive.enabled = True
                            
                            intClick = 0
                            pctUserTxtPassword.Text = DEnv1.rstblUsers.Fields("password")
                            pctUserLblPassword.Caption = "Password:"
                            If pwJubill = True Then
                            MsgBox "Thankyou, your password has been successfully updated", , "Password Updated!"
                            pwJubill = False
                            passupdated = True
                            Else
                            MsgBox "Password successfully changed", , "Password Changed!"
                        
                        End If
                    End If
                End If
            End If
        End If
    End Select

    If passincorrect = True Then
        pctUserTxtPassword.enabled = False
        pctUserCmdChangePassword.Caption = "Change"
        pctUserLblPassword.FontBold = False
        pctUserTxtFullName.enabled = True
        pctUserTxtJobTitle.enabled = True
        pctUserTxtHomeTel.enabled = True
        pctUserTxtHomeEmail.enabled = True
        pctUserTxtMobile.enabled = True
        pctUserTxtWorkTel.enabled = True
        pctUserTxtWorkFax.enabled = True
        pctUserTxtWorkEmail.enabled = True
        
        pctUserLblFullName.enabled = True
        pctUserLblJobTitle.enabled = True
        pctUserLblHomeTel.enabled = True
        pctUserLblHomeEmail.enabled = True
        pctUserLblMobile.enabled = True
        pctUserLblWorkTel.enabled = True
        pctUserLblWorkFax.enabled = True
        pctUserLblWorkEmail.enabled = True
        pctUserLblUsername.enabled = True
        pctUserLblActive.enabled = True
        
        pctUserCmdActive.enabled = True
        pctUserCmdCancel.enabled = True
        pctUserCmdUpdate.enabled = True
        pctUserCmdChangePassword.Default = False
        intClick = 0
        pctUserLblPassword.Caption = "Password:"
        pctUserTxtPassword.Text = DEnv1.rstblUsers.Fields("password")
        passincorrect = False
        If pwJubill = True Then
            intClick = 1
            pctUserCmdChangePassword_Click
        End If
    End If
    If passupdated = True Then
        passupdated = False
        Unload Me
    End If
End Sub

Private Sub pctUserCmdUpdate_Click()
    
    If Flexgrid.Text <> g_strUser Then
        If userPermOK = False Then
            MsgBox "Sorry you can not update someone elses account."
            populatetabs False, "", ""
            Exit Sub:
        End If
    End If
    If addnew = False Then
        If MsgBox("Are you sure you wish to update this account?", vbYesNo + vbQuestion, "Update?") = vbYes Then
            With DEnv1.rstblUsers
                .Fields("Fullname") = pctUserTxtFullName.Text
                .Fields("jobdescription") = pctUserTxtJobTitle.Text
                .Fields("hometel") = pctUserTxtHomeTel.Text
                .Fields("homeemail") = pctUserTxtHomeEmail.Text
                .Fields("Worktel") = pctUserTxtWorkTel.Text
                .Fields("workfax") = pctUserTxtWorkFax.Text
                .Fields("workemail") = pctUserTxtWorkEmail.Text
                .Fields("mobile") = pctUserTxtMobile.Text
                
                .Update
            End With
            populateFlex
            populatetabs False, "", ""
        End If
    Else
        If MsgBox("Are you sure you wish to add this account?", vbYesNo + vbQuestion, "Update?") = vbYes Then
            With DEnv1.rstblUsers
                .addnew
                .Fields("Username") = pctUserTxtUsername.Text
                .Fields("Password") = pctUserTxtPassword.Text
                .Fields("Fullname") = pctUserTxtFullName.Text
                .Fields("jobdescription") = pctUserTxtJobTitle.Text
                .Fields("hometel") = pctUserTxtHomeTel.Text
                .Fields("homeemail") = pctUserTxtHomeEmail.Text
                .Fields("Worktel") = pctUserTxtWorkTel.Text
                .Fields("workfax") = pctUserTxtWorkFax.Text
                .Fields("workemail") = pctUserTxtWorkEmail.Text
                .Fields("mobile") = pctUserTxtMobile.Text
                .Fields("Status") = 1
                .Update
                               
            End With
            With DEnv1.rstblPermissions
            On Error Resume Next
            .Open
            Dim recArray(1, 3)
            .MoveFirst
           ''''''''''''''''''''''''''''''''''''''''''''''''''''
            For I = 0 To 3
                recArray(0, I) = .Fields("PermissionID")
                recArray(1, I) = .Fields("PermissionName")
                .MoveNext
            Next I
            End With
            
            For I = 0 To 3
                With DEnv1.rstblUserPermissions
                    .addnew
                    .Fields("permissionid") = recArray(0, I)
                    .Fields("permission") = recArray(1, I)
                    .Fields("value") = "No"
                    .Fields("username") = pctUserTxtUsername.Text
                    .Update
                End With
            Next I
            ''''''''''''''''''''''''''''''''''''''''''''''''''''
            PopulateFrm "Users", "option1", True, ""
            populateFlex
            populatetabs False, "", ""
        End If
        addnew = False
        passchange = False
    End If
    
End Sub



Private Sub pctUserTxtFullName_click()
recAltered = True
End Sub

Private Sub pctUserTxtHomeEmail_click()
recAltered = True
End Sub

Private Sub pctUserTxtHomeTel_click()
recAltered = True
End Sub

Private Sub pctUserTxtJobTitle_click()
recAltered = True
End Sub

Private Sub pctUserTxtMobile_click()
recAltered = True
End Sub

Private Sub pctUserTxtPassword_click()
recAltered = True
End Sub

Private Sub pctUserTxtUsername_click()
recAltered = True
End Sub

Private Sub pctUserTxtWorkEmail_click()
recAltered = True
End Sub

Private Sub pctUserTxtWorkFax_click()
recAltered = True
End Sub

Private Sub pctUserTxtWorkTel_click()
recAltered = True
End Sub

Private Sub TabMain_Click(PreviousTab As Integer)
If passchange = True Then
    Select Case TabMain.Tab
        Case 1
            TabMain.Tab = 0
            If addnew = False Then
                pctUserTxtPassword.SetFocus
            End If
        Case 2
            TabMain.Tab = 0
            If addnew = False Then
                pctUserTxtPassword.SetFocus
            End If
        Case 3
            TabMain.Tab = 0
            If addnew = False Then
                pctUserTxtPassword.SetFocus
            End If
    End Select
Else
    On Error Resume Next
End If


End Sub

Private Function populatetabs(enabled As Boolean, tabEntity As String, recselector As String)
Dim I As Integer
    If enabled = False Then
        
        TabMain.enabled = False
        pctUser.Visible = False
        pctPrepare.Visible = True
        Flexgrid.enabled = True
        Flexgrid.ForeColor = &H80000008
        fraSort.enabled = True
        With TabMain
            .Tab = 0
            .Caption = "User Details"
            .Tab = 1
            .Caption = "Permissions"
            .Tab = 2
            .Caption = "Statistics"
            .Tab = 3
            .Caption = "Actions"
            .Tab = 0
        End With
    Else
        TabMain.enabled = True
        Flexgrid.enabled = False
        Flexgrid.ForeColor = &H8000000E
        fraSort.enabled = False
        pctUserTxtUsername.enabled = False
        pctUserTxtPassword.enabled = False
        Select Case tabEntity
    
        Case "Users"
            With TabMain
                .Tab = 0
                .Caption = "User Details"
                .Tab = 1
                .Caption = "Permissions"
                .Tab = 2
                .Caption = "Statistics"
                .Tab = 3
                .Caption = "Actions"
                .Tab = 0
            End With
            pctUser.Visible = False
            With DEnv1.rstblUsers
                .MoveFirst
                Do Until .Fields("Username") = recselector
                    .MoveNext
                Loop
                
                If recselector = g_strUser Then
                    pctUserTxtPassword.PasswordChar = "*"
                    pctUserCmdChangePassword.Visible = True
                    If userPermOK = True Then
                        pctUserCmdChangePassword.Visible = True
                        pctUserCmdChangePassword.enabled = True
                        pctUserLblActive.enabled = True
                        pctUserTxtPassword.PasswordChar = "*"
                        pctUser4.Visible = True
                        pctUser4lblActions.Caption = g_strUser & "'s actions"
                        pctUser2NoPerm.Visible = False
                        'pctUser3NoPerm.Visible = False
                        pctUser4NoPerm.Visible = False
                        pctUserCmdActive.enabled = True
                    Else
                        pctUser2NoPerm.Visible = True
                        PctUser3NoPerm.Visible = True
                        pctUser4NoPerm.Visible = True
                        
                        pctUser4.Visible = False
                        pctUserCmdActive.enabled = False
                    End If
                Else
                    If userPermOK = True Then
                        pctUserCmdActive.enabled = True
                        pctUserCmdChangePassword.Visible = True
                        pctUserCmdChangePassword.enabled = True
                        pctUserLblActive.enabled = True
                        pctUserTxtPassword.PasswordChar = "*"
                        pctUser4.Visible = True
                        pctUser4lblActions.Caption = g_strUser & "'s actions"
                        pctUser2NoPerm.Visible = False
                        'pctUser3NoPerm.Visible = False
                        pctUser4NoPerm.Visible = False
                    Else
                        pctUser2NoPerm.Visible = True
                        PctUser3NoPerm.Visible = True
                        pctUser4NoPerm.Visible = True
                        pctUserCmdChangePassword.Visible = False
                        pctUserTxtPassword.PasswordChar = "*"
                        pctUser4.Visible = False
                        pctUserCmdActive.enabled = False
                    End If
                End If
                
                pctUserLblUserDetails.Caption = .Fields("Username") & "'s details"
                pctUser.Visible = True
            
                pctUserTxtUsername.Text = .Fields("Username")
                pctUserTxtPassword.Text = .Fields("password")
            
                If .Fields("Status") = 1 Then
                    pctUserLblActive.Caption = .Fields("username") & "'s account is active."
                    pctUserLblActive.ForeColor = &H4000&
                    pctUserCmdActive.Caption = "Deactivate"
                Else
                    pctUserLblActive.Caption = .Fields("username") & "'s account has been deactivated."
                    pctUserLblActive.ForeColor = &HC0&
                    pctUserCmdActive.Caption = "Activate"
                End If
            
            
                If .Fields("fullname") <> "" Then
                    pctUserTxtFullName.Text = .Fields("fullname")
                Else
                    pctUserTxtFullName.Text = ""
                End If
            
                If .Fields("jobdescription") <> "" Then
                    pctUserTxtJobTitle.Text = .Fields("jobdescription")
                Else
                    pctUserTxtJobTitle.Text = ""
                End If
            
                If .Fields("hometel") <> "" Then
                    pctUserTxtHomeTel.Text = .Fields("hometel")
                Else
                    pctUserTxtHomeTel.Text = ""
                End If
            
                If .Fields("homeemail") <> "" Then
                    pctUserTxtHomeEmail.Text = .Fields("homeemail")
                Else
                    pctUserTxtHomeEmail.Text = ""
                End If
            
                If .Fields("worktel") <> "" Then
                    pctUserTxtWorkTel.Text = .Fields("worktel")
                Else
                    pctUserTxtWorkTel.Text = ""
                End If
            
                If .Fields("workfax") <> "" Then
                    pctUserTxtWorkFax.Text = .Fields("workfax")
                Else
                    pctUserTxtWorkFax.Text = ""
                End If
            
                If .Fields("workemail") <> "" Then
                    pctUserTxtWorkEmail.Text = .Fields("workemail")
                Else
                    pctUserTxtWorkEmail.Text = ""
                End If
            
                If .Fields("mobile") <> "" Then
                    pctUserTxtMobile.Text = .Fields("mobile")
                Else
                    pctUserTxtMobile.Text = ""
                End If
        
                If userPermOK = True Then
                    TabMain.Tab = 1
                    pctUser2LblUserPerms.Caption = .Fields("username") & "'s permissions"
        
                    With DEnv1.rstblUserPermissions
                        .Close
                        .Source = ("Select * from tblUserPermissions where username = '" & Flexgrid.Text & "'")
                        .Open
                        On Error Resume Next
                        .MoveFirst
                        irPermCount = 0
                        Do Until .EOF
                            irPermCount = irPermCount + 1
                            .MoveNext
                        Loop
                    End With
                    With pctUser2Flexgrid
                        .Clear
                        .Rows = irPermCount + 1
                        .Cols = 3
                        .ColWidth(0) = 0
                        .ColWidth(1) = 2750
                        .ColWidth(2) = 1100
                        .TextMatrix(0, 0) = "ID"
                        .TextMatrix(0, 1) = "Description"
                        .TextMatrix(0, 2) = "Available"
                        On Error Resume Next
                        .Row = 1
                        .Col = 0
                        .RowSel = irPermCount
                        .ColSel = 2
                        
                       ' For i = 1 To irPermCount
                      '      .Col = 2
                      '      If .Text = -1 Then
                      '          .Text = "Yes"
                      '      Else
                      '          .Text = "No"
                      '      End If
                      '      .Row = .Row + 1
                      '  Next i
                      '      .Col = 2
                     '   If .Text = -1 Then
                     '       .Text = "Yes"
                      '  Else
                     '       .Text = "No"
                     '   End If
                        
                    End With
        
                    On Error Resume Next
                    DEnv1.rstblUserPermissions.MoveFirst
                    With pctUser2Flexgrid
                        .Visible = False
                        .Clip = DEnv1.rstblUserPermissions.GetString
                        .Col = 0
                        .Sort = flexSortStringAscending
                        .Row = 1
                        .Col = 0
                        .RowSel = 1
                        '.ColSel = 1
                        .Visible = True
                        .ColAlignment(0) = 1
                        .ColAlignment(1) = 1
                    End With
         
                Else
                    pctUser2.Visible = True
                End If
            End With
        End Select
    End If
pctUser2.Visible = True
    TabMain.Tab = 0

End Function

Private Sub txtsearch_Change()
If Option1.Value = True Then
    If Option3.Value = True Then
        PopulateFrm latestentity, "option1", True, txtsearch.Text
    Else
        PopulateFrm latestentity, "option1", False, txtsearch.Text
    End If
Else
    If Option3.Value = True Then
        PopulateFrm latestentity, "option2", True, txtsearch.Text
    Else
        PopulateFrm latestentity, "option2", False, txtsearch.Text
    End If
End If
populateFlex

End Sub

