VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "SP Helper"
   ClientHeight    =   7185
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9705
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   9705
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8940
      Top             =   180
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
            Picture         =   "frmMain.frx":0442
            Key             =   "SP"
            Object.Tag             =   "SP"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CD4
            Key             =   "vb"
            Object.Tag             =   "vb"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1566
            Key             =   "asp"
            Object.Tag             =   "asp"
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtfTemp 
      Height          =   135
      Left            =   1860
      TabIndex        =   17
      Top             =   3180
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   238
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmMain.frx":19B8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   16
      Top             =   6870
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3588
            MinWidth        =   2292
            Picture         =   "frmMain.frx":1A41
            Text            =   " Server:"
            TextSave        =   " Server:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4012
            MinWidth        =   2716
            Picture         =   "frmMain.frx":1E93
            Text            =   " Database:"
            TextSave        =   " Database:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5000
            MinWidth        =   3704
            Picture         =   "frmMain.frx":22E5
            Text            =   " Procedure:"
            TextSave        =   " Procedure:"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3836
            Picture         =   "frmMain.frx":243F
            Text            =   "Parameters:"
            TextSave        =   "Parameters:"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab tabCode 
      Height          =   3015
      Left            =   60
      TabIndex        =   15
      Top             =   3840
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   5318
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   " Stored Procedure Code"
      TabPicture(0)   =   "frmMain.frx":2891
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblSPCode"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Container"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   " VB Code"
      TabPicture(1)   =   "frmMain.frx":28AD
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "rtbVB"
      Tab(1).Control(1)=   "lblVBCode"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   " ASP Code"
      TabPicture(2)   =   "frmMain.frx":28C9
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "rtbASP"
      Tab(2).Control(1)=   "lblASPCode"
      Tab(2).ControlCount=   2
      Begin VB.PictureBox Container 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         ForeColor       =   &H80000008&
         Height          =   2235
         Left            =   120
         ScaleHeight     =   2205
         ScaleWidth      =   9345
         TabIndex        =   19
         Top             =   360
         Width           =   9375
         Begin VB.PictureBox picLines 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2295
            Left            =   0
            ScaleHeight     =   2295
            ScaleWidth      =   480
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin RichTextLib.RichTextBox rtbSQL 
            Height          =   2250
            Left            =   480
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   -30
            Width           =   8885
            _ExtentX        =   15663
            _ExtentY        =   3969
            _Version        =   393217
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            ScrollBars      =   3
            Appearance      =   0
            RightMargin     =   99999
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"frmMain.frx":28E5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin RichTextLib.RichTextBox rtbVB 
         Height          =   2250
         Left            =   -74880
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   360
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   3969
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         RightMargin     =   99999
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmMain.frx":2965
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox rtbASP 
         Height          =   2250
         Left            =   -74880
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   360
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   3969
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         RightMargin     =   99999
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmMain.frx":29E5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblASPCode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ASP Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -74880
         TabIndex        =   23
         Top             =   120
         Width           =   9375
      End
      Begin VB.Label lblVBCode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Visual Basic Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -74880
         TabIndex        =   22
         Top             =   120
         Width           =   9375
      End
      Begin VB.Label lblSPCode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Stored Procedure Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   9375
      End
   End
   Begin VB.ComboBox cboSprocs 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   180
      TabIndex        =   14
      Top             =   3360
      Width           =   2535
   End
   Begin VB.ComboBox cboDatabases 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   180
      TabIndex        =   11
      Top             =   2700
      Width           =   2535
   End
   Begin VB.Frame FrameParams 
      Height          =   3735
      Left            =   2760
      TabIndex        =   6
      Top             =   -60
      Width           =   6915
      Begin MSComctlLib.ListView lvSP 
         Height          =   3135
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   6675
         _ExtentX        =   11774
         _ExtentY        =   5530
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   14
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "PARAMETER_NAME"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ORDINAL_POSITION"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "PARAMETER_TYPE"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "PARAMETER_HASDEFAULT"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "PARAMETER_DEFAULT"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "IS_NULLABLE"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "DATA_TYPE"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "CHAR_MAX_LENGTH"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "CHAR_OCTET_LENGTH"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "NUMERIC_PRECISION"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "NUMERIC_SCALE"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "DESCRIPTION"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "TYPE_NAME"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "LOCAL_TYPE_NAME"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblParameters 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Parameters"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   6675
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   60
      TabIndex        =   0
      Top             =   -60
      Width           =   2655
      Begin VB.CommandButton cmdConnect 
         Appearance      =   0  'Flat
         Caption         =   "&Connect"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2040
         Width           =   1035
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   4
         Text            =   "admin"
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox txtUser 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "sa"
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox txtServer 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "bbender2000s"
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   1035
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "User"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Server"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SProcs"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   180
      TabIndex        =   13
      Top             =   3120
      Width           =   1035
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Databases"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   180
      TabIndex        =   12
      Top             =   2460
      Width           =   1035
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuParameters 
         Caption         =   "Parameters"
         Begin VB.Menu mnuSelectAll 
            Caption         =   "Select &All"
         End
         Begin VB.Menu mnuDeselectAll 
            Caption         =   "&Deselect All"
         End
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "&Settings"
      Begin VB.Menu mnuoptions 
         Caption         =   "&Options"
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About SP Helper"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bDBLoading As Boolean
Dim bSPLoading As Boolean
Dim bServerLoading As Boolean

Private Sub Form_Load()
    '-- Subclass the rtb so we can scroll the line numbers
    lPrevWndProc = SetWindowLong(rtbSQL.hwnd, GWL_WNDPROC, AddressOf WindowProc)
    Load_Settings
    tabCode.TabPicture(0) = ImageList1.ListImages("SP").Picture
    tabCode.TabPicture(1) = ImageList1.ListImages("vb").Picture
    tabCode.TabPicture(2) = ImageList1.ListImages("asp").Picture
    Adjust_Listview_Columns lvSP
    txtServer.Text = gCurrentServer
    txtUser.Text = gCurrentUser
    txtPassword.Text = gCurrentPassword
 
    Me.Top = oReg.GetSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "MainTop", -1)
    Me.Left = oReg.GetSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "MainLeft", -1)
    Me.Width = oReg.GetSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "MainWidth", 9825)
    Me.Height = oReg.GetSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "MainHeight", 7875)
    
    If Me.Top = -1 And Me.Left = -1 Then Center_Form Me
    Unload frmSplash
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '-- kill the subclass so we don't screw up someone's machine
    Call SetWindowLong(rtbSQL.hwnd, GWL_WNDPROC, lPrevWndProc)
    Save_Settings
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If frmMain.Height < 5925 Then frmMain.Height = 5925
    If frmMain.Width < 3450 Then frmMain.Width = 3450
    FrameParams.Width = Me.Width - 2920
    lvSP.Width = FrameParams.Width - 240
    lblParameters.Width = lvSP.Width
    
    tabCode.Width = Me.Width - 240
    DoEvents
    Container.Width = tabCode.Width - 240
    Container.Height = tabCode.Height - 780
    rtbSQL.Height = Container.Height + 10
    picLines.Height = Container.Height
    lblSPCode.Width = Container.Width
    rtbSQL.Width = Container.Width - picLines.Width
    rtbVB.Height = Container.Height
    lblVBCode.Width = Container.Width
    rtbVB.Width = Container.Width
    rtbASP.Height = Container.Height
    lblASPCode.Width = Container.Width
    rtbASP.Width = Container.Width
    tabCode.Height = Me.Height - 4900
      
    ' Refresh the line numbers
    DrawLines picLines
    If Container.Height <> tabCode.Height - 780 Then Form_Resize
End Sub

Public Sub DrawLines(picTo As PictureBox)

    Dim lLine As Long
    Dim lCount As Long
    Dim lCurrent As Long
    Dim hBr As Long
    Dim lEnd As Long
    Dim lhDC As Long
    Dim bComplete As Boolean
    Dim tR As RECT
    Dim tTR As RECT
    Dim oCol As OLE_COLOR
    Dim lStart As Long
    Dim lEndLine As Long
    Dim tPO As POINTAPI
    Dim lLineHeight As Long
    Dim hPen As Long
    Dim hPenOld As Long
 
    lhDC = picTo.hdc
    DrawText lhDC, "Hy", 2, tTR, DT_CALCRECT
    lLineHeight = tTR.Bottom - tTR.Top
    
    lCount = LineCount(rtbSQL.hwnd)
    If lCount < 50 Then lCount = 50
    lCurrent = SendMessageLong(rtbSQL.hwnd, EM_LINEFROMCHAR, rtbSQL.SelStart, 0&)
    lStart = rtbSQL.SelStart
    lEnd = rtbSQL.SelStart + rtbSQL.SelLength - 1
    If (lEnd > lStart) Then
       lEndLine = LineForCharacterIndex(lEnd, rtbSQL.hwnd)
    Else
       lEndLine = lCurrent
    End If
    lLine = GetFirstVisibleLine(rtbSQL.hwnd)
    GetClientRect picTo.hwnd, tR
    lEnd = tR.Bottom - tR.Top
    
    hBr = CreateSolidBrush(TranslateColor(picTo.BackColor))
    FillRect lhDC, tR, hBr
    DeleteObject hBr
    tR.Left = 2
    tR.Right = tR.Right - 2
    tR.Top = 0
    tR.Bottom = tR.Top + lLineHeight
    
    SetTextColor lhDC, TranslateColor(vbButtonShadow)
    
    Do
       ' Ensure correct colour:
       If (lLine = lCurrent) Then
          SetTextColor lhDC, TranslateColor(vbWindowText)
       ElseIf (lLine = lEndLine + 1) Then
          SetTextColor lhDC, TranslateColor(vbButtonShadow)
       End If
       ' Draw the line number:
       DrawText lhDC, CStr(lLine + 1), -1, tR, DT_RIGHT
       
       ' Increment the line:
       lLine = lLine + 1
       ' Increment the position:
       OffsetRect tR, 0, lLineHeight
       If (tR.Bottom > lEnd) Or (lLine + 1 > lCount) Then
          bComplete = True
       End If
    Loop While Not bComplete
    
    ' Draw a line...
    MoveToEx lhDC, tR.Right + 1, 0, tPO
    hPen = CreatePen(PS_SOLID, 1, TranslateColor(vbButtonShadow))
    hPenOld = SelectObject(lhDC, hPen)
    LineTo lhDC, tR.Right + 1, lEnd
    SelectObject lhDC, hPenOld
    DeleteObject hPen
    If picTo.AutoRedraw Then
       picTo.Refresh
    End If
   
End Sub

Private Sub lvSP_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Refresh_Code
    lvSP.SetFocus
End Sub

Private Sub mnuAbout_Click()
    frmSplash.Show vbModal
End Sub

Private Sub mnuDeselectAll_Click()
    Dim I As Integer
    For I = 1 To lvSP.ListItems.Count
        lvSP.ListItems(I).Checked = False
    Next
    Refresh_Code
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuoptions_Click()
    frmOptions.Show , Me
End Sub

Private Sub mnuRefresh_Click()
    rtbASP.Text = ""
    rtbVB.Text = ""
    rtbSQL.Text = ""
    Dim db As String
    db = gCurrentDatabase
    Dim sp
    sp = gCurrentSproc
    If Trim(txtServer.Text) <> "" And Trim(txtPassword.Text) <> "" And Trim(txtUser) <> "" Then
        cmdConnect_Click
        If db <> "" Then
            cboDatabases.Text = db
            cboDatabases_Click
            If sp <> "" Then
                cboSprocs.Text = sp
                cboSprocs_Click
            End If
        End If
    End If
        
End Sub

Private Sub mnuSelectAll_Click()
    Dim I As Integer
    For I = 1 To lvSP.ListItems.Count
        lvSP.ListItems(I).Checked = True
    Next
    Refresh_Code
End Sub

Private Sub rtbsql_KeyUp(KeyCode As Integer, Shift As Integer)
    DrawLines picLines
End Sub

Private Sub rtbsql_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    DrawLines picLines
End Sub

Private Sub rtbsql_SelChange()
    DrawLines picLines
End Sub

Private Sub rtbsql_Change()
    DrawLines picLines
End Sub

Private Sub cboDatabases_Click()
    DoEvents
    If bDBLoading Then Exit Sub
    If cboDatabases.Text = gCurrentDatabase And Not bServerLoading Then Exit Sub
    gCurrentDatabase = cboDatabases.Text
    gCurrentSproc = ""
    Screen.MousePointer = vbHourglass
    bSPLoading = True
    Create_Connection db_sql, cboDatabases.Text, Trim(txtServer.Text), Trim(txtUser.Text), Trim(txtPassword.Text)
    Set oRS = ExecuteSP("sp_stored_procedures", sp_Select)
    cboSprocs.Clear
    Do Until oRS.EOF
        If oRS("Procedure_Owner") <> "system_function_schema" Then cboSprocs.AddItem Left(oRS("Procedure_Name"), InStr(1, oRS("Procedure_Name"), ";") - 1)
        oRS.MoveNext
    Loop
    If cboSprocs.ListCount > 0 Then
        cboSprocs.Text = "Choose Sproc"
        cboSprocs.Enabled = True
    Else
        cboSprocs.Text = ""
        gCurrentSproc = ""
        cboSprocs.Enabled = False
    End If
    Screen.MousePointer = vbDefault
    bSPLoading = False
    Update_Status
End Sub

Private Sub cboSprocs_Click()
    DoEvents
    If bDBLoading Then Exit Sub
    If cboDatabases.Text = "Choose Sproc" Then Exit Sub
    If cboDatabases.Text = gCurrentSproc Then Exit Sub
    gCurrentSproc = cboSprocs.Text
    Screen.MousePointer = vbHourglass
    Create_Connection db_sql, gCurrentDatabase, Trim(txtServer.Text), Trim(txtUser.Text), Trim(txtPassword.Text)
    Load_ListView "exec sp_procedure_params_rowset '" & gCurrentSproc & "'", lvSP
    rtbSQL.Text = ""
    rtbVB.Text = ""
    rtbASP.Text = ""
    ColorizeSQL rtbSQL, Load_Sproc
    Write_Command cmd_VB, rtbVB, lvSP
    Write_Command cmd_asp, rtbASP, lvSP
    Update_Status
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdConnect_Click()
    bServerLoading = True
    cboDatabases.Clear
    cboSprocs.Clear
    DoEvents
    Select Case True
        Case Trim(txtServer.Text) = ""
            MsgBox "Please enter a server name"
        Case Trim(txtUser.Text) = ""
            MsgBox "Please Enter a user name"
        Case Else
            Screen.MousePointer = vbHourglass
            gCurrentServer = Trim(txtServer.Text)
            gCurrentUser = Trim(txtUser.Text)
            gCurrentPassword = Trim(txtPassword.Text)
            bDBLoading = True
            Create_Connection db_sql, "Master", gCurrentServer, gCurrentUser, gCurrentPassword
            Set oRS = ExecuteSP("sp_databases", sp_Select)
            cboDatabases.Clear
            If Not oRS Is Nothing Then
                Do Until oRS.EOF
                    cboDatabases.AddItem oRS("Database_Name")
                    oRS.MoveNext
                Loop
                cboDatabases.Text = "Master"
                cboDatabases.Enabled = True
            End If
            bDBLoading = False
            cboDatabases_Click
            sResult = oReg.SaveSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "CurrentServer", gCurrentServer, REG_EXPAND_SZ)
            sResult = oReg.SaveSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "CurrentUser", gCurrentUser, REG_EXPAND_SZ)
            sResult = oReg.SaveSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "CurrentPassword", gCurrentPassword, REG_EXPAND_SZ)
            sResult = oReg.SaveSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "Currentdatabase", gCurrentDatabase, REG_EXPAND_SZ)
                           
            Screen.MousePointer = vbDefault
    End Select
    bServerLoading = False
    Update_Status
End Sub


Private Sub Update_Status()
    SB.Panels(1).Text = "Server: " & gCurrentServer
    SB.Panels(2).Text = "Database: " & gCurrentDatabase
    SB.Panels(3).Text = "Procedure: " & gCurrentSproc
    If lvSP.ListItems.Count > 0 Then
        SB.Panels(4).Text = "Parameters: " & lvSP.ListItems.Count
    Else
        SB.Panels(4).Text = "Parameters: "
    End If
    If gCurrentSproc = "" Then
        rtbSQL.Text = ""
        rtbVB.Text = ""
        rtbASP.Text = ""
        lvSP.ListItems.Clear
    End If
End Sub

Public Sub Refresh_Code()
    Dim iTab As Integer
    'iTab = tabCode.Tab
    Dim iLine As Long
    iLine = LastVisibleLine(rtbVB.hwnd)
    Write_Command cmd_VB, rtbVB, lvSP
    'tabCode.Tab = 1
    rtbVB.SetFocus
    rtbVB.SelStart = 0
    
    Dim iCount As Integer
    For iCount = 1 To iLine
        keybd_event VK_DOWN, 0, 0, 0
        keybd_event VK_DOWN, 0, KEYEVENTF_KEYUP, 0
    Next
    iLine = LastVisibleLine(rtbASP.hwnd)
    Write_Command cmd_asp, rtbASP, lvSP
    'tabCode.Tab = 2
    rtbASP.SetFocus
    rtbASP.SelStart = 0
    For iCount = 1 To iLine
        keybd_event VK_DOWN, 0, 0, 0
        keybd_event VK_DOWN, 0, KEYEVENTF_KEYUP, 0
    Next
    'tabCode.Tab = iTab
End Sub
