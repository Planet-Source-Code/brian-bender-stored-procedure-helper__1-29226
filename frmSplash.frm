VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5790
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   5535
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   315
         Left            =   300
         TabIndex        =   2
         Top             =   2520
         Width           =   975
      End
      Begin VB.Image Image2 
         Height          =   795
         Left            =   240
         Picture         =   "frmSplash.frx":1CFA
         Top             =   360
         Width           =   3000
      End
      Begin VB.Image Image1 
         Height          =   435
         Left            =   2880
         Picture         =   "frmSplash.frx":2B4F
         Top             =   1740
         Width           =   450
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Stored Procedure Helper "
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
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   5055
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   2760
         X2              =   5280
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Licensed to Beta User"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1980
         TabIndex        =   5
         Top             =   2400
         Width           =   3315
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Copyright BHC Solutions 2001. All Rights Reserved"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   780
         TabIndex        =   4
         Top             =   2640
         Width           =   4515
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         Caption         =   "Version 1.0.0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   3
         Top             =   1920
         Width           =   2415
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3545
      Left            =   0
      TabIndex        =   0
      Top             =   -310
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   6271
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
    gbSplash = False
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    Center_Form frmSplash
       
End Sub



