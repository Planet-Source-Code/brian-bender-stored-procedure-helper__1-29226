VERSION 5.00
Begin VB.Form frmOptions 
   Caption         =   "Options"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6240
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5580
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRestore 
      Caption         =   "Restore &Default"
      Height          =   255
      Left            =   60
      TabIndex        =   21
      Top             =   5220
      Width           =   1275
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   255
      Left            =   4920
      TabIndex        =   14
      Top             =   5220
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   255
      Left            =   3300
      TabIndex        =   13
      Top             =   5220
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   255
      Left            =   1680
      TabIndex        =   12
      Top             =   5220
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   5115
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.CheckBox chkReturnRecordset 
         Appearance      =   0  'Flat
         Caption         =   "Return Recordset"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   840
         TabIndex        =   22
         Top             =   4320
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   120
         TabIndex        =   18
         Top             =   4020
         Width           =   5895
         Begin VB.CheckBox chkConstants 
            Appearance      =   0  'Flat
            Caption         =   "Use ADO Constants"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3480
            TabIndex        =   23
            Top             =   600
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.CheckBox chkRecordsEffected 
            Appearance      =   0  'Flat
            Caption         =   "Return Records Affected"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   720
            TabIndex        =   20
            Top             =   600
            Value           =   1  'Checked
            Width           =   2115
         End
         Begin VB.CheckBox chkComment 
            Appearance      =   0  'Flat
            Caption         =   "Comment Code"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3480
            TabIndex        =   19
            Top             =   300
            Value           =   1  'Checked
            Width           =   1395
         End
      End
      Begin VB.CheckBox chkCreateConnectionString 
         Appearance      =   0  'Flat
         Caption         =   "Create Connection String"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3600
         TabIndex        =   17
         Top             =   3420
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox chkCreateCommand 
         Appearance      =   0  'Flat
         Caption         =   "Create Command Object"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3600
         TabIndex        =   16
         Top             =   2520
         Value           =   1  'Checked
         Width           =   2115
      End
      Begin VB.CheckBox chkCreateRecordset 
         Appearance      =   0  'Flat
         Caption         =   "Create Recordset Object"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3600
         TabIndex        =   15
         Top             =   1620
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.OptionButton optConnection 
         Appearance      =   0  'Flat
         Caption         =   "Use SQL Native Driver"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   2340
         TabIndex        =   11
         Top             =   3720
         Width           =   1935
      End
      Begin VB.OptionButton optConnection 
         Appearance      =   0  'Flat
         Caption         =   "Use OLE DB Provider"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   3720
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.TextBox txtConnectionString 
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
         Left            =   240
         TabIndex        =   8
         Text            =   "strConnectionString"
         Top             =   3360
         Width           =   3135
      End
      Begin VB.TextBox txtCommand 
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
         Left            =   240
         TabIndex        =   6
         Text            =   "objCmd"
         Top             =   2460
         Width           =   3135
      End
      Begin VB.TextBox txtRecordset 
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
         Left            =   240
         TabIndex        =   4
         Text            =   "objRS"
         Top             =   1560
         Width           =   3135
      End
      Begin VB.CheckBox chkCreatConnection 
         Appearance      =   0  'Flat
         Caption         =   "Create Connection Object"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3600
         TabIndex        =   3
         Top             =   720
         Value           =   1  'Checked
         Width           =   2235
      End
      Begin VB.TextBox txtConnection 
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
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Text            =   "objConn"
         Top             =   600
         Width           =   3135
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   6000
         Y1              =   3975
         Y2              =   3975
      End
      Begin VB.Line Line7 
         BorderColor     =   &H80000003&
         X1              =   120
         X2              =   6000
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Connection String"
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
         TabIndex        =   9
         Top             =   3120
         Width           =   1755
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   6000
         Y1              =   2835
         Y2              =   2835
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000003&
         X1              =   120
         X2              =   6000
         Y1              =   2100
         Y2              =   2100
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Command Object"
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
         TabIndex        =   7
         Top             =   2220
         Width           =   1755
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   6000
         Y1              =   1935
         Y2              =   1935
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000003&
         X1              =   120
         X2              =   6000
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Recordset Object"
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
         TabIndex        =   5
         Top             =   1320
         Width           =   1755
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   6000
         Y1              =   1035
         Y2              =   1035
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         X1              =   120
         X2              =   6000
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Connection Object"
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
         TabIndex        =   2
         Top             =   360
         Width           =   1755
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function Save_Changes() As Boolean
    If Trim(txtConnection.Text) = "" Then
        MsgBox "Please specify a connection object", vbCritical, App.EXEName
        txtConnection.SetFocus
        Exit Function
    Else
        gsConnectionObject = txtConnection.Text
    End If
    If Trim(txtRecordset.Text) = "" Then
        MsgBox "Please specify a recordset object", vbCritical, App.EXEName
        txtRecordset.SetFocus
        Exit Function
    Else
        gsRecordsetObject = txtRecordset.Text
    End If
    If Trim(txtCommand.Text) = "" Then
        MsgBox "Please specify a command object", vbCritical, App.EXEName
        txtCommand.SetFocus
        Exit Function
    Else
        gsCommandObject = txtCommand.Text
    End If
    If Trim(txtConnectionString.Text) = "" Then
        MsgBox "Please specify a connection string", vbCritical, App.EXEName
        txtConnectionString.SetFocus
        Exit Function
    Else
        gsConnectionString = txtConnectionString.Text
    End If
    
    gbCreateConnection = CBool(chkCreatConnection.Value)
    gbCreateRecordset = CBool(chkCreateRecordset.Value)
    gbCreateCommand = CBool(chkCreateCommand.Value)
    gbCreateConnectionString = CBool(chkCreateConnectionString.Value)
    Dim I As Integer
    For I = 0 To optConnection.Count - 1
        If optConnection(I).Value = True Then goptDriver = I
    Next
    gbRecordsAffected = CBool(chkRecordsEffected.Value)
    gbCommentCode = CBool(chkComment.Value)
    gbConstants = CBool(chkConstants.Value)
    gbReturnRecordset = CBool(chkReturnRecordset.Value)
    Save_Changes = True
End Function

Private Sub Load_Settings()
    txtConnection.Text = gsConnectionObject
    txtRecordset.Text = gsRecordsetObject
    txtCommand.Text = gsCommandObject
    txtConnectionString = gsConnectionString
    chkCreatConnection.Value = IIf(gbCreateConnection = True, 1, 0)
    chkCreateRecordset.Value = IIf(gbCreateRecordset = True, 1, 0)
    chkCreateCommand.Value = IIf(gbCreateCommand = True, 1, 0)
    chkCreateConnectionString.Value = IIf(gbCreateConnectionString = True, 1, 0)
    optConnection(goptDriver).Value = True
    chkRecordsEffected.Value = IIf(gbRecordsAffected = True, 1, 0)
    chkComment.Value = IIf(gbCommentCode, 1, 0)
    chkConstants.Value = IIf(gbConstants, 1, 0)
    chkReturnRecordset.Value = IIf(gbReturnRecordset, 1, 0)
End Sub

Private Sub chkReturnRecordset_Click()
    If chkReturnRecordset.Value = vbUnchecked Then chkCreateRecordset.Value = vbUnchecked
End Sub

Private Sub cmdApply_Click()
    If Save_Changes = True Then
        Save_Settings
    End If
    frmMain.Refresh_Code
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    cmdApply_Click
    Unload Me
End Sub

Private Sub cmdRestore_Click()
    txtConnection.Text = "objConnection"
    txtRecordset.Text = "objRecordset"
    txtCommand.Text = "objCommand"
    txtConnectionString = "strConnectionString"
    chkCreatConnection.Value = vbChecked
    chkCreateRecordset.Value = vbChecked
    chkCreateCommand.Value = vbChecked
    chkCreateConnectionString.Value = vbChecked
    optConnection(0).Value = True
    chkRecordsEffected.Value = False
    chkComment.Value = vbChecked
    chkConstants.Value = vbChecked
    chkReturnRecordset.Value = vbChecked
End Sub

Private Sub Form_Load()
    Center_Form Me
    Load_Settings
End Sub
