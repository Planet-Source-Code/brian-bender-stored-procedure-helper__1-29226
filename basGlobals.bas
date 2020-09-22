Attribute VB_Name = "basGlobals"
Option Explicit

Public oReg As New clsRegistry

Public Const RegSettingsPath = "Software\BHC Solutions\SPHelper\Settings"

Public ConnectionString As String
Public gCurrentDatabase As String
Public gCurrentServer As String
Public gCurrentSproc As String
Public gCurrentUser As String
Public gCurrentPassword As String
Public gSQLDriver As String
Public gbReturnRecordset As Boolean
Public gbConstants As Boolean
Public gbSplash As Boolean

Public sResult As Variant

Public gbColorizing As Boolean

'-- Options
Public gbCreateConnection As Boolean
Public gbCreateRecordset As Boolean
Public gbCreateCommand As Boolean
Public gbCreateConnectionString As Boolean
Public gsConnectionObject As String
Public gsRecordsetObject As String
Public gsCommandObject As String
Public gsConnectionString As String
Public goptDriver As String
Public gbCommentCode As Boolean
Public gbRecordsAffected As Boolean

Public oRS As New ADODB.Recordset

Public Enum dbType
    db_Access = 0
    db_sql = 1
End Enum

Public Enum spType
    sp_Select = 0
    sp_Insert = 1
    sp_Update = 2
End Enum

Public Enum CommandType
    cmd_VB = 0
    cmd_asp = 1
End Enum

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function PutFocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function SendMessageBynum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)


Public Const LVM_FIRST = &H1000
Public Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)
Public Const LVSCW_AUTOSIZE            As Long = -1
Public Const LVSCW_AUTOSIZE_USEHEADER  As Long = -2

Public Const VK_DOWN = &H28
Public Const VK_M = &H4D
Public Const KEYEVENTF_KEYUP = &H2
