Attribute VB_Name = "basSettings"
Option Explicit


Public Sub Load_Settings()
    gCurrentServer = oReg.GetSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "CurrentServer", "")
    gCurrentDatabase = oReg.GetSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "CurrentDB", "")
    gCurrentUser = oReg.GetSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "CurrentUser", "")
    gCurrentPassword = oReg.GetSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "CurrentPassword", "")
    gSQLDriver = oReg.GetSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "SQLDriverString", "")
    gbCommentCode = oReg.GetSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "CommentCode", True)
    gbConstants = oReg.GetSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "ADOConstants", True)
    
    goptDriver = oReg.GetSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "SQLDriver", "0")
    If Trim(gSQLDriver) = "" Then
        If goptDriver = 0 Then
            gSQLDriver = "Provider=SQLOLEDB"
        ElseIf gSQLDriver = "1" Then
            gSQLDriver = "Driver={SQL Server}"
        End If
    End If
    gsConnectionObject = oReg.GetSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "ConnectionObject", "objConn")
    gbCreateConnection = oReg.GetSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "CreateConnection", True)
    gsRecordsetObject = oReg.GetSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "RecordsetObject", "objRS")
    gbCreateRecordset = oReg.GetSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "CreateRecordset", True)
    gsCommandObject = oReg.GetSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "CommandObject", "objCmd")
    gbCreateCommand = oReg.GetSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "CreateCommand", True)
    gsConnectionString = oReg.GetSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "ConnectionString", "strConnectionString")
    gbCreateConnectionString = oReg.GetSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "CreateConnectionString", True)
    gbRecordsAffected = oReg.GetSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "RecordsAffected", False)
    gbReturnRecordset = oReg.GetSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "ReturnRecordset", True)
    
     
    
End Sub

Public Sub Save_Settings()
    '-- Save Size Settings
    If frmMain.WindowState = 0 Then
        sResult = oReg.SaveSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "MainTop", frmMain.Top, REG_EXPAND_SZ)
        sResult = oReg.SaveSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "MainLeft", frmMain.Left, REG_EXPAND_SZ)
        sResult = oReg.SaveSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "Mainwidth", frmMain.Width, REG_EXPAND_SZ)
        sResult = oReg.SaveSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "Mainheight", frmMain.Height, REG_EXPAND_SZ)
    End If
    sResult = oReg.SaveSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "CurrentServer", gCurrentServer, REG_EXPAND_SZ)
    sResult = oReg.SaveSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "CurrentDB", gCurrentDatabase, REG_EXPAND_SZ)
    sResult = oReg.SaveSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "CurrentUser", gCurrentUser, REG_EXPAND_SZ)
    sResult = oReg.SaveSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "CurrentPassword", gCurrentPassword, REG_EXPAND_SZ)
    sResult = oReg.SaveSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "SQLDriverString", gSQLDriver, REG_EXPAND_SZ)
    sResult = oReg.SaveSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "CommentCode", gbCommentCode, REG_EXPAND_SZ)
    sResult = oReg.SaveSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "ADOConstants", gbConstants, REG_EXPAND_SZ)
    sResult = oReg.SaveSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "SQLDriver", goptDriver, REG_EXPAND_SZ)
    sResult = oReg.SaveSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "ConnectionObject", gsConnectionObject, REG_EXPAND_SZ)
    sResult = oReg.SaveSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "CreateConnection", gbCreateConnection, REG_EXPAND_SZ)
    sResult = oReg.SaveSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "RecordsetObject", gsRecordsetObject, REG_EXPAND_SZ)
    sResult = oReg.SaveSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "CreateRecordset", gbCreateRecordset, REG_EXPAND_SZ)
    sResult = oReg.SaveSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "CommandObject", gsCommandObject, REG_EXPAND_SZ)
    sResult = oReg.SaveSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "CreateCommand", gbCreateCommand, REG_EXPAND_SZ)
    sResult = oReg.SaveSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "ConnectionString", gsConnectionString, REG_EXPAND_SZ)
    sResult = oReg.SaveSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "CreateConnectionString", gbCreateConnectionString, REG_EXPAND_SZ)
    sResult = oReg.SaveSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "RecordsAffected", gbRecordsAffected, REG_EXPAND_SZ)
    sResult = oReg.SaveSetting(HKEY_LOCAL_MACHINE, RegSettingsPath, "ReturnRecordset", gbReturnRecordset, REG_EXPAND_SZ)
   
End Sub

Public Sub Center_Form(frm As Form)
    frm.Left = (Screen.Width - frm.Width) / 2
    frm.Top = (Screen.Height - frm.Height) / 2

End Sub
