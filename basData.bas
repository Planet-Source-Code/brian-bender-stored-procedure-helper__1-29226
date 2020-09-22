Attribute VB_Name = "basData"

Public Sub Main()
    If App.PrevInstance = True Then End
    frmSplash.Show
    gbSplash = True
    Pause 3
    frmMain.Show
End Sub
Private Sub Pause(Interval As Integer)
    Dim iTimer As Long
    iTimer = Timer
    Do Until Timer > iTimer + Interval
        If gbSplash = False Then Exit Sub
        DoEvents
        If Timer < iTimer Then
            Exit Do
            End If
    Loop
    
End Sub
Public Sub Create_Connection(Database_Type As dbType, db As String, Optional ServerName As String, Optional user As String, Optional Password As String)
    On Error GoTo Err_Handler
    
    Select Case Database_Type

        Case 0 '-- Access Database
            ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                               "Data Source=" & db & gsDBName & ";" & _
                               "User Id=" & user & ";" & _
                               "Jet OLEDB:Database Password=" & Password & ";"
            
        Case 1 '-- SQL Database
            ConnectionString = "Provider=SQLOLEDB;" & _
                               "Server=" & ServerName & ";" & _
                               "Database=" & db & ";" & _
                               "UID=" & user & ";PWD=" & Password & ";"
        Case Else
            MsgBox "Incorect Database Type", vbOKOnly + vbExclamation, "Create Connection Error!"
            Exit Sub
    
    End Select
    
    Set OpenConnection = New ADODB.Connection
    OpenConnection.Open ConnectionString
    If OpenConnection.State = adStateOpen Then
        gCurrentDatabase = db
        gCurrentServer = ServerName
        OpenConnection.Close
    Else
        gCurrentDatabase = ""
        gCurrentServer = ""
    End If
    Set OpenConnection = Nothing
    Exit Sub
    
Err_Handler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ExecuteSQL(ByVal SQL As String) As ADODB.Recordset

    Dim cnn As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim sTokens() As String

    On Error GoTo Err_Handler

    sTokens = Split(SQL)
    Set cnn = New ADODB.Connection
    cnn.Open ConnectionString

    If InStr(1, Left(UCase(SQL), 7), "INSERT") _
        Or InStr(1, Left(UCase(SQL), 7), "DELETE") _
        Or InStr(1, Left(UCase(SQL), 7), "UPDATE") Then
        cnn.Execute SQL
    Else
        Set RS = New ADODB.Recordset
        RS.Open Trim(SQL), cnn, adOpenKeyset, adLockOptimistic
        Set ExecuteSQL = RS
    End If

Exit_Execute_SQL:
    Set RS = Nothing
    Set cnn = Nothing
    Exit Function

Err_Handler:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, "Execute SQL Error!"
    Resume Exit_Execute_SQL

End Function

Public Function ExecuteSP(ByVal SQL As String, Procedure_Type As spType) As ADODB.Recordset

    Dim cnn As ADODB.Connection
    Dim RS As ADODB.Recordset

    On Error GoTo Err_Handler

    Set cnn = New ADODB.Connection
    cnn.Open ConnectionString

    Select Case Procedure_Type

        Case 0
            Set RS = New ADODB.Recordset
            RS.Open Trim(SQL), cnn ', adOpenKeyset, adLockOptimistic, adCmdStoredProc
            Set ExecuteSP = RS

        Case 1, 2
            cnn.Execute SQL

    End Select

Exit_Execute_SP:
    Set RS = Nothing
    Set cnn = Nothing
    Exit Function

Err_Handler:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, "Execute SQL Error!"
    Resume Exit_Execute_SP

End Function

Public Sub Load_ListView(sp As String, olv As ListView)
    Dim rsLV As New ADODB.Recordset
    Set rsLV = ExecuteSP(sp, sp_Select)
    If rsLV.EOF Then Exit Sub
    olv.ListItems.Clear
    Dim I As Integer
    LockWindowUpdate olv.hwnd
    Do Until rsLV.EOF
        Set LI = olv.ListItems.Add(, , rsLV(3).Value)
        LI.Checked = True
        For I = 4 To rsLV.Fields.Count - 1
            If IsNull(rsLV(I).Value) Then
                LI.ListSubItems.Add , , "NULL"
            Else
                If I = 5 Then
                    LI.ListSubItems.Add , , Get_ParameterDirection_Enum_Type(rsLV(I).Value) & "(" & rsLV(I).Value & ")"
                ElseIf I = 9 Then
                    LI.ListSubItems.Add , , Get_Enum_Type(rsLV(I).Value) & "(" & rsLV(I).Value & ")"
                Else
                    LI.ListSubItems.Add , , rsLV(I).Value
                End If
            End If
        Next
        rsLV.MoveNext
    Loop
    LockWindowUpdate 0&
    Adjust_Listview_Columns olv
End Sub

Public Sub Adjust_Listview_Columns(olv As ListView)
    For I = 0 To olv.ColumnHeaders.Count - 1
        SendMessage olv.hwnd, LVM_SETCOLUMNWIDTH, I, ByVal LVSCW_AUTOSIZE_USEHEADER
    Next
End Sub

Public Function Get_ParameterDirection_Enum_Type(ParameterType As Integer) As String
    Select Case ParameterType
        Case 0
            Get_ParameterDirection_Enum_Type = "adParamUnknown"
        Case 1
            Get_ParameterDirection_Enum_Type = "adParamInput"
        Case 2
            Get_ParameterDirection_Enum_Type = "adParamOutput"
        Case 3
            Get_ParameterDirection_Enum_Type = "adParamInputOutput"
        Case 4
            Get_ParameterDirection_Enum_Type = "adParamReturnValue"
    End Select
End Function
    

Public Function Get_Enum_Type(DataType As Integer) As String

    Select Case DataType
        Case 0
            Get_Enum_Type = "adEmpty"
        Case 2
            Get_Enum_Type = "adSmallInt"
        Case 3
            Get_Enum_Type = "adInteger"
        Case 4
            Get_Enum_Type = "adSingle"
        Case 5
            Get_Enum_Type = "adDouble"
        Case 6
            Get_Enum_Type = "adCurrency "
        Case 7
            Get_Enum_Type = "adDate"
        Case 8
            Get_Enum_Type = "adBSTR"
        Case 9
            Get_Enum_Type = "adIDispatch"
        Case 10
            Get_Enum_Type = "adError"
        Case 11
            Get_Enum_Type = "adBoolean"
        Case 12
            Get_Enum_Type = "adVariant"
        Case 13
            Get_Enum_Type = "adIUnknown"
        Case 14
            Get_Enum_Type = "adDecimal"
        Case 16
            Get_Enum_Type = "adTinyInt"
        Case 17
            Get_Enum_Type = "adUnsignedTinyInt"
        Case 18
            Get_Enum_Type = "adUnsignedSmallInt"
        Case 19
            Get_Enum_Type = "adUnsignedInt"
        Case 20
            Get_Enum_Type = "adBigInt"
        Case 21
            Get_Enum_Type = "adUnsignedBigInt"
        Case 64
            Get_Enum_Type = "adFileTime"
        Case 72
            Get_Enum_Type = "adGUID"
        Case 128
            Get_Enum_Type = "adBinary"
        Case 129
            Get_Enum_Type = "adChar"
        Case 130
            Get_Enum_Type = "adWChar"
        Case 131
            Get_Enum_Type = "adNumeric"
        Case 132
            Get_Enum_Type = "adUserDefined"
        Case 133
            Get_Enum_Type = "adDBDate"
        Case 134
            Get_Enum_Type = "adDBTime"
        Case 135
            Get_Enum_Type = "adDBTimeStamp"
        Case 136
            Get_Enum_Type = "adChapter"
        Case 138
            Get_Enum_Type = "adPropVariant"
        Case 139
            Get_Enum_Type = "adVarNumeric"
        Case 200
            Get_Enum_Type = "adVarChar"
        Case 201
            Get_Enum_Type = "adLongVarChar"
        Case 202
            Get_Enum_Type = "adVarWChar"
        Case 203
            Get_Enum_Type = "adLongVarWChar"
        Case 204
            Get_Enum_Type = "adVarBinary"
        Case 205
            Get_Enum_Type = "adLongVarBinary"
        Case 8192
            Get_Enum_Type = "adArray"
    End Select
End Function

Private Function Account_For_NULL(sValue As String)
    If IsNull(sValue) Then
        Account_For_NULL = "NULL"
    Else
        Account_For_NULL = Trim(sValue)
    End If
End Function

Public Function Load_Sproc() As String
    Set oRS = ExecuteSP("exec sp_helptext '" & gCurrentSproc & "'", sp_Select)
    Dim strSproc As String
    If oRS.State = adStateOpen Then
        Do Until oRS.EOF
            strSproc = strSproc & oRS("text")
            oRS.MoveNext
        Loop
        Load_Sproc = strSproc
    Else
        gCurrentSproc = ""
    End If
End Function

Public Sub Write_Command(cmd As CommandType, rtb As RichTextBox, lv As ListView)
    
    rtb.Text = ""
    
    Dim strCode As String
    Dim strParamCode As String
    Dim I As Integer
    
    strCode = vbCrLf
    
    If cmd = cmd_VB Then
        
        '-- VB Code
        
        '-- Create Connection String
        If gbCreateConnectionString = True Then
            If gbCommentCode = True Then strCode = strCode & "   '-- Create Connection String" & vbCrLf
            strCode = strCode & "   Dim " & gsConnectionString & " as String" & vbCrLf
            strCode = strCode & "   " & gsConnectionString & " = " & Chr(34) & gSQLDriver & ";" & Chr(34) & " & _" & vbCrLf
            strCode = strCode & "       " & Chr(34) & "Server=" & gCurrentServer & ";" & Chr(34) & " & _" & vbCrLf
            strCode = strCode & "       " & Chr(34) & "Database=" & gCurrentDatabase & ";" & Chr(34) & " & _" & vbCrLf
            strCode = strCode & "       " & Chr(34) & "Uid=" & gCurrentUser & ";" & Chr(34) & " & _" & vbCrLf
            strCode = strCode & "       " & Chr(34) & "Pwd=" & gCurrentPassword & ";" & Chr(34) & vbCrLf & vbCrLf
        End If
                
        '-- Create Connection
        If gbCreateConnection = True Then
            If gbCommentCode = True Then strCode = strCode & "   '-- Create ADODB Connection Object" & vbCrLf
            strCode = strCode & "   Dim " & gsConnectionObject & " as New ADODB.Connection" & vbCrLf & vbCrLf
        End If
        
        '-- Bind ConnectionString
        If gbCreateConnectionString = True Then
            If gbCommentCode = True Then strCode = strCode & "   '-- Bind Connection String" & vbCrLf
            strCode = strCode & "   " & gsConnectionObject & ".ConnectionString = " & gsConnectionString & vbCrLf & vbCrLf
        End If
        
        '-- Create Command
        If gbCreateCommand = True Then
            If gbCommentCode = True Then strCode = strCode & "   '-- Create ADODB Command Object" & vbCrLf
            strCode = strCode & "   Dim " & gsCommandObject & " as New ADODB.Command" & vbCrLf & vbCrLf
        End If
        
        '-- Set Command Properties
        If gbCommentCode = True Then strCode = strCode & "   '-- Set properties of Command Object" & vbCrLf
        strCode = strCode & "   With " & gsCommandObject & vbCrLf
        strCode = strCode & "       .ActiveConnection = " & gsConnectionObject & ".ConnectionString " & vbCrLf
        strCode = strCode & "       .CommandText = " & Chr(34) & gCurrentSproc & Chr(34) & vbCrLf
        strCode = strCode & "       .CommandType = " & IIf(gbConstants = True, "adCmdStoredProc", "4") & vbCrLf & vbCrLf
        
        '-- Creat Command Parameters
        If gbCommentCode = True Then strCode = strCode & "       '-- Create ADODB Command Parameters" & vbCrLf
        
        For I = 1 To lv.ListItems.Count
            If lv.ListItems(I).Checked = True Then
                Select Case Mid$(lv.ListItems(I).ListSubItems(2).Text, Len(lv.ListItems(I).ListSubItems(2).Text) - 1, 1)
                    
                    Case 1 '-- Input Parameter
                        strCode = strCode & "       .Parameters.Append " & gsCommandObject & ".CreateParameter(" & Chr(34) & lv.ListItems(I).Text & Chr(34) & ", " & Get_Parameter_Data(lv, I) & ", " & Get_Parameter_Type(lv, I)
                        If IsNumeric(lv.ListItems(I).ListSubItems(7).Text) Then
                            strCode = strCode & ", " & lv.ListItems(I).ListSubItems(7).Text
                        End If
                        strCode = strCode & ")" & vbCrLf
                        strParamCode = strParamCode & "       .Parameters(" & Chr(34) & lv.ListItems(I).Text & Chr(34) & ") = '[ENTER VALUE]" & vbCrLf
                    
                    Case 2 '-- Output Parameter
                        strCode = strCode & "       .Parameters.Append " & gsCommandObject & ".CreateParameter(" & Chr(34) & lv.ListItems(I).Text & Chr(34) & ", " & Get_Parameter_Data(lv, I) & ", " & Get_Parameter_Type(lv, I) & ")" & vbCrLf
                        
                    Case 3 '-- Unknown Parameter
                        strCode = strCode & "       .Parameters.Append " & gsCommandObject & ".CreateParameter(" & Chr(34) & lv.ListItems(I).Text & Chr(34) & ", " & Get_Parameter_Data(lv, I) & ", " & Get_Parameter_Type(lv, I)
                        If IsNumeric(lv.ListItems(I).ListSubItems(7).Text) Then
                            strCode = strCode & ", " & lv.ListItems(I).ListSubItems(7).Text
                        End If
                        strCode = strCode & ")" & vbCrLf
                        strParamCode = strParamCode & "       .Parameters(" & Chr(34) & lv.ListItems(I).Text & Chr(34) & ") = '[ENTER VALUE]" & vbCrLf
                        
                    Case 4 '-- Return Parameter
                        strCode = strCode & "       .Parameters.Append " & gsCommandObject & ".CreateParameter(" & Chr(34) & lv.ListItems(I).Text & Chr(34) & ", " & Get_Parameter_Data(lv, I) & ", " & Get_Parameter_Type(lv, I) & ")" & vbCrLf
                        
                End Select
            End If
        Next
        
        strCode = strCode & vbCrLf
        '-- Set Parameter Values
        If strParamCode <> "" Then
            If gbCommentCode = True Then strCode = strCode & "       '-- Set Parameter Values" & vbCrLf
            strCode = strCode & strParamCode
        End If
        
        strCode = strCode & "   End With" & vbCrLf
        strCode = strCode & vbCrLf
        If gbCommentCode = True Then strCode = strCode & "   '-- Run the stored procedure" & vbCrLf
            
        '-- Run The Sproc
        If gbCreateRecordset Then
            strCode = strCode & "   Dim " & gsRecordsetObject & " as Recordset" & vbCrLf
        End If
            
        If gbRecordsAffected = True Then
            strCode = strCode & "   Dim iRecordsEffected as Integer" & vbCrLf & vbCrLf
        End If
        
        If gbReturnRecordset Then
            If gbRecordsAffected = True Then
                strCode = strCode & "   Set " & gsRecordsetObject & " = " & gsCommandObject & ".Execute(iRecordsEffected)" & vbCrLf & vbCrLf
            End If
        Else
            If gbRecordsAffected = True Then
                strCode = strCode & "   " & gsCommandObject & ".Execute iRecordsEffected" & vbCrLf & vbCrLf
            Else
                strCode = strCode & "   " & gsCommandObject & ".Execute " & vbCrLf & vbCrLf
            End If
        End If
            
        If gbReturnRecordset Then
            If gbCommentCode = True Then strCode = strCode & "   '-- Loop through the recordset" & vbCrLf
            If gbConstants = False Then
                strCode = strCode & "   While Not " & gsRecordsetObject & ".EOF Or Not " & gsRecordsetObject & ".State = 1 & vbCrLf"
            Else
                strCode = strCode & "   While Not " & gsRecordsetObject & ".EOF Or Not " & gsRecordsetObject & ".State = adStateClosed" & vbCrLf
            End If
            strCode = strCode & vbclf & vbCrLf
            strCode = strCode & "       '[ADD CODE HERE]" & vbCrLf
            strCode = strCode & vbclf & vbCrLf
            strCode = strCode & "       " & gsRecordsetObject & ".MoveNext" & vbCrLf
            strCode = strCode & "   Wend" & vbCrLf
            strCode = strCode & vbCrLf
            If gbConstants = False Then
                strCode = strCode & "   If " & gsCommandObject & ".State = 1 Then " & gsRecordsetObject & ".Close" & vbCrLf
            Else
                strCode = strCode & "   If " & gsCommandObject & ".State = adStateOpen Then " & gsRecordsetObject & ".Close" & vbCrLf
            End If
            strCode = strCode & "   Set " & gsRecordsetObject & " = Nothing" & vbCrLf
            strCode = strCode & "   Set " & gsCommandObject & " = Nothing" & vbCrLf
            If gbConstants = True Then
                strCode = strCode & "   If " & gsConnectionObject & ".State = adStateOpen Then " & gsConnectionObject & ".Close" & vbCrLf
            Else
                strCode = strCode & "   If " & gsConnectionObject & ".State = 1 Then " & gsConnectionObject & ".Close" & vbCrLf
            End If
            strCode = strCode & "   Set " & gsConnectionObject & " = Nothing" & vbCrLf
        End If
    
    ElseIf cmd = cmd_asp Then
        
        '-- ASP CODE
        
        '-- Create Connection String
        If gbCreateConnectionString = True Then
            If gbCommentCode = True Then strCode = strCode & "   '-- Create Connection String" & vbCrLf
            strCode = strCode & "   Dim " & gsConnectionString & vbCrLf
            strCode = strCode & "   " & gsConnectionString & " = " & Chr(34) & gSQLDriver & ";" & Chr(34) & " & _" & vbCrLf
            strCode = strCode & "       " & Chr(34) & "Server=" & gCurrentServer & ";" & Chr(34) & " & _" & vbCrLf
            strCode = strCode & "       " & Chr(34) & "Database=" & gCurrentDatabase & ";" & Chr(34) & " & _" & vbCrLf
            strCode = strCode & "       " & Chr(34) & "Uid=" & gCurrentUser & ";" & Chr(34) & " & _" & vbCrLf
            strCode = strCode & "       " & Chr(34) & "Pwd=" & gCurrentPassword & ";" & Chr(34) & vbCrLf & vbCrLf
        End If
                
        '-- Create Connection
        If gbCreateConnection = True Then
            If gbCommentCode = True Then strCode = strCode & "   '-- Create ADODB Connection Object" & vbCrLf
            strCode = strCode & "   Dim " & gsConnectionObject & vbCrLf
            strCode = strCode & "   Set " & gsConnectionObject & " = Server.CreateObject(" & Chr(34) & "ADODB.Connection" & Chr(34) & ")" & vbCrLf & vbCrLf
        End If
        
        '-- Bind ConnectionString
        If gbCreateConnectionString = True Then
            If gbCommentCode = True Then strCode = strCode & "   '-- Bind Connection String" & vbCrLf
            strCode = strCode & "   " & gsConnectionObject & ".ConnectionString = " & gsConnectionString & vbCrLf & vbCrLf
        End If
        
        '-- Create Command
        If gbCreateCommand = True Then
            If gbCommentCode = True Then strCode = strCode & "   '-- Create ADODB Command Object" & vbCrLf
            strCode = strCode & "   Dim " & gsCommandObject & vbCrLf
            strCode = strCode & "   Set " & gsCommandObject & " = Server.CreateObject(" & Chr(34) & "ADODB.Command" & Chr(34) & ")" & vbCrLf & vbCrLf
        End If
        
        '-- Set Command Properties
        If gbCommentCode = True Then strCode = strCode & "   '-- Set properties of Command Object" & vbCrLf
        strCode = strCode & "   With " & gsCommandObject & vbCrLf
        strCode = strCode & "       .ActiveConnection = " & gsConnectionObject & ".ConnectionString " & vbCrLf
        strCode = strCode & "       .CommandText = " & Chr(34) & gCurrentSproc & Chr(34) & vbCrLf
        strCode = strCode & "       .CommandType = " & IIf(gbConstants = True, "adCmdStoredProc", "4") & vbCrLf & vbCrLf
        
        '-- Creat Command Parameters
        If gbCommentCode = True Then strCode = strCode & "       '-- Create ADODB Command Parameters" & vbCrLf
        
        For I = 1 To lv.ListItems.Count
            If lv.ListItems(I).Checked = True Then
                Select Case Mid$(lv.ListItems(I).ListSubItems(2).Text, Len(lv.ListItems(I).ListSubItems(2).Text) - 1, 1)
                    
                    Case 1 '-- Input Parameter
                        strCode = strCode & "       .Parameters.Append " & gsCommandObject & ".CreateParameter(" & Chr(34) & lv.ListItems(I).Text & Chr(34) & ", " & Get_Parameter_Data(lv, I) & ", " & Get_Parameter_Type(lv, I)
                        If IsNumeric(lv.ListItems(I).ListSubItems(7).Text) Then
                            strCode = strCode & ", " & lv.ListItems(I).ListSubItems(7).Text
                        End If
                        strCode = strCode & ")" & vbCrLf
                        strParamCode = strParamCode & "       .Parameters(" & Chr(34) & lv.ListItems(I).Text & Chr(34) & ") = '[ENTER VALUE]" & vbCrLf
                    
                    Case 2 '-- Output Parameter
                        strCode = strCode & "       .Parameters.Append " & gsCommandObject & ".CreateParameter(" & Chr(34) & lv.ListItems(I).Text & Chr(34) & ", " & Get_Parameter_Data(lv, I) & ", " & Get_Parameter_Type(lv, I) & ")" & vbCrLf
                        
                    Case 3 '-- Unknown Parameter
                        strCode = strCode & "       .Parameters.Append " & gsCommandObject & ".CreateParameter(" & Chr(34) & lv.ListItems(I).Text & Chr(34) & ", " & Get_Parameter_Data(lv, I) & ", " & Get_Parameter_Type(lv, I)
                        If IsNumeric(lv.ListItems(I).ListSubItems(7).Text) Then
                            strCode = strCode & ", " & lv.ListItems(I).ListSubItems(7).Text
                        End If
                        strCode = strCode & ")" & vbCrLf
                        strParamCode = strParamCode & "       .Parameters(" & Chr(34) & lv.ListItems(I).Text & Chr(34) & ") = '[ENTER VALUE]" & vbCrLf
                        
                    Case 4 '-- Return Parameter
                        strCode = strCode & "       .Parameters.Append " & gsCommandObject & ".CreateParameter(" & Chr(34) & lv.ListItems(I).Text & Chr(34) & ", " & Get_Parameter_Data(lv, I) & ", " & Get_Parameter_Type(lv, I) & ")" & vbCrLf
                        
                End Select
            End If
        Next
        
        strCode = strCode & vbCrLf
        '-- Set Parameter Values
        If strParamCode <> "" Then
            If gbCommentCode = True Then strCode = strCode & "       '-- Set Parameter Values" & vbCrLf
            strCode = strCode & strParamCode
        End If
        
        strCode = strCode & "   End With" & vbCrLf
        strCode = strCode & vbCrLf
        If gbCommentCode = True Then strCode = strCode & "   '-- Run the stored procedure" & vbCrLf
            
        '-- Run The Sproc
        If gbCreateRecordset Then
            strCode = strCode & "   Dim " & gsRecordsetObject & vbCrLf
        End If
            
        If gbRecordsAffected = True Then
            strCode = strCode & "   Dim iRecordsEffected" & vbCrLf & vbCrLf
        End If
        
        If gbReturnRecordset Then
            If gbRecordsAffected = True Then
                strCode = strCode & "   Set " & gsRecordsetObject & " = " & gsCommandObject & ".Execute(iRecordsEffected)" & vbCrLf & vbCrLf
            End If
        Else
            If gbRecordsAffected = True Then
                strCode = strCode & "   " & gsCommandObject & ".Execute iRecordsEffected" & vbCrLf & vbCrLf
            Else
                strCode = strCode & "   " & gsCommandObject & ".Execute " & vbCrLf & vbCrLf
            End If
        End If
            
        If gbReturnRecordset Then
            If gbCommentCode = True Then strCode = strCode & "   '-- Loop through the recordset" & vbCrLf
            If gbConstants = False Then
                strCode = strCode & "   While Not " & gsRecordsetObject & ".EOF Or Not " & gsRecordsetObject & ".State = 1 & vbCrLf"
            Else
                strCode = strCode & "   While Not " & gsRecordsetObject & ".EOF Or Not " & gsRecordsetObject & ".State = adStateClosed" & vbCrLf
            End If
            strCode = strCode & vbclf & vbCrLf
            strCode = strCode & "       '[ADD CODE HERE]" & vbCrLf
            strCode = strCode & vbclf & vbCrLf
            strCode = strCode & "       " & gsRecordsetObject & ".MoveNext" & vbCrLf
            strCode = strCode & "   Wend" & vbCrLf
            strCode = strCode & vbCrLf
            If gbConstants = False Then
                strCode = strCode & "   If " & gsCommandObject & ".State = 1 Then " & gsRecordsetObject & ".Close" & vbCrLf
            Else
                strCode = strCode & "   If " & gsCommandObject & ".State = adStateOpen Then " & gsRecordsetObject & ".Close" & vbCrLf
            End If
            strCode = strCode & "   Set " & gsRecordsetObject & " = Nothing" & vbCrLf
            strCode = strCode & "   Set " & gsCommandObject & " = Nothing" & vbCrLf
            If gbConstants = True Then
                strCode = strCode & "   If " & gsConnectionObject & ".State = adStateOpen Then " & gsConnectionObject & ".Close" & vbCrLf
            Else
                strCode = strCode & "   If " & gsConnectionObject & ".State = 1 Then " & gsConnectionObject & ".Close" & vbCrLf
            End If
            strCode = strCode & "   Set " & gsConnectionObject & " = Nothing" & vbCrLf
        End If
    
    End If
        
    Colorize rtb, strCode
End Sub

Private Function Get_Parameter_Type(lv As ListView, Row As Integer)
    Dim sText As String
    sText = lv.ListItems(Row).ListSubItems(2).Text
    If gbConstants = False Then
        Get_Parameter_Type = Right(sText, Len(sText) - InStr(1, sText, "("))
        Get_Parameter_Type = Left(Get_Parameter_Type, Len(Get_Parameter_Type) - 1)
    Else
        Get_Parameter_Type = Left(sText, InStr(1, sText, "(") - 1)
    End If
End Function

Private Function Get_Parameter_Data(lv As ListView, Row As Integer)
    Dim sText As String
    sText = lv.ListItems(Row).ListSubItems(6).Text
    If gbConstants = False Then
        Get_Parameter_Data = Right(sText, Len(sText) - InStr(1, sText, "("))
        Get_Parameter_Data = Left(Get_Parameter_Data, Len(Get_Parameter_Data) - 1)
    Else
        Get_Parameter_Data = Left(sText, InStr(1, sText, "(") - 1)
    End If
End Function
