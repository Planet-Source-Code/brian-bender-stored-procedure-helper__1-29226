Attribute VB_Name = "ColorizeCode"
Option Explicit

Public Declare Function LockWindowUpdate Lib "USER32" (ByVal hwndLock As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Const m_strBlueKeyWords = "#Const*#Else*#ElseIf*#End If*#If*Alias*And*As*Base*Binary*Boolean*Byte*ByVal*Call*Case*CBool*CByte*CCur*CDate*CDbl*CDec*CInt*CLng*Close*Compare*Const*CSng*CStr*Currency*CVar*CVErr*Decimal*Declare*DefBool*DefByte*DefCur*DefDate*DefDbl*DefDec*DefInt*DefLng*DefObj*DefSng*DefStr*DefVar*Dim*Do*Double*Each*Else*ElseIf*End*Enum*Eqv*Erase*Error*Exit*Explicit*False*For*Function*Get*Global*GoSub*GoTo*If*Imp*In*Input*Input*Integer*Is*LBound*Let*Lib*Like*Line*Lock*Long*Loop*LSet*Name*New*Next*Not*Object*On*Open*Option*Or*Output*Print*Private*Property*Public*Put*Random*Read*ReDim*Resume*Return*RSet*Seek*Select*Set*Single*Spc*Static*String*Stop*Sub*Tab*Then*Then*True*Type*UBound*Unlock*Variant*Wend*While*With*Xor*Nothing*To*"
Const m_strBlackKeywords = "*Abs*Add*AddItem*AppActivate*Array*Asc*Atn*Beep*Begin*BeginProperty*ChDir*ChDrive*Choose*Chr*Clear*Collection*Command*Cos*CreateObject*CurDir*DateAdd*DateDiff*DatePart*DateSerial*DateValue*Day*DDB*DeleteSetting*Dir*DoEvents*EndProperty*Environ*EOF*Err*Exp*FileAttr*FileCopy*FileDateTime*FileLen*Fix*Format*FV*GetAllSettings*GetAttr*GetObject*GetSetting*Hex*Hide*Hour*InputBox*InStr*Int*Int*IPmt*IRR*IsArray*IsDate*IsEmpty*IsError*IsMissing*IsNull*IsNumeric*IsObject*Item*Kill*LCase*Left*Len*Load*Loc*LOF*Log*LTrim*Me*Mid*Minute*MIRR*MkDir*Month*Now*NPer*NPV*Oct*Pmt*PPmt*PV*QBColor*Raise*Randomize*Rate*Remove*RemoveItem*Reset*RGB*Right*RmDir*Rnd*RTrim*SaveSetting*Second*SendKeys*SetAttr*Sgn*Shell*Sin*Sin*SLN*Space*Sqr*Str*StrComp*StrConv*Switch*SYD*Tan*Text*Time*Time*Timer*TimeSerial*TimeValue*Trim*TypeName*UCase*Unload*Val*VarType*WeekDay*Width*Year*"
Const m_strBlueSQLKeyWords = "*ABSOLUTE*ADD*ALTER*AS*ASC*AT*AUTHORIZATION*BEGIN*BIT*BY*CASCADE*CHAR*CHARACTER*CHECK*CLOSE*COLUMN*COMMIT*CONNECT*CONNECTION*CONSTRAINT*CONTINUE*CREATE*CURRENT*CURRENT_DATE*CURRENT_TIME*CURSOR*DATE*DEALLOCATE*DECIMAL*DECLARE*DEFAULT*DELETE*DESC*DISTINCT*DOUBLE*DROP*ELSE*END*END-EXEC*ESCAPE*EXCEPT*EXEC*EXECUTE*FALSE*FETCH*FIRST*FLOAT*FOR*FOREIGN*FROM*FULL*GLOBAL*GOTO*GRANT*GROUP*HAVING*HOUR*IF*INDEX*INNER*INSENSITIVE*INSERT*INT*INTEGER*INTERSECT*INTO*IS*ISOLATION*KEY*LAST*LEVEL*LOCAL*MAX*MIN*MINUTE*NATIONAL*NCHAR*NEXT*NUMERIC*NVARCHAR*OF*ON*ONLY*OPEN*OPTION*ORDER*PRECISION*PREPARE*PRIMARY*PRIOR*PRIVILEGES*PROC*PROCEDURE*PUBLIC*REFERENCES*RELATIVE*RESTRICT*RETURN**REVOKE*ROLLBACK*ROWS*SCHEMA*SCROLL*SECOND*SECTION*SELECT*SEQUENCE*SET*SIZE*SMALLINT*TABLE*TEMPORARY*THEN*TIMESTAMP*TO*TRANSACTION*TRANSLATION*TRUE*UNION*UNIQUE*UPDATE*VALUES*VARBINARY*VARCHAR*VARYING*VIEW*WHEN*WHERE*WITH*WORK*"
Const m_strGreySQLKeyWords = "*+*-*=*/*(*)*>*<*%*ALL*AND*ANY*BETWEEN*CROSS*EXISTS*IN*JOIN*LIKE*NOT*NULL*OR*OUTER*SOME*"
Const m_strPurpleSQLKeyWords = "*@@ERROR*@@IDENTITY*@@CURSOR_ROWS*@@CPU_BUSY*@@DATEFIRST*@@DBTS*@@FETCH_STATUS*@@IDLE*@@IO_BUSY*@@LANGID*@@LANGUAGE*@@LOCK_TIMEOUT*@@MAX_CONNECTIONS*@@MAX_PRECISION*@@NESTLEVEL*@@OPTIONS*@@PACK_RECEIVED*@@PACK_SENT*@@PACKET_ERRORS*@@PROCID*@@REMSERVER*@@ROWCOUNT*@@SERVERNAME*@@ROWCOUNT*@@SERVERNAME*@@SERVICENAME*@@SPID*@@TEXTSIZE*@@TIMETICKS*@@TOTAL_ERRORS*@@TOTAL_READ*@@TOTAL_WRITE*@@TRANCOUNT*@@VERSION*SIN*COS*ACOS*AVG*CASE*CAST*COALESCE*CONVERT*COUNT*CURRENT_TIMESTAMP*CURRENT_USER*DAY*LEFT*LOWER*MONTH*NULLIF*RIGHT*ROUND*SESSION_USER*SPACE*SUBSTRING*SUM*SYSTEM_USER*UPPER*USER*YEAR*LTRIM*RTRIM*"
Const m_strBlackSQLKeywords = ""

Const m_Colour_SQL_Blue = vbBlue
Const m_Colour_SQL_Grey = &H808080
Const m_Colour_SQL_Purple = &HFF00FF
Const m_Colour_SQL_Comment = &H808000

Const m_Colour_Comment = &H8000&
Const m_Colour_Text = vbBlack
Const m_Colour_Keyword = &H800000

Public rtfTemp As RichTextBox

Public Sub Colorize(rtb As RichTextBox, sText As String, Optional Paste As Boolean = False)
    Dim sBuffer    As String
    Dim sTmpWord   As String
    Dim nStartPos  As Long
    Dim nSelLen    As Long
    Dim nWordPos   As Long
    Dim nBufferlen As Long
    Dim i As Long
    Dim iUpdate As Integer
    Dim sComment As String
    
    gbColorizing = True
    LockWindowUpdate frmMain.rtfTemp.hWnd
    
    frmMain.rtfTemp.Text = sText
    With frmMain.rtfTemp
        .MousePointer = vbHourglass
        .SelStart = 0
        .SelLength = Len(.Text)
        .SelColor = m_Colour_Text
        sBuffer = .Text
        sTmpWord = ""
        nBufferlen = Len(sBuffer)
                
        For i = 1 To nBufferlen + 1
            DoEvents
            Select Case Mid$(sBuffer, i, 1)
                Case "A" To "Z", "a" To "z", "_"
                    If sTmpWord = "" Then nStartPos = i
                    sTmpWord = sTmpWord & Mid(sBuffer, i, 1)
                Case Chr(34) '-- Quote
                    i = InStr(i + 1, sBuffer, Chr(34))
                    If i = 0 Then i = nBufferlen
                Case Chr(39) '-- Apostrophe
CaseComment:
                    
                    .SelStart = i - 1
                    nSelLen = InStr(i, sBuffer, vbCrLf)
                    If nSelLen = 0 Then
                        nSelLen = nBufferlen - i
                    Else
                        nSelLen = nSelLen - i
                    End If
                    .SelLength = nSelLen
                    If InStr(1, .SelText, "[") Then GoTo CaseRequired
                    .SelColor = m_Colour_Comment
                    i = i + nSelLen
                    
CaseRequired:
                Case "["
                    .SelStart = i - 1
                    nSelLen = InStr(i, sBuffer, "]") + 1
                    If nSelLen = 0 Then
                        nSelLen = nBufferlen - i
                    Else
                        nSelLen = nSelLen - i
                    End If
                    .SelLength = nSelLen
                    .SelColor = vbRed
                    i = i + nSelLen
                    
                Case Else
                    If Trim(sTmpWord) <> "" Then
                        .SelStart = nStartPos - 1
                        .SelLength = Len(sTmpWord)
                        nWordPos = InStr(1, m_strBlackKeywords, "*" & sTmpWord & "*", 1)
                        If nWordPos <> 0 Then
                            .SelColor = m_Colour_Text
                            .SelText = Mid$(m_strBlackKeywords, nWordPos + 1, Len(sTmpWord))
                            GoTo ExitSelect
                        End If
                        nWordPos = InStr(1, m_strBlueKeyWords, "*" & sTmpWord & "*", 1)
                        If nWordPos <> 0 Then
                            .SelColor = m_Colour_Keyword
                            .SelText = Mid$(m_strBlueKeyWords, nWordPos + 1, Len(sTmpWord))
                            GoTo ExitSelect
                        End If
                        If UCase(sTmpWord) = "REM" Then
                            i = i - 3
                            GoTo CaseComment
                        End If
                    End If
ExitSelect:
                    sTmpWord = ""
           End Select
        Next
theEnd:
        .SelStart = 0
        .MousePointer = vbDefault
        LockWindowUpdate 0
        rtb.TextRTF = frmMain.rtfTemp.TextRTF
        If Paste Then
            rtb.SetFocus
            SendKeys "{BS}"
        End If
        gbColorizing = False
    End With
End Sub

Public Sub ColorizeSQL(rtb As RichTextBox, sText As String, Optional Paste As Boolean = False)
    Dim sBuffer    As String
    Dim sTmpWord   As String
    Dim nStartPos  As Long
    Dim nSelLen    As Long
    Dim nWordPos   As Long
    Dim nBufferlen As Long
    Dim i As Long
    Dim iUpdate As Integer
    Dim bComment As Boolean
    
    gbColorizing = True
    LockWindowUpdate frmMain.rtfTemp.hWnd
    frmMain.rtfTemp.Text = sText
    With frmMain.rtfTemp
        .MousePointer = vbHourglass
        .SelStart = 0
        .SelLength = Len(.Text)
        .SelColor = m_Colour_Text
        sBuffer = .Text
        sTmpWord = ""
        nBufferlen = Len(sBuffer)
        For i = 1 To nBufferlen
            DoEvents
            Select Case Mid$(sBuffer, i, 1)
                Case "A" To "Z", "a" To "z", "_", "@"
                    If sTmpWord = "" Then nStartPos = i
                    sTmpWord = sTmpWord & Mid(sBuffer, i, 1)
                
                Case Chr(39) '-- Apostrophe
                    .SelStart = i - 1
                    .SelLength = InStr(i + 1, sBuffer, Chr(39)) - (i - 1)
                    .SelColor = vbRed
                    i = InStr(i + 1, sBuffer, Chr(39))
                
                Case "/", "-"
                    If (Mid$(sBuffer, i, 2) = "/*") Or (Mid$(sBuffer, i, 2) = "--") Or bComment Then
CaseComment:
                        If (Mid$(sBuffer, i, 2) = "/*") Then bComment = True
                        .SelStart = i - 1
                        nSelLen = InStr(i, sBuffer, vbCrLf)
                        If nSelLen = 0 Then
                            nSelLen = nBufferlen - i
                        Else
                            nSelLen = nSelLen - i
                        End If
                        .SelLength = nSelLen + 1
                        .SelColor = m_Colour_SQL_Comment
                        i = i + nSelLen
                        If Right(Trim(.SelText), 2) = "*/" Then bComment = False
                    Else
                        GoTo CaseSymbol
                    End If
                
                Case "+", "*", "(", ")", "=", ">", "<", "%"
CaseSymbol:
                    .SelStart = i - 1
                    nSelLen = 1
                    .SelLength = nSelLen
                    .SelColor = m_Colour_SQL_Grey
                    GoTo CaseElse
                
                Case Else
CaseElse:
                    If bComment Then GoTo CaseComment
                    If Trim(sTmpWord) <> "" Then
                        .SelStart = nStartPos - 1
                        .SelLength = Len(sTmpWord)
                        nWordPos = InStr(1, m_strBlackSQLKeywords, "*" & sTmpWord & "*", 1)
                        If nWordPos <> 0 Then
                            .SelColor = m_Colour_Text
                            .SelText = sTmpWord 'Mid$(m_strBlackSQLKeywords, nWordPos + 1, Len(sTmpWord))
                            GoTo ExitSelect
                        End If
                        nWordPos = InStr(1, m_strBlueSQLKeyWords, "*" & sTmpWord & "*", 1)
                        If nWordPos <> 0 Then
                            .SelColor = m_Colour_SQL_Blue
                            .SelText = sTmpWord 'Mid$(m_strBlueSQLKeyWords, nWordPos + 1, Len(sTmpWord))
                            GoTo ExitSelect
                        End If
                        nWordPos = InStr(1, m_strGreySQLKeyWords, "*" & sTmpWord & "*", 1)
                        If nWordPos <> 0 Then
                            .SelColor = m_Colour_SQL_Grey
                            .SelText = sTmpWord 'Mid$(m_strGreySQLKeyWords, nWordPos + 1, Len(sTmpWord))
                            GoTo ExitSelect
                        End If
                        nWordPos = InStr(1, m_strPurpleSQLKeyWords, "*" & sTmpWord & "*", 1)
                        If nWordPos <> 0 Then
                            .SelColor = m_Colour_SQL_Purple
                            .SelText = sTmpWord 'Mid$(m_strPurpleSQLKeyWords, nWordPos + 1, Len(sTmpWord))
                            GoTo ExitSelect
                        End If
                    End If
ExitSelect:
                    sTmpWord = ""
           End Select
        Next
theEnd:
        .SelStart = 0
        .MousePointer = vbDefault
        LockWindowUpdate 0
        rtb.SelRTF = frmMain.rtfTemp.TextRTF
        If Paste Then
            rtb.SetFocus
            SendKeys "{BS}"
        End If
        gbColorizing = False
    End With
End Sub



