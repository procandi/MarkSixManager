Attribute VB_Name = "basINI"
'/******************************************************************/
'/*�����G�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*��m��iniŪ���B�g�J��������ƪ��a��C                           */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�ҲաG�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@    �@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�ѦҡG    �@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*����G    �@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�`�N�ƶ��G�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/******************************************************************/
Option Explicit


'/*��ini�B�z������Win32API�`��*/
Public Declare Function ReadINI Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long 'Ū��ini�ɩһݭn�Ψ쪺WindowsAPI
'/**/

'/**************************��ini�B�z�������ܼ�***********************************/
Public ReadINI_Return As Integer '�ΰ��^��ini���e�ɡA�O���^�Ǧr�����
Public ReadINI_String As String * 256 '�Ω�^��ini���e�ɡA�ȥΪ��w�Ħr��
'/**************************�p�حק諸(2009/02/25)***********************************/

''iniŪ���{�ǡA�Y�s�binisetup.ref�A�h��ĸ�ƮwŪ���覡
'Public Function InputINI(ByVal ClassName As String, ByVal TitleName As String, ByVal FileName As String) As String
'    Dim temp$, sql$, t1$, t2$, t3$
'    Dim Inbyte() As Byte
'    Dim Outbyte() As Byte
'
'    If Not flg_INI_Trans Then
'        flg_INI_Trans = True
''        flg_IS_INI_Trans = False
'        INISetupName = ""
'        temp$ = App.Path & "\inisetup.ref"
'
'        'INI�]�w�ɬO�_�s�b
'        If FSO.FileExists(temp$) Or FSO.FileExists(Replace(UCase(temp$), ".REF", ".REX")) Then
'            INIConnection_String = ""
'            If FSO.FileExists(temp$) Then
'                Open temp$ For Binary Access Read As #1
'                    ReDim Inbyte(LOF(1) - 1)
'                    Get #1, , Inbyte
'                Close #1
'                INIConnection_String = StrConv(Inbyte, vbUnicode)
'
'                '�[�K���X
'                If xDecoding(Inbyte(), Outbyte()) Then
'                    '��X��INX��
'                    If FSO.FileExists(Replace(UCase(temp$), ".REF", ".REX")) Then
'                        FSO.DeleteFile (Replace(UCase(temp$), ".REF", ".REX"))
'                    End If
'                    Open Replace(UCase(temp$), ".REF", ".REX") For Binary Access Write As #1
'                        Put #1, , Outbyte
'                    Close #1
'                Else
'                    OutLog ("104�A�]�w�ɥ[�K���~!")
'                End If
'            End If
'            temp$ = Replace(UCase(temp$), ".REF", ".REX")
'            If FSO.FileExists(temp$) Then
'                Open temp$ For Binary Access Read As #1
'                    ReDim Inbyte(LOF(1) - 1)
'                    Get #1, , Inbyte
'                Close #1
'                If xDecoding(Inbyte(), Outbyte()) Then
'                    INIConnection_String = StrConv(Outbyte, vbUnicode)
'                Else
'                    OutLog ("108�AINI�]�w�ɸѱK���~!")
'                    MsgBox "���~�N�X108�A�{���Y�N�۰������A���p���a�Q�u�{�v�B�z"
'                    End
'                End If
'            Else
'                If Len(INIConnection_String) > 0 Then
'                    OutLog ("106�A�L�k�}�ҥ[�K��INI�ɡA�����s���r��!")
'                Else
'                    OutLog ("107�A�L�k�}�ҥ[�K��INI�ɡA�B�L�s���r��!")
'                    MsgBox "���~�N�X107�A�{���Y�N�۰������A���p���a�Q�u�{�v�B�z"
'                    End
'                End If
'            End If
'
'            '�O�_���s���r��
'            If Len(INIConnection_String) > 0 Then
'
'                '�O�_�s���r�꦳�ġA�i�H�s�W��Ʈw
'                If Connect2DataBase(INIConnection_String, INIConnection) Then
'                    INIIP = GetFullIPAddress()
'                    sql$ = "select * from CRIS_INI_IPCONNECT where IPVALUE = '" & INIIP & "' "
'
'                    '�O�_�i�H�}��IP���p��
'                    If OpenRecordset(sql$, INIConnection, INIRecordset) Then
'
'                        '�O�_����IP�����p�]�w��
'                        If Not INIRecordset.EOF Then
'                            INISetupName = NoNull(INIRecordset("SETUPNAME"))
'                            '�ˬd���Ĥ��
'                            t1$ = NoNull(INIRecordset("DATEVALID"))
'                            If t1$ < Format(Now, "YYYY/MM/DD") Then
'                                Call DBINILog(INISetupName, INIIP, "202", "��IP�w�L��")
'                                MsgBox "���~�N�X202�A�{���Y�N�۰������A���p���a�Q�u�{�v�B�z"
'                                End
'                            End If
'                            '�ˬd�ϺЧǸ�/�D���O�Ǹ�/���d�Ǹ����ȬO�_����
'                            '�����ŮɡA�ˬd�ϺЧǸ�/�D���O�Ǹ�/���d�Ǹ��O�_�ۦP
'                            t1$ = NoNull(INIRecordset("MBvalue"))
'                            t2$ = NoNull(INIRecordset("NWvalue"))
'                            t3$ = NoNull(INIRecordset("DKvalue"))
'
'                            If t1$ = "" Then
'                                Call DBINILog(INISetupName, INIIP, "301", "��IP�|����g�D���O�Ǹ��A�w��s")
'                                sql$ = "update CRIS_INI_IPCONNECT set "
'                                sql$ = sql$ & " MBvalue = '" & Replace(Get_MB_SNo, "'", "''") & "' "
'                                sql$ = sql$ & " where IPVALUE = '" & INIIP & "' "
'                                INIConnection.Execute sql$
'                            Else
'                                If t1$ <> Get_MB_SNo Then
'                                    Call DBINILog(INISetupName, INIIP, "203-1", "��IP���D���O�Ǹ����� : " & Get_MB_SNo)
'                                End If
'                            End If
'                            If t2$ = "" Then
'                                Call DBINILog(INISetupName, INIIP, "301", "��IP�|����g���d�Ǹ��A�w��s")
'                                sql$ = "update CRIS_INI_IPCONNECT set "
'                                sql$ = sql$ & " NWvalue = '" & Replace(GetPhysicalAddress, "'", "''") & "' "
'                                sql$ = sql$ & " where IPVALUE = '" & INIIP & "' "
'                                INIConnection.Execute sql$
'                            Else
'                                If t2$ <> GetPhysicalAddress Then
'                                    Call DBINILog(INISetupName, INIIP, "203-2", "��IP�����d�Ǹ����� : " & GetPhysicalAddress)
'                                End If
'                            End If
'                            If t3$ = "" Then
'                                Call DBINILog(INISetupName, INIIP, "301", "��IP�|����g�ϺЧǸ��A�w��s")
'                                sql$ = "update CRIS_INI_IPCONNECT set "
'                                sql$ = sql$ & " DKvalue = '" & Replace(GetDiskSerialNumber("C:\"), "'", "''") & "' "
'                                sql$ = sql$ & " where IPVALUE = '" & INIIP & "' "
'                                INIConnection.Execute sql$
'                            Else
'                                If t3$ <> GetDiskSerialNumber("C:\") Then
'                                    Call DBINILog(INISetupName, INIIP, "203-3", "��IP���ϺЧǸ����� : " & GetDiskSerialNumber("C:\"))
'                                End If
'                            End If
'
'                            '======================================
'                            'Ū����Ʈw����INI�]�w
'                            '======================================
'                            sql$ = " select * from CRIS_INI_ITEMS "
'                            sql$ = sql$ & " where SETUPNAME = '" & INISetupName & "' and ProgramName = '" & ProgramName & "' "
'                            If OpenRecordset(sql$, INIConnection, INIRecordset) Then
'                                If Not INIRecordset.EOF Then
'                                    Ini_Variables_Count = 0
'                                    While Not INIRecordset.EOF
'                                        Ini_Variables_Name(Ini_Variables_Count, 0) = ProgramName
'                                        Ini_Variables_Name(Ini_Variables_Count, 1) = UCase(NoNull(INIRecordset("ITEMNAME")))
'                                        Ini_Variables_Name(Ini_Variables_Count, 2) = NoNull(INIRecordset("ITEMVALUE"))
'                                        Ini_Variables_Count = Ini_Variables_Count + 1
'                                        INIRecordset.MoveNext
'                                    Wend
'                                    Call DBINILog(INISetupName, INIIP, "302", "���`�s�uŪ���]�w")
'                                    InputINI = zInputINI(ClassName, TitleName, FileName)
'                                Else
'                                    Call DBINILog(INISetupName, INIIP, "206", "�d�L��IP�P�{���ҥΪ�INI��ƶ��ذO��")
'                                    MsgBox "���~�N�X206�A�{���Y�N�۰������A���p���a�Q�u�{�v�B�z"
'                                    End
'                                End If
'                            Else
'                                Call DBINILog(INISetupName, INIIP, "205", "�}��INI��ƶ����ɥ���")
'                                MsgBox "���~�N�X205�A�{���Y�N�۰������A���p���a�Q�u�{�v�B�z"
'                                End
'                            End If
'                        Else
'                            Call DBINILog("�L", INIIP, "201", "�d�LIP���p���!")
''                            OutLog ("201�A�d�LIP���p���!")
'                            MsgBox "���~�N�X201�A�{���Y�N�۰������A���p���a�Q�u�{�v�B�z"
'                            End
'                        End If
'                    Else
'                    '�L�k�}��IP���p��
'                        OutLog ("103�A�L�k�}��INI�]�w��ƪ�!")
'                        MsgBox "���~�N�X103�A�{���Y�N�۰������A���p���a�Q�u�{�v�B�z"
'                        End
'                    End If
'                Else
'                '�L�k�s���W��Ʈw
'                    OutLog ("102�A�]�w�ɪ��s���r�꦳�~�A�L�k�s�W��Ʈw!")
'                    MsgBox "���~�N�X102�A�{���Y�N�۰������A���p���a�Q�u�{�v�B�z"
'                    End
'                End If
'            Else
'            '�L�s���r��
'                OutLog ("102�A�L�]�w�ɪ��s���r��!")
'                MsgBox "���~�N�X102�A�{���Y�N�۰������A���p���a�Q�u�{�v�B�z"
'                End
'            End If
'        Else
'        '�LINI�]�w�ɮ�
'            OutLog ("101�A���o�{INI�]�w��!")
'            MsgBox "���~�N�X101�A�{���Y�N�۰������A���p���a�Q�u�{�v�B�z"
'            End
'        End If
'    Else
'        InputINI = zInputINI(ClassName, TitleName, FileName)
'    End If
'End Function
    
'/************�Ω�����K���B�zŪini�ɪ����D�A���禡�i�H�۰ʥh�ťո�B�zŪ������ini�����D********/
Public Function InputINI(ByVal ClassName As String, ByVal TitleName As String, ByVal FileName As String) As String
    Dim InputINI_Return As Integer
    Dim InputINI_String As String * 256
    Dim Result_String As String
    Dim Inbyte() As Byte
    Dim Outbyte() As Byte
    Dim xStr As String
    Dim adoDB As New adoDB.Connection
    Dim adoRS As New adoDB.Recordset
    Dim tSql$
    Dim SQL_String As String

    '�ȮɫʦL�PŪ�[�K��INX�ɥ\��A����w�Ҳէ������p�A�}��20121228
'    If FSO.FileExists(Replace(UCase(FileName), ".INI", ".INX")) Then
'        InputINI = xInputINI(ClassName, TitleName, Replace(UCase(FileName), ".INI", ".INX"))
'    Else
'    If False Then
'        '�Ĥ@���i�J�ɡAŪ��INI�ɫ�A�[�K��INX�ɡA�äW�Ǹ�Ʈw
'        If Not flg_INI_Trans Then
'            'Ū��INI��
'            If FSO.FileExists(FileName) Then
'                Open FileName For Binary Access Read As #1
'                    ReDim Inbyte(LOF(1) - 1)
'                    Get #1, , Inbyte
'                Close #1
'            Else
'                MsgBox "�䤣��t�γ]�w�ɮסA���p���a�Q�u�{�v!!!"
'                End
'            End If
'
'            xStr = StrConv(Inbyte, vbUnicode)
'            If Not Trans_Ini_Array(xStr) Then
'                MsgBox "�t�γ]�w�ѪR���~�A���p���a�Q�u�{�v!!!"
'                End
'            Else
'                '�[�K
'                If xDecoding(Inbyte(), Outbyte()) Then
'                    '��X��INX��
'                    If FSO.FileExists(Replace(UCase(FileName), ".INI", ".INX")) Then
'                        FSO.DeleteFile (Replace(UCase(FileName), ".INI", ".INX"))
'                    End If
'                    Open Replace(UCase(FileName), ".INI", ".INX") For Binary Access Write As #1
'                        Put #1, , Outbyte
'                    Close #1
'                Else
'                    MsgBox "�t�γ]�w�ѱK���~�A���p���a�Q�u�{�v!!!"
'                    End
'                End If
'            End If
'
'            '�W�Ǹ�Ʈw
'            flg_INI_Trans = True    '���]�w�X�СA�~�i���j�I�s�A���M�|���J���j��
'            dbConnection$ = xInputINI("Database", "Connection", FileName)
'            adoDB.Open dbConnection$
'            '�O�_�w��������ơA�Y�L�h�W�ǡF�Y���h����ƬO�_���T
'            Dim LocalRecordset As New adoDB.Recordset
'
'            SQL_String = "select * from cris_decode_main where computer_id = '" & GetIPAddress() & "' "
'            Call OpenRecordset(SQL_String, adoDB, LocalRecordset)
'            If Not LocalRecordset.EOF Then
'                SQL_String = "UPDATE cris_decode_main set MB_SN = '" & Replace(Get_MB_SNo, "/", "") & "', "
'                SQL_String = SQL_String & " DISK_SN = '" & GetDiskSerialNumber("C:\") & ", "
'                SQL_String = SQL_String & " Program_Type = '���i�t��' "
'                SQL_String = SQL_String & " where computer_id = '" & GetIPAddress() & "' "
'            Else
'                SQL_String = "INSERT cris_decode_main "
'            End If
'
'            adoDB.Close
'            Set adoDB = Nothing
'        End If
'    End If
        InputINI_Return = ReadINI(ClassName, TitleName, "", InputINI_String, Len(InputINI_String), FileName)
        Result_String = Trim(Left(InputINI_String, InputINI_Return))

        If Len(Result_String) < 1 Then
            InputINI = ""
        Else
            Do While Len(Result_String) > 0 And (Asc(Right(Result_String, 1)) = 32 Or Asc(Right(Result_String, 1)) = 0 Or Asc(Right(Result_String, 1)) = 77 Or Asc(Right(Result_String, 1)) = 121)
                Result_String = Left(Result_String, Len(Result_String) - 1)
            Loop

            InputINI = Result_String
        End If
'    End If
End Function

'���INI�]�w����ƶ��جO�_�s�b�A���s�b���Ǧ^�Ŧr��
Public Function zInputINI(ByVal ClassName As String, ByVal TitleName As String, ByVal FileName As String) As String
    Dim xStr As String
    Dim tf As Boolean, i As Integer
    
    'INI�����]�w��Ʈɤ~�d��
    xStr = ""
    tf = False
    If Ini_Variables_Count > 0 Then
        For i = 0 To Ini_Variables_Count
            If UCase(TitleName) = Ini_Variables_Name(i, 1) Then
                xStr = Ini_Variables_Name(i, 2)
                tf = True
                Exit For
            End If
        Next
    End If
    zInputINI = xStr
    If Not tf Then
        Call DBINILog(INISetupName, INIIP, "204", "�L�ҭnŪ����INI��ƶ��� : " & TitleName)
    End If
End Function

'�NLOG�T���g�J���a�ɮפ�(�]�|����s�W�u�e�A�L�k�g�J�u�W��Ʈw�A�u��g�b���a�ɮ�)
Public Sub OutLog(ByRef Log_String As String)
    Dim FSO_FileExist As New FileSystemObject
    Dim SavePath As String
    Dim SaveFile As String
    Dim SaveDate As String
    Dim SaveTime As String
    
    
    SavePath = App.Path & "\log\"
    SaveFile = Format(DateTime.Date, "yyyyMMdd") & ".log"
    SaveDate = DateTime.Date
    SaveTime = DateTime.time
    
    
    If Not FSO_FileExist.FolderExists(SavePath) Then
        Call CreatePath(SavePath)
    End If
    

    FreeFilePort = FreeFile
    Open SavePath & SaveFile For Append As #FreeFilePort
        Print #FreeFilePort, SaveDate, SaveTime, "����-" & Log_String
    Close #FreeFilePort

End Sub

'2013/06/13
'�NLOG�T���g�Jcris_ini_log
Public Sub DBINILog(SETUPNAME, IPVALUE, errType, errNote)
    Dim sql$, temp$
    Dim adoDB As adoDB.Connection
    
    Set adoDB = New adoDB.Connection
    adoDB.Open INIConnection_String
    
    On Error GoTo err
        sql$ = "insert into cris_INI_Log ( "
        sql$ = sql$ & "SETUPNAME, IPVALUE, errType, errNote, xDate, xTime, ProgramName "
        sql$ = sql$ & " ) values ( "
        sql$ = sql$ & " '" & SETUPNAME & "', '" & IPVALUE & "', "
        sql$ = sql$ & " '" & errType & "', '" & Replace(errNote, "'", "''") & "', "
        sql$ = sql$ & " '" & Format(Now, "YYYY/MM/DD") & "', '" & Format(Now, "hh:mm:ss") & "', "
        sql$ = sql$ & " '" & ProgramName & "' )"
        adoDB.Execute sql$
        adoDB.Close
        Set adoDB = Nothing
    
    If False Then
err:
        temp$ = "105�A�g�JINI��log��ƪ���~!" & vbCrLf
        temp$ = temp$ & SETUPNAME & " / " & IPVALUE & " / " & errType & " / " & errNote & vbCrLf
        temp$ = temp$ & sql$
        OutLog (temp$)
    End If
    
End Sub
