Attribute VB_Name = "basINI"
'/******************************************************************/
'/*說明：　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*放置跟ini讀取、寫入等相關資料的地方。                           */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*模組：　　　　　　　　　　　　　　　    　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*參考：    　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*元件：    　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*注意事項：　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/******************************************************************/
Option Explicit


'/*跟ini處理有關的Win32API常數*/
Public Declare Function ReadINI Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long '讀取ini檔所需要用到的WindowsAPI
'/**/

'/**************************跟ini處理有關的變數***********************************/
Public ReadINI_Return As Integer '用除回傳ini內容時，記錄回傳字串長度
Public ReadINI_String As String * 256 '用於回傳ini內容時，暫用的緩衝字串
'/**************************小華修改的(2009/02/25)***********************************/

''ini讀取程序，若存在inisetup.ref，則改採資料庫讀取方式
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
'        'INI設定檔是否存在
'        If FSO.FileExists(temp$) Or FSO.FileExists(Replace(UCase(temp$), ".REF", ".REX")) Then
'            INIConnection_String = ""
'            If FSO.FileExists(temp$) Then
'                Open temp$ For Binary Access Read As #1
'                    ReDim Inbyte(LOF(1) - 1)
'                    Get #1, , Inbyte
'                Close #1
'                INIConnection_String = StrConv(Inbyte, vbUnicode)
'
'                '加密後輸出
'                If xDecoding(Inbyte(), Outbyte()) Then
'                    '輸出為INX檔
'                    If FSO.FileExists(Replace(UCase(temp$), ".REF", ".REX")) Then
'                        FSO.DeleteFile (Replace(UCase(temp$), ".REF", ".REX"))
'                    End If
'                    Open Replace(UCase(temp$), ".REF", ".REX") For Binary Access Write As #1
'                        Put #1, , Outbyte
'                    Close #1
'                Else
'                    OutLog ("104，設定檔加密錯誤!")
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
'                    OutLog ("108，INI設定檔解密錯誤!")
'                    MsgBox "錯誤代碼108，程式即將自動關閉，請聯絡軒崴工程師處理"
'                    End
'                End If
'            Else
'                If Len(INIConnection_String) > 0 Then
'                    OutLog ("106，無法開啟加密的INI檔，但有連結字串!")
'                Else
'                    OutLog ("107，無法開啟加密的INI檔，且無連結字串!")
'                    MsgBox "錯誤代碼107，程式即將自動關閉，請聯絡軒崴工程師處理"
'                    End
'                End If
'            End If
'
'            '是否有連結字串
'            If Len(INIConnection_String) > 0 Then
'
'                '是否連結字串有效，可以連上資料庫
'                If Connect2DataBase(INIConnection_String, INIConnection) Then
'                    INIIP = GetFullIPAddress()
'                    sql$ = "select * from CRIS_INI_IPCONNECT where IPVALUE = '" & INIIP & "' "
'
'                    '是否可以開啟IP關聯檔
'                    If OpenRecordset(sql$, INIConnection, INIRecordset) Then
'
'                        '是否有此IP的關聯設定檔
'                        If Not INIRecordset.EOF Then
'                            INISetupName = NoNull(INIRecordset("SETUPNAME"))
'                            '檢查有效日期
'                            t1$ = NoNull(INIRecordset("DATEVALID"))
'                            If t1$ < Format(Now, "YYYY/MM/DD") Then
'                                Call DBINILog(INISetupName, INIIP, "202", "此IP已過期")
'                                MsgBox "錯誤代碼202，程式即將自動關閉，請聯絡軒崴工程師處理"
'                                End
'                            End If
'                            '檢查磁碟序號/主機板序號/網卡序號欄位值是否為空
'                            '不為空時，檢查磁碟序號/主機板序號/網卡序號是否相同
'                            t1$ = NoNull(INIRecordset("MBvalue"))
'                            t2$ = NoNull(INIRecordset("NWvalue"))
'                            t3$ = NoNull(INIRecordset("DKvalue"))
'
'                            If t1$ = "" Then
'                                Call DBINILog(INISetupName, INIIP, "301", "此IP尚未填寫主機板序號，已更新")
'                                sql$ = "update CRIS_INI_IPCONNECT set "
'                                sql$ = sql$ & " MBvalue = '" & Replace(Get_MB_SNo, "'", "''") & "' "
'                                sql$ = sql$ & " where IPVALUE = '" & INIIP & "' "
'                                INIConnection.Execute sql$
'                            Else
'                                If t1$ <> Get_MB_SNo Then
'                                    Call DBINILog(INISetupName, INIIP, "203-1", "此IP的主機板序號不對 : " & Get_MB_SNo)
'                                End If
'                            End If
'                            If t2$ = "" Then
'                                Call DBINILog(INISetupName, INIIP, "301", "此IP尚未填寫網卡序號，已更新")
'                                sql$ = "update CRIS_INI_IPCONNECT set "
'                                sql$ = sql$ & " NWvalue = '" & Replace(GetPhysicalAddress, "'", "''") & "' "
'                                sql$ = sql$ & " where IPVALUE = '" & INIIP & "' "
'                                INIConnection.Execute sql$
'                            Else
'                                If t2$ <> GetPhysicalAddress Then
'                                    Call DBINILog(INISetupName, INIIP, "203-2", "此IP的網卡序號不對 : " & GetPhysicalAddress)
'                                End If
'                            End If
'                            If t3$ = "" Then
'                                Call DBINILog(INISetupName, INIIP, "301", "此IP尚未填寫磁碟序號，已更新")
'                                sql$ = "update CRIS_INI_IPCONNECT set "
'                                sql$ = sql$ & " DKvalue = '" & Replace(GetDiskSerialNumber("C:\"), "'", "''") & "' "
'                                sql$ = sql$ & " where IPVALUE = '" & INIIP & "' "
'                                INIConnection.Execute sql$
'                            Else
'                                If t3$ <> GetDiskSerialNumber("C:\") Then
'                                    Call DBINILog(INISetupName, INIIP, "203-3", "此IP的磁碟序號不對 : " & GetDiskSerialNumber("C:\"))
'                                End If
'                            End If
'
'                            '======================================
'                            '讀取資料庫內的INI設定
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
'                                    Call DBINILog(INISetupName, INIIP, "302", "正常連線讀取設定")
'                                    InputINI = zInputINI(ClassName, TitleName, FileName)
'                                Else
'                                    Call DBINILog(INISetupName, INIIP, "206", "查無此IP與程式所用的INI資料項目記錄")
'                                    MsgBox "錯誤代碼206，程式即將自動關閉，請聯絡軒崴工程師處理"
'                                    End
'                                End If
'                            Else
'                                Call DBINILog(INISetupName, INIIP, "205", "開啟INI資料項目檔失敗")
'                                MsgBox "錯誤代碼205，程式即將自動關閉，請聯絡軒崴工程師處理"
'                                End
'                            End If
'                        Else
'                            Call DBINILog("無", INIIP, "201", "查無IP關聯資料!")
''                            OutLog ("201，查無IP關聯資料!")
'                            MsgBox "錯誤代碼201，程式即將自動關閉，請聯絡軒崴工程師處理"
'                            End
'                        End If
'                    Else
'                    '無法開啟IP關聯檔
'                        OutLog ("103，無法開啟INI設定資料表!")
'                        MsgBox "錯誤代碼103，程式即將自動關閉，請聯絡軒崴工程師處理"
'                        End
'                    End If
'                Else
'                '無法連結上資料庫
'                    OutLog ("102，設定檔的連結字串有誤，無法連上資料庫!")
'                    MsgBox "錯誤代碼102，程式即將自動關閉，請聯絡軒崴工程師處理"
'                    End
'                End If
'            Else
'            '無連結字串
'                OutLog ("102，無設定檔的連結字串!")
'                MsgBox "錯誤代碼102，程式即將自動關閉，請聯絡軒崴工程師處理"
'                End
'            End If
'        Else
'        '無INI設定檔時
'            OutLog ("101，未發現INI設定檔!")
'            MsgBox "錯誤代碼101，程式即將自動關閉，請聯絡軒崴工程師處理"
'            End
'        End If
'    Else
'        InputINI = zInputINI(ClassName, TitleName, FileName)
'    End If
'End Function
    
'/************用於比較方便的處理讀ini檔的問題，此函式可以自動去空白跟處理讀取中文ini的問題********/
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

    '暫時封印判讀加密的INX檔功能，等資安模組完成部署再開放20121228
'    If FSO.FileExists(Replace(UCase(FileName), ".INI", ".INX")) Then
'        InputINI = xInputINI(ClassName, TitleName, Replace(UCase(FileName), ".INI", ".INX"))
'    Else
'    If False Then
'        '第一次進入時，讀取INI檔後，加密成INX檔，並上傳資料庫
'        If Not flg_INI_Trans Then
'            '讀取INI檔
'            If FSO.FileExists(FileName) Then
'                Open FileName For Binary Access Read As #1
'                    ReDim Inbyte(LOF(1) - 1)
'                    Get #1, , Inbyte
'                Close #1
'            Else
'                MsgBox "找不到系統設定檔案，請聯絡軒崴工程師!!!"
'                End
'            End If
'
'            xStr = StrConv(Inbyte, vbUnicode)
'            If Not Trans_Ini_Array(xStr) Then
'                MsgBox "系統設定解析錯誤，請聯絡軒崴工程師!!!"
'                End
'            Else
'                '加密
'                If xDecoding(Inbyte(), Outbyte()) Then
'                    '輸出為INX檔
'                    If FSO.FileExists(Replace(UCase(FileName), ".INI", ".INX")) Then
'                        FSO.DeleteFile (Replace(UCase(FileName), ".INI", ".INX"))
'                    End If
'                    Open Replace(UCase(FileName), ".INI", ".INX") For Binary Access Write As #1
'                        Put #1, , Outbyte
'                    Close #1
'                Else
'                    MsgBox "系統設定解密錯誤，請聯絡軒崴工程師!!!"
'                    End
'                End If
'            End If
'
'            '上傳資料庫
'            flg_INI_Trans = True    '先設定旗標，才可遞迴呼叫，不然會陷入死迴圈
'            dbConnection$ = xInputINI("Database", "Connection", FileName)
'            adoDB.Open dbConnection$
'            '是否已有本機資料，若無則上傳；若有則比對資料是否正確
'            Dim LocalRecordset As New adoDB.Recordset
'
'            SQL_String = "select * from cris_decode_main where computer_id = '" & GetIPAddress() & "' "
'            Call OpenRecordset(SQL_String, adoDB, LocalRecordset)
'            If Not LocalRecordset.EOF Then
'                SQL_String = "UPDATE cris_decode_main set MB_SN = '" & Replace(Get_MB_SNo, "/", "") & "', "
'                SQL_String = SQL_String & " DISK_SN = '" & GetDiskSerialNumber("C:\") & ", "
'                SQL_String = SQL_String & " Program_Type = '報告系統' "
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

'比對INI設定的資料項目是否存在，不存在的傳回空字串
Public Function zInputINI(ByVal ClassName As String, ByVal TitleName As String, ByVal FileName As String) As String
    Dim xStr As String
    Dim tf As Boolean, i As Integer
    
    'INI內有設定資料時才查詢
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
        Call DBINILog(INISetupName, INIIP, "204", "無所要讀取的INI資料項目 : " & TitleName)
    End If
End Function

'將LOG訊息寫入本地檔案內(因尚未能連上線前，無法寫入線上資料庫，只能寫在本地檔案)
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
        Print #FreeFilePort, SaveDate, SaveTime, "說明-" & Log_String
    Close #FreeFilePort

End Sub

'2013/06/13
'將LOG訊息寫入cris_ini_log
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
        temp$ = "105，寫入INI的log資料表錯誤!" & vbCrLf
        temp$ = temp$ & SETUPNAME & " / " & IPVALUE & " / " & errType & " / " & errNote & vbCrLf
        temp$ = temp$ & sql$
        OutLog (temp$)
    End If
    
End Sub
