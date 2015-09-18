Attribute VB_Name = "Module4"
'
Function GetHISCase(ByVal uni_key As String, ByVal chartno As String) As Boolean
    Dim SQLString As String
    Dim UPDSQLString As String
    Dim tmp$
    Dim xzone As String
    Dim searchcode As String
    
    '20130304修改撈單時，檢查代碼只撈屬於本科室的，所以增加cris_examtype的參照
    searchcode = ""
'    SQLString = "select "
'    SQLString = SQLString & "code "
'    SQLString = SQLString & "from "
'    SQLString = SQLString & "cris_reference where class = 'ExamType' "
    SQLString = "select "
    SQLString = SQLString & "a.code "
    SQLString = SQLString & "from "
    SQLString = SQLString & "cris_reference a, cris_examtype b "
    SQLString = SQLString & "where "
    SQLString = SQLString & "a.type=b.type and b.divisions='" & xDivision_On$ & "' and a.class='ExamType' "
    Call PrintLog(SQLString)
    If OpenRecordset(SQLString, Connection, Recordset) Then
        Do Until Recordset.EOF
            If searchcode <> "" Then
                searchcode = searchcode & ","
            End If
            searchcode = searchcode & "'" & Recordset("code") & "'"
            Recordset.MoveNext
        Loop
    End If

    On Error GoTo errout:
    
    SQLString = "select "
    SQLString = SQLString & "WorkListId, "
    SQLString = SQLString & "WorkListTime, "
    SQLString = SQLString & "TranTime, "
    SQLString = SQLString & "HospitalCode, "
    SQLString = SQLString & "Zone, "
    SQLString = SQLString & "ChartNo, "
    SQLString = SQLString & "PtName, "
    SQLString = SQLString & "IdNo, "
    SQLString = SQLString & "BirthDay, "
    SQLString = SQLString & "Sex, "
    SQLString = SQLString & "BedNo, "
    SQLString = SQLString & "VisitType, "
    SQLString = SQLString & "OrderDoctor, "
    SQLString = SQLString & "OrderTime, "
    SQLString = SQLString & "ItemCode, "
    SQLString = SQLString & "ItemName, "
    SQLString = SQLString & "DepartName, "
    SQLString = SQLString & "MedSummary, "
    SQLString = SQLString & "ReferDoctor, "
    SQLString = SQLString & "AccessionNo, "
    SQLString = SQLString & "ScheduleTime, "
    SQLString = SQLString & "ScheduleModality, "
    SQLString = SQLString & "ScheduleAETitle, "
    SQLString = SQLString & "ScheduleStatus, "
    SQLString = SQLString & "ExecUser, "
    SQLString = SQLString & "ExecTime "
    SQLString = SQLString & "from "
    SQLString = SQLString & "DicomWorkList "
    SQLString = SQLString & "where "

    SQLString = SQLString & "AccessionNo like '%"
'    SQLString = SQLString & "1" & uni_key & "%'"
    SQLString = SQLString & Trim(uni_key) & "%'"
    SQLString = SQLString & " and ChartNo='"
    SQLString = SQLString & chartno
    SQLString = SQLString & "' and ItemCode in ("
    SQLString = SQLString & searchcode
    SQLString = SQLString & ") and (ScheduleModality='US' "
    SQLString = SQLString & "or ScheduleModality='OT') "
    SQLString = SQLString & "order by "
    SQLString = SQLString & "ScheduleModality desc, AccessionNo "
'    SQLString = SQLString & "order by "
'    SQLString = SQLString & "ScheduleModality desc  "
    Call PrintLog(SQLString)
    Call Connect2DataBase(Target_Connection_String, Target_Connection)
    Call OpenRecordset(SQLString, Target_Connection, Target_Recordset)
    If Not Target_Recordset.EOF Then
        If IsNull(Target_Recordset("PtName")) Then
            chartname = ""
        Else
            chartname = Trim(Target_Recordset("PtName"))
        End If
        xzone = NoNull(Target_Recordset("PtName"))
        If IsNull(Target_Recordset("BirthDay")) Then
            chartbirthday = ""
            chartage = ""
        Else
            chartbirthday = Trim(Target_Recordset("BirthDay"))
            If chartbirthday = "" Then
                chartage = ""
            Else
                chartbirthday = Left(chartbirthday, 4) & "/" & Mid(chartbirthday, 5, 2) & "/" & Right(chartbirthday, 2)
                chartage = DateTime.DateDiff("yyyy", chartbirthday, DateTime.Now)
            End If
        End If
        If IsNull(Target_Recordset("BedNo")) Then
            Dr_from = ""
        Else
            Dr_from = Trim(Target_Recordset("BedNo"))
        End If
        If Dr_from = "" Then
            If IsNull(Target_Recordset("VisitType")) Then
                Dr_from = ""
            Else
                Dr_from = Target_Recordset("VisitType")
                Select Case Dr_from
                Case "I"
                    Dr_from = "住院"
                Case "O"
                    Dr_from = "門診"
                Case "E"
                    Dr_from = "急診"
                End Select
            End If
        End If

        If IsNull(Target_Recordset("Sex")) Then
            chartsex = ""
        Else
            chartsex = Trim(Target_Recordset("Sex"))
            Select Case chartsex
            Case "M"
                chartsex = "男"
            Case "F"
                chartsex = "女"
            Case "O"
                chartsex = "中"
            End Select
        End If
        
        If IsNull(Target_Recordset("MedSummary")) Then
            MedSummary = ""
        Else
            MedSummary = Trim(Replace(Target_Recordset("MedSummary"), "'", "''"))
        End If
        If IsNull(Target_Recordset("ExecUser")) Then
            Dr_report = ""
        Else
            Dr_report = Trim(Target_Recordset("ExecUser"))
            
            SQLString = "select "
            SQLString = SQLString & "name "
            SQLString = SQLString & "from "
            SQLString = SQLString & "cris_user "
            SQLString = SQLString & "where "
            SQLString = SQLString & "userid='"
            SQLString = SQLString & Dr_report
            SQLString = SQLString & "' "
            Call PrintLog(SQLString)
            Call OpenRecordset(SQLString, Connection, Recordset)
            If Not Recordset.EOF Then
                Call PrintLog("changed")
                Dr_report = Dr_report & Recordset(0)
            End If
        End If
        If IsNull(Target_Recordset("ExecTime")) Then
            examdate = Format(DateTime.Now, "yyyy/MM/dd")
            examtime = Format(DateTime.Now, "hh:mm:ss")
        Else
            examdate = Trim(Target_Recordset("ExecTime"))
            If examdate = "" Then
                examdate = Format(DateTime.Now, "yyyy/MM/dd")
                examtime = Format(DateTime.Now, "hh:mm:ss")
            Else
                examtime = Mid(examdate, 9, 2) & ":" & Mid(examdate, 11, 2) & ":" & Right(examdate, 2)
                examdate = Left(examdate, 4) & "/" & Mid(examdate, 5, 2) & "/" & Mid(examdate, 7, 2)
            End If
        End If
        
        If IsNull(Target_Recordset("ItemName")) Then
            itemname = ""
        Else
            itemname = Trim(Target_Recordset("ItemName"))
        End If
        
        '如果有多筆記錄時 , 須把不同的檢查細項合併
        itemcode = ""
        While Not Target_Recordset.EOF
            If Not IsNull(Target_Recordset("ItemCode")) Then
                If itemcode <> "" Then
                    itemcode = itemcode & ","
                End If
                itemcode = itemcode & Target_Recordset("ItemCode")
            End If
            Target_Recordset.MoveNext
        Wend
        
        If itemcode <> "" Then
            
            If InStr(itemcode, ",") > 0 Then
                tmp$ = Left(itemcode, InStr(itemcode, ",") - 1)
            Else
                tmp$ = itemcode
            End If

            '如果在cris_reference有對應tpye時，就取代itemname
'                    MsgBox "{" & Left(itemcode, InStr(itemcode, ",")) & "}"
            SQLString = "select "
            SQLString = SQLString & "type "
            SQLString = SQLString & "from "
            SQLString = SQLString & "cris_reference "
            SQLString = SQLString & "where "
            SQLString = SQLString & "code='"
            SQLString = SQLString & tmp$
            SQLString = SQLString & "' "
            Call PrintLog(SQLString)
            Call OpenRecordset(SQLString, Connection, Recordset)
            If Not Recordset.EOF Then
                Call PrintLog("changed itemname")
                itemname = Recordset(0)
            End If
            
        End If
                
        Call PrintLog("ready to change patient name become " & chartname)
        Call PrintLog("ready to change patient sex become " & chartsex)
        Call PrintLog("ready to change patient birthday become " & chartbirthday)
        
        SQLString = "select * from cris_patient_info where chartno='" & chartno & "' "
        Call PrintLog(SQLString)
        Call OpenRecordset(SQLString, Connection, Recordset)
        If Not Recordset.EOF Then
            UPDSQLString = "update "
            UPDSQLString = UPDSQLString & "cris_patient_info "
            UPDSQLString = UPDSQLString & "set "
            UPDSQLString = UPDSQLString & "name='" & chartname & "', "
            UPDSQLString = UPDSQLString & "sex='" & chartsex & "', "
            UPDSQLString = UPDSQLString & "birthday='" & chartbirthday & "' "
            UPDSQLString = UPDSQLString & "where "
            UPDSQLString = UPDSQLString & "chartno='"
            UPDSQLString = UPDSQLString & chartno
            UPDSQLString = UPDSQLString & "' "
        '若無此病歷號時，改為新增
        Else
            '設定enabled為true，以防止USER撈單後才新增病患資料
            frmAddNew.cmdGet.Enabled = True
            UPDSQLString = "INSERT INTO cris_patient_info ("
            UPDSQLString = UPDSQLString & "name, sex, birthday, chartno) VALUES ('" & _
                chartname & "','" & chartsex & "', '" & chartbirthday & "', '" & chartno & "')"
        End If
        Call PrintLog(UPDSQLString)
        Connection.Execute UPDSQLString
        
        UPDSQLString = "update "
        UPDSQLString = UPDSQLString & "cris_exam_online "
        UPDSQLString = UPDSQLString & "set "
        UPDSQLString = UPDSQLString & "system='HIS_IN', "
        UPDSQLString = UPDSQLString & "status='待檢查', "
        UPDSQLString = UPDSQLString & "division_on='" & xDivision_On$ & "', "
        UPDSQLString = UPDSQLString & "examdetail='" & itemcode & "', "
        UPDSQLString = UPDSQLString & "type='" & itemname & "', "
        UPDSQLString = UPDSQLString & "CLINICALIMP='" & MedSummary & "', "
        UPDSQLString = UPDSQLString & "age='" & chartage & "', "
        UPDSQLString = UPDSQLString & "zone='" & xzone & "', "
        UPDSQLString = UPDSQLString & "examdate='" & examdate & "', "
        UPDSQLString = UPDSQLString & "examtime='" & examtime & "', "
        UPDSQLString = UPDSQLString & "orderdate='" & examdate & "', "
        UPDSQLString = UPDSQLString & "ordertime='" & examtime & "', "
        UPDSQLString = UPDSQLString & "dr_report='" & Dr_report & "', "
        UPDSQLString = UPDSQLString & "dr_from='" & Dr_from & "' "
        UPDSQLString = UPDSQLString & "where status<>'已刪除' and "
        UPDSQLString = UPDSQLString & "uni_key='"
        UPDSQLString = UPDSQLString & uni_key
        UPDSQLString = UPDSQLString & "' "
        Call DBRecordLogA(uni_key, chartno, "update", UPDSQLString, "GetHISCase，更新cris_exam_online")
        Connection.Execute UPDSQLString
        
        GetHISCase = True
    Else
        GetHISCase = False
    End If
    
    
    If False Then
errout:
        Call PrintLog("Error at function GetHISCase.")
        GetHISCase = False
    End If
End Function

Sub GetHisSync(ByVal uni_key As String)
    Dim SQLString As String
    Dim UPDSQLString As String
    Dim tmp$
    Dim searchcode As String
    Dim xzone As String
    
    SQLString = "select * from cris_exam_online where uni_key='" & Trim(uni_key) & "' and status<>'已刪除' "
    Call OpenRecordset(SQLString, Connection, Recordset)
    If Recordset.EOF Then
        searchcode = ""
        SQLString = "select "
        SQLString = SQLString & "a.code "
        SQLString = SQLString & "from "
        SQLString = SQLString & "cris_reference a, cris_examtype b "
        SQLString = SQLString & "where "
        SQLString = SQLString & "a.type=b.type and b.divisions='" & xDivision_On$ & "' and a.class='ExamType' "
        Call PrintLog(SQLString)
        If OpenRecordset(SQLString, Connection, Recordset) Then
            Do Until Recordset.EOF
                If searchcode <> "" Then
                    searchcode = searchcode & ","
                End If
                searchcode = searchcode & "'" & Recordset("code") & "'"
                Recordset.MoveNext
            Loop
        End If
        
        On Error GoTo errout:
        
        SQLString = "select "
        SQLString = SQLString & "WorkListId, "
        SQLString = SQLString & "WorkListTime, "
        SQLString = SQLString & "TranTime, "
        SQLString = SQLString & "HospitalCode, "
        SQLString = SQLString & "Zone, "
        SQLString = SQLString & "ChartNo, "
        SQLString = SQLString & "PtName, "
        SQLString = SQLString & "IdNo, "
        SQLString = SQLString & "BirthDay, "
        SQLString = SQLString & "Sex, "
        SQLString = SQLString & "BedNo, "
        SQLString = SQLString & "VisitType, "
        SQLString = SQLString & "OrderDoctor, "
        SQLString = SQLString & "OrderTime, "
        SQLString = SQLString & "ItemCode, "
        SQLString = SQLString & "ItemName, "
        SQLString = SQLString & "DepartName, "
        SQLString = SQLString & "MedSummary, "
        SQLString = SQLString & "ReferDoctor, "
        SQLString = SQLString & "AccessionNo, "
        SQLString = SQLString & "ScheduleTime, "
        SQLString = SQLString & "ScheduleModality, "
        SQLString = SQLString & "ScheduleAETitle, "
        SQLString = SQLString & "ScheduleStatus, "
        SQLString = SQLString & "ExecUser, "
        SQLString = SQLString & "ExecTime "
        SQLString = SQLString & "from "
        SQLString = SQLString & "DicomWorkList "
        SQLString = SQLString & "where "

        SQLString = SQLString & "AccessionNo like '%"
        SQLString = SQLString & Trim(uni_key) & "%'"
        SQLString = SQLString & " and ItemCode in ("
        SQLString = SQLString & searchcode
        SQLString = SQLString & ") and (ScheduleModality='US' "
        SQLString = SQLString & "or ScheduleModality='OT') "
        SQLString = SQLString & "order by "
        SQLString = SQLString & "ScheduleModality desc, AccessionNo "
        Call PrintLog(SQLString)
        Call Connect2DataBase(Target_Connection_String, Target_Connection)
        Call OpenRecordset(SQLString, Target_Connection, Target_Recordset)
        If Not Target_Recordset.EOF Then
            Call PrintLog("uni_key = " & Target_Recordset("AccessionNo"))
            uni_key = Trim(Target_Recordset("AccessionNo"))
            xzone = NoNull(Target_Recordset("Zone"))
            If InStr(uni_key, "-") > 0 Then
                uni_key = Left(uni_key, InStr(uni_key, "-") - 1)
            End If
            uni_key = Right(uni_key, Len(uni_key) - 1)
            
            If IsNull(Target_Recordset("PtName")) Then
                chartname = ""
            Else
                chartname = Trim(Target_Recordset("PtName"))
            End If
            chartno = Target_Recordset("ChartNo")
            If IsNull(Target_Recordset("BirthDay")) Then
                chartbirthday = ""
                chartage = ""
            Else
                chartbirthday = Trim(Target_Recordset("BirthDay"))
                If chartbirthday = "" Then
                    chartage = ""
                Else
                    chartbirthday = Left(chartbirthday, 4) & "/" & Mid(chartbirthday, 5, 2) & "/" & Right(chartbirthday, 2)
                    chartage = DateTime.DateDiff("yyyy", chartbirthday, DateTime.Now)
                End If
            End If
            If IsNull(Target_Recordset("BedNo")) Then
                Dr_from = ""
            Else
                Dr_from = Trim(Target_Recordset("BedNo"))
            End If
            If Dr_from = "" Then
                If IsNull(Target_Recordset("VisitType")) Then
                    Dr_from = ""
                Else
                    Dr_from = Target_Recordset("VisitType")
                    Select Case Dr_from
                    Case "I"
                        Dr_from = "住院"
                    Case "O"
                        Dr_from = "門診"
                    Case "E"
                        Dr_from = "急診"
                    End Select
                End If
            End If
    
            If IsNull(Target_Recordset("Sex")) Then
                chartsex = ""
            Else
                chartsex = Trim(Target_Recordset("Sex"))
                Select Case chartsex
                Case "M"
                    chartsex = "男"
                Case "F"
                    chartsex = "女"
                Case "O"
                    chartsex = "中"
                End Select
            End If
            
            If IsNull(Target_Recordset("MedSummary")) Then
                MedSummary = ""
            Else
                MedSummary = Trim(Replace(Target_Recordset("MedSummary"), "'", "''"))
            End If
            If IsNull(Target_Recordset("ExecUser")) Then
                Dr_report = ""
            Else
                Dr_report = Trim(Target_Recordset("ExecUser"))
                
                SQLString = "select "
                SQLString = SQLString & "name "
                SQLString = SQLString & "from "
                SQLString = SQLString & "cris_user "
                SQLString = SQLString & "where "
                SQLString = SQLString & "userid='"
                SQLString = SQLString & Dr_report
                SQLString = SQLString & "' "
                Call PrintLog(SQLString)
                Call OpenRecordset(SQLString, Connection, Recordset)
                If Not Recordset.EOF Then
                    Call PrintLog("changed")
                    Dr_report = Dr_report & Recordset(0)
                End If
            End If
            If IsNull(Target_Recordset("ExecTime")) Then
                examdate = Format(DateTime.Now, "yyyy/MM/dd")
                examtime = Format(DateTime.Now, "hh:mm:ss")
            Else
                examdate = Trim(Target_Recordset("ExecTime"))
                If examdate = "" Then
                    examdate = Format(DateTime.Now, "yyyy/MM/dd")
                    examtime = Format(DateTime.Now, "hh:mm:ss")
                Else
                    examtime = Mid(examdate, 9, 2) & ":" & Mid(examdate, 11, 2) & ":" & Right(examdate, 2)
                    examdate = Left(examdate, 4) & "/" & Mid(examdate, 5, 2) & "/" & Mid(examdate, 7, 2)
                End If
            End If
            
            If IsNull(Target_Recordset("ItemName")) Then
                itemname = ""
            Else
                itemname = Trim(Target_Recordset("ItemName"))
            End If
            
            '如果有多筆記錄時 , 須把不同的檢查細項合併
            itemcode = ""
            While Not Target_Recordset.EOF
                If Not IsNull(Target_Recordset("ItemCode")) Then
                    If itemcode <> "" Then
                        itemcode = itemcode & ","
                    End If
                    itemcode = itemcode & Target_Recordset("ItemCode")
                End If
                Target_Recordset.MoveNext
            Wend
            
            If itemcode <> "" Then
                
                If InStr(itemcode, ",") > 0 Then
                    tmp$ = Left(itemcode, InStr(itemcode, ",") - 1)
                Else
                    tmp$ = itemcode
                End If
    
                '如果在cris_reference有對應tpye時，就取代itemname
    '                    MsgBox "{" & Left(itemcode, InStr(itemcode, ",")) & "}"
                SQLString = "select "
                SQLString = SQLString & "type "
                SQLString = SQLString & "from "
                SQLString = SQLString & "cris_reference "
                SQLString = SQLString & "where "
                SQLString = SQLString & "code='"
                SQLString = SQLString & tmp$
                SQLString = SQLString & "' "
                Call PrintLog(SQLString)
                Call OpenRecordset(SQLString, Connection, Recordset)
                If Not Recordset.EOF Then
                    Call PrintLog("changed itemname")
                    itemname = Recordset(0)
                End If
                
            End If
                    
            Call PrintLog("ready to change patient name become " & chartname)
            Call PrintLog("ready to change patient sex become " & chartsex)
            Call PrintLog("ready to change patient birthday become " & chartbirthday)
            
            SQLString = "select * from cris_patient_info where chartno='" & chartno & "' "
            Call PrintLog(SQLString)
            Call OpenRecordset(SQLString, Connection, Recordset)
            If Not Recordset.EOF Then
                UPDSQLString = "update "
                UPDSQLString = UPDSQLString & "cris_patient_info "
                UPDSQLString = UPDSQLString & "set "
                UPDSQLString = UPDSQLString & "name='" & chartname & "', "
                UPDSQLString = UPDSQLString & "sex='" & chartsex & "', "
                UPDSQLString = UPDSQLString & "birthday='" & chartbirthday & "' "
                UPDSQLString = UPDSQLString & "where "
                UPDSQLString = UPDSQLString & "chartno='"
                UPDSQLString = UPDSQLString & chartno
                UPDSQLString = UPDSQLString & "' "
            '若無此病歷號時，改為新增
            Else
                '設定enabled為true，以防止USER撈單後才新增病患資料
                frmAddNew.cmdGet.Enabled = True
                UPDSQLString = "INSERT INTO cris_patient_info ("
                UPDSQLString = UPDSQLString & "name, sex, birthday, chartno) VALUES ('" & _
                    chartname & "','" & chartsex & "', '" & chartbirthday & "', '" & chartno & "')"
            End If
            Call PrintLog(UPDSQLString)
            Connection.Execute UPDSQLString
            currForm.txtReqNo.Text = uni_key
            SQLString = "select * from cris_exam_online where uni_key='" & Trim(uni_key) & "' and status<>'已刪除' "
            Call OpenRecordset(SQLString, Connection, Recordset)
            If Recordset.EOF Then
                If Len(examdate) <> 10 Then
                    examdate = Format(Date, "YYYY/MM/DD")
                    examtime = Format(time, "hh:mm:ss")
                End If
                
                UPDSQLString = "Insert into cris_exam_online (uni_key, chartno, "
                UPDSQLString = UPDSQLString & "system, status, division_on, examdetail, "
                UPDSQLString = UPDSQLString & "type, CLINICALIMP, age, "
                UPDSQLString = UPDSQLString & "examdate, examtime, orderdate, "
                UPDSQLString = UPDSQLString & "ordertime, dr_report, dr_from, zone "
                UPDSQLString = UPDSQLString & ") values ('" & uni_key & "', '" & chartno & "', "
                UPDSQLString = UPDSQLString & "'HIS_IN',  '待檢查', '" & xDivision_On$ & "', '" & itemcode & "', "
                UPDSQLString = UPDSQLString & "'" & itemname & "',  '" & MedSummary & "', '" & chartage & "', "
                UPDSQLString = UPDSQLString & "'" & examdate & "', '" & examtime & "', '" & examdate & "',  "
                UPDSQLString = UPDSQLString & "'" & examtime & "', '" & Dr_report & "', '" & Dr_from & "', '" & xzone & "' "
                UPDSQLString = UPDSQLString & ")"
                
                Call DBRecordLogA(uni_key, chartno, "insert", UPDSQLString, "GetHisSync更新cris_exam_online")
                Connection.Execute UPDSQLString
            End If
        End If
    End If
    
    If False Then
errout:
        Call PrintLog("Error at function GetHISSync.")
        
    End If
End Sub
