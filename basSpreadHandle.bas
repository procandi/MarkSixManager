Attribute VB_Name = "basSpreadHandle"
 
 Public Function Spread2Text(ByVal uni_key As String, ByVal chartno As String, ByRef SpreadForm As fpSpread) As String
 Dim ReportTemp As String
    ReportTemp = "[Indication]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Indication#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Pre-medication]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Pre-medication#" & vbCrLf
 
    ReportTemp = ReportTemp & "[Colon cleansing agent]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#cleansingagent#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Colon cleansing level]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#cleansinglevel#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Insertion Level]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Inserttion#" & vbCrLf
    
    '因尚不知這個內容的來源，所以暫時填空白
    ReportTemp = ReportTemp & "[Degree of difficulty]" & vbCrLf
'    ReportTemp = ReportTemp & "  " & "#Degree#" & vbCrLf
    ReportTemp = ReportTemp & "  " & "Nil" & vbCrLf
    
    ReportTemp = ReportTemp & "[Endoscopic finding]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#finding#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Diagnosis]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Diagnosis#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Limitation of examination]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#examination#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Complication]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Complication#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Suggestion]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Suggestion#" & vbCrLf
    
    
    For i = 1 To SpreadForm.MaxRows
        SpreadForm.row = i
        For j = 1 To SpreadForm.MaxCols
        
            If i = 3 Or i = 4 Or i = 5 Or i = 6 Then
                Exit For '/因為3~4~5都有checkbox 要算所以跳過
            End If
            SpreadForm.Col = j
            
            If i = 2 Then
                '#Indication位置
                ReportTemp = Replace(ReportTemp, "#Indication#", GetCheckBox(SpreadForm, 6))
            ElseIf i = 8 And SpreadForm.CellType = CellTypeComboBox Then
                '#Pre-medication#位置
'                Dim ColonPre As String
'                ColonPre = "1.Gascon drop oral 10-15c.c." & vbCrLf
'                ColonPre = ColonPre & "2.Xylocaine(10%) pump spray 1-3puffs"
'                If UCase(Trim(SpreadForm.Text)) <> "NIL" And Trim(SpreadForm.Text <> "") Then
'                   ColonPre = ColonPre & vbCrLf & "3." & SpreadForm.Text
'                End If
                Call ReplaceReport("#Pre-medication#", SpreadForm.Text, ReportTemp)
                
            ElseIf i = 9 And SpreadForm.CellType = CellTypeComboBox Then
                '#cleansingagent#
                Call ReplaceReport("#cleansingagent#", SpreadForm.Text, ReportTemp)
                
            ElseIf i = 10 And SpreadForm.CellType = CellTypeComboBox Then
                '#cleansinglevel#
                Call ReplaceReport("#cleansinglevel#", SpreadForm.Text, ReportTemp)
            ElseIf i = 11 And SpreadForm.CellType = CellTypeComboBox Then
                '#Inserttion#
                Call ReplaceReport("#Inserttion#", SpreadForm.Text, ReportTemp)
            ElseIf i = 13 And SpreadForm.CellType = CellTypeEdit Then
                '#finding#
                Call ReplaceReport("#finding#", SpreadForm.Text, ReportTemp)
            ElseIf i = 16 And SpreadForm.CellType = CellTypeEdit Then
                '#Diagnosis#
                Call ReplaceReport("#Diagnosis#", SpreadForm.Text, ReportTemp)
            ElseIf i = 20 And SpreadForm.CellType = CellTypeComboBox Then
                '#examination#
                Call ReplaceReport("#examination#", SpreadForm.Text, ReportTemp)
            ElseIf i = 21 And SpreadForm.CellType = CellTypeComboBox Then
                '#Complication#
                Call ReplaceReport("#Complication#", SpreadForm.Text, ReportTemp)
            ElseIf i = 22 And SpreadForm.CellType = CellTypeEdit Then
                '#Complication#
                Call ReplaceReport("#Suggestion#", SpreadForm.Text, ReportTemp)
            End If
        Next
    Next
'    Call UpdateExamOnline(uni_key, Chartno, ReportTemp)
    Spread2Text = Trim(ReportTemp)
 End Function
 
 '/20130913大腸鏡報表(腸胃科)修改
 Public Function Spread2Text_1(ByVal uni_key As String, ByVal chartno As String, ByRef SpreadForm As fpSpread) As String
 Dim ReportTemp As String
    ReportTemp = "[Indication]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Indication#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Pre-medication]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Pre-medication#" & vbCrLf
 
    ReportTemp = ReportTemp & "[Colon cleansing agent]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#cleansingagent#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Colon cleansing level]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#cleansinglevel#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Insertion Level]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Inserttion#" & vbCrLf
    
    '因尚不知這個內容的來源，所以暫時填空白
    ReportTemp = ReportTemp & "[Degree of difficulty]" & vbCrLf
'    ReportTemp = ReportTemp & "  " & "#Degree#" & vbCrLf
    ReportTemp = ReportTemp & "  " & "Nil" & vbCrLf
    
    ReportTemp = ReportTemp & "[Endoscopic finding]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#finding#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Diagnosis]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Diagnosis#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Limitation of examination]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#examination#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Complication]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Complication#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Suggestion]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Suggestion#" & vbCrLf
    
    For i = 1 To SpreadForm.MaxRows
        SpreadForm.row = i
        For j = 1 To SpreadForm.MaxCols
        
            If i = 3 Or i = 4 Or i = 5 Or i = 6 Then
                Exit For '/因為3~4~5都有checkbox 要算所以跳過
            End If
            SpreadForm.Col = j
            
            If i = 2 Then
                '#Indication位置
                ReportTemp = Replace(ReportTemp, "#Indication#", GetCheckBox(SpreadForm, 6))
            ElseIf i = 8 And SpreadForm.CellType = CellTypeComboBox Then
                Dim ColPre As String
                '/記錄三個pre-medication
                Dim Pr1  As String, Pr2 As String, Pr3 As String
                '/記錄pre_medication 得編號,從3開始
                Dim PrCount As Integer
                PrCount = 1
                Dim l As Integer, m As Integer
                For l = 8 To 10
                    SpreadForm.row = l
                    For m = 1 To SpreadForm.MaxCols
                        SpreadForm.Col = m
                        If SpreadForm.CellType = CellTypeComboBox Then
                            If UCase(Trim(SpreadForm.Text)) <> "NIL" And Trim(SpreadForm.Text <> "") Then
                                If ColPre = "" Then
                                    ColPre = PrCount & "." & SpreadForm.Text
                                Else
                                    ColPre = ColPre & vbCrLf & PrCount & "." & SpreadForm.Text
                                End If
                                PrCount = PrCount + 1
                            End If
                            Exit For
                        End If
                    Next
                Next
                Call ReplaceReport("#Pre-medication#", ColPre, ReportTemp)
                
            ElseIf i = 12 And SpreadForm.CellType = CellTypeComboBox Then
                '#cleansingagent#
                Call ReplaceReport("#cleansingagent#", SpreadForm.Text, ReportTemp)
                
            ElseIf i = 13 And SpreadForm.CellType = CellTypeComboBox Then
                '#cleansinglevel#
                Call ReplaceReport("#cleansinglevel#", SpreadForm.Text, ReportTemp)
            ElseIf i = 14 And SpreadForm.CellType = CellTypeComboBox Then
                '#Inserttion#
                Call ReplaceReport("#Inserttion#", SpreadForm.Text, ReportTemp)
            ElseIf i = 16 And SpreadForm.CellType = CellTypeEdit Then
                '#finding#
                Call ReplaceReport("#finding#", SpreadForm.Text, ReportTemp)
            ElseIf i = 19 And SpreadForm.CellType = CellTypeEdit Then
                '#Diagnosis#
                Call ReplaceReport("#Diagnosis#", SpreadForm.Text, ReportTemp)
            ElseIf i = 23 And SpreadForm.CellType = CellTypeComboBox Then
                '#examination#
                Call ReplaceReport("#examination#", SpreadForm.Text, ReportTemp)
            ElseIf i = 24 And SpreadForm.CellType = CellTypeComboBox Then
                '#Complication#
                Call ReplaceReport("#Complication#", SpreadForm.Text, ReportTemp)
            ElseIf i = 25 And SpreadForm.CellType = CellTypeEdit Then
                '#Complication#
                Call ReplaceReport("#Suggestion#", SpreadForm.Text, ReportTemp)
            End If
        Next
    Next
'    Call UpdateExamOnline(uni_key, Chartno, ReportTemp)
    Spread2Text_1 = Trim(ReportTemp)
 End Function
 
 
 
 '/20130902將大腸直腸外科的colon_out_1報表拆成三個頁籤。
 Public Function Spread2Text_ColonOut_1(ByVal uni_key As String, ByVal chartno As String, ByRef SpreadForm As fpSpread) As String
 Dim ReportTemp As String
    ReportTemp = "[Indication]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Indication#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Pre-medication]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Pre-medication#" & vbCrLf
 
    ReportTemp = ReportTemp & "[Colon cleansing agent]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#cleansingagent#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Colon cleansing level]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#cleansinglevel#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Insertion Level]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Inserttion#" & vbCrLf
    
    '因尚不知這個內容的來源，所以暫時填空白
    ReportTemp = ReportTemp & "[Degree of difficulty]" & vbCrLf
'    ReportTemp = ReportTemp & "  " & "#Degree#" & vbCrLf
    ReportTemp = ReportTemp & "  " & "Nil" & vbCrLf
    
    ReportTemp = ReportTemp & "[Endoscopic finding]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#finding#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Diagnosis]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Diagnosis#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Limitation of examination]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#examination#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Complication]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Complication#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Suggestion]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Suggestion#" & vbCrLf
    
    '/頁籤(1)Main
    SpreadForm.sheet = 1
    For i = 1 To SpreadForm.MaxRows
        SpreadForm.row = i
        For j = 1 To SpreadForm.MaxCols
        
            If i = 3 Or i = 4 Or i = 5 Or i = 6 Then
                Exit For '/因為3~4~5都有checkbox 要算所以跳過
            End If
            SpreadForm.Col = j
            
            If i = 2 Then
                '#Indication位置
                ReportTemp = Replace(ReportTemp, "#Indication#", GetCheckBox(SpreadForm, 6))
            ElseIf i = 8 And SpreadForm.CellType = CellTypeComboBox Then
                '#Pre-medication#位置
'                Dim ColonPre As String
'                ColonPre = "1.Gascon drop oral 10-15c.c." & vbCrLf
'                ColonPre = ColonPre & "2.Xylocaine(10%) pump spray 1-3puffs"
'                If UCase(Trim(SpreadForm.Text)) <> "NIL" And Trim(SpreadForm.Text <> "") Then
'                   ColonPre = ColonPre & vbCrLf & "3." & SpreadForm.Text
'                End If

                Call ReplaceReport("#Pre-medication#", SpreadForm.Text, ReportTemp)
                
            ElseIf i = 9 And SpreadForm.CellType = CellTypeComboBox Then
                '#cleansingagent#
                Call ReplaceReport("#cleansingagent#", SpreadForm.Text, ReportTemp)
                
            ElseIf i = 10 And SpreadForm.CellType = CellTypeComboBox Then
                '#cleansinglevel#
                Call ReplaceReport("#cleansinglevel#", SpreadForm.Text, ReportTemp)
            ElseIf i = 11 And SpreadForm.CellType = CellTypeComboBox Then
                '#Inserttion#
                Call ReplaceReport("#Inserttion#", SpreadForm.Text, ReportTemp)
'            ElseIf i = 13 And SpreadForm.CellType = CellTypeEdit Then
'                '#finding#
'                Call ReplaceReport("#finding#", SpreadForm.Text, ReportTemp)
'            ElseIf i = 16 And SpreadForm.CellType = CellTypeEdit Then
'                '#Diagnosis#
'                Call ReplaceReport("#Diagnosis#", SpreadForm.Text, ReportTemp)
            ElseIf i = 14 And SpreadForm.CellType = CellTypeComboBox Then
                '#examination#
                Call ReplaceReport("#examination#", SpreadForm.Text, ReportTemp)
            ElseIf i = 15 And SpreadForm.CellType = CellTypeComboBox Then
                '#Complication#
                Call ReplaceReport("#Complication#", SpreadForm.Text, ReportTemp)
            ElseIf i = 16 And SpreadForm.CellType = CellTypeEdit Then
                '#Complication#
                Call ReplaceReport("#Suggestion#", SpreadForm.Text, ReportTemp)
            End If
        Next
    Next
    '/頁籤(2)Finding
    SpreadForm.sheet = 2
    For i = 1 To SpreadForm.row
        SpreadForm.row = i
        For j = 1 To SpreadForm.Col
            SpreadForm.Col = j
            If i = 2 And SpreadForm.CellType = CellTypeEdit Then
                '#finding#
                Call ReplaceReport("#finding#", SpreadForm.Text, ReportTemp)
            End If
        Next
    Next
    
    '/頁籤(3)Diagnosis
    SpreadForm.sheet = 3
    For i = 1 To SpreadForm.row
        SpreadForm.row = i
        For j = 1 To SpreadForm.Col
            SpreadForm.Col = j
            If i = 2 And SpreadForm.CellType = CellTypeEdit Then
                '#Diagnosis#
                Call ReplaceReport("#Diagnosis#", SpreadForm.Text, ReportTemp)
            End If
        Next
    Next
    
    
'    Call UpdateExamOnline(uni_key, Chartno, ReportTemp)
    Spread2Text_ColonOut_1 = Trim(ReportTemp)
 End Function
  
 
''/原本做的~不能用了
'Public Function Spread2Text_bak(ByVal uni_key As String, ByVal Chartno As String, ByRef SpreadForm As fpSpread) As String
'    Dim ReturnSpreadReport As String
'    '/最後10行都是圖片所以就不判斷了
'    Dim RowInfoTemp As String
'    Dim RowInfoFlag As Boolean
'    Dim CheckTemp As String
'    Dim CountIndex As Integer
'    CountIndex = 1
'    For i = 1 To SpreadForm.MaxRows - 10
'        SpreadForm.row = i
'        '/預設都是沒資料
'        RowInfoFlag = False
'        RowInfoTemp = ""
'        For j = 1 To SpreadForm.MaxCols
'            If i = 3 Or i = 4 Or i = 5 Then
'                Exit For '/因為3~4~5都有checkbox 要算所以跳過
'            End If
'
'            SpreadForm.Col = j
'            Select Case SpreadForm.CellType
'                Case CellTypeStaticText: 'Label組成
'                    If SpreadForm.Text <> "" Then
'                        RowInfoTemp = RowInfoTemp & SpreadForm.Text & " "
'                    End If
'
'                Case CellTypeCheckBox: 'checkBox組成
'                    CheckTemp = GetCheckBox(SpreadForm)
'                     If CheckTemp <> "" Then
'                        RowInfoTemp = RowInfoTemp & CheckTemp
'                        RowInfoFlag = True
'                        Exit For
'                     End If
'                Case CellTypeComboBox: '/ComBo組成
'                    If SpreadForm.Text <> "" Then
'                        RowInfoTemp = RowInfoTemp & SpreadForm.Text
'                        RowInfoFlag = True
'                    End If
'
'                Case CellTypeEdit:
'                    If SpreadForm.Text <> "" Then
'                        If i = 13 Then
'                            RowInfoTemp = RowInfoTemp & Replace(SpreadForm.Text, vbCrLf, vbCrLf & "              ")
'                        Else
'                            RowInfoTemp = RowInfoTemp & SpreadForm.Text
'                        End If
'                        RowInfoFlag = True
'                    End If
'            End Select
'        Next
'        If RowInfoFlag = True Then
'            RowInfoTemp = CountIndex & "." & RowInfoTemp
'            ReturnSpreadReport = ReturnSpreadReport & RowInfoTemp & vbCrLf
'        End If
'    Next
'    Call UpdateExamOnline(uni_key, Chartno, ReturnSpreadReport)
'    Spread2Text = ReturnSpreadReport
'
'End Function

'20131007，修改為多一個參數X，表示抓取從第二行到第X行的值，以因應不同報表
Public Function GetCheckBox(ByRef sp As fpSpread, ByVal x As Integer) As String
    Dim CheckTemp As String
    For i = 2 To x
        sp.row = i
        For j = 1 To sp.MaxCols
            sp.Col = j
            Select Case sp.CellType
                Case CellTypeCheckBox:
                    If sp.Value = "1" Then
                        sp.Col = sp.Col + 1
                        CheckTemp = CheckTemp & sp.Text & ","
                    End If
                Case CellTypeEdit:
                    If sp.Text <> "" Then
                        CheckTemp = CheckTemp & sp.Text & ","
                    End If
            End Select
        Next
    Next
    '/判斷Others
    If CheckTemp = "" Then
        CheckTemp = "Nil"
    Else
        CheckTemp = RejectTxt(Mid(CheckTemp, 1, Len(CheckTemp) - 1))
    End If
    
    GetCheckBox = CheckTemp
End Function

'20130912，ERCP_1.rps專用，mRow啟始行，nRow最後行，抓出mRow~nRow間所有已打勾的checkbox的文字，含other文字
Public Function GetCheckBox_1(ByRef sp As fpSpread, mRow As Integer, nRow As Integer) As String
    Dim CheckTemp As String
    For i = mRow To nRow
        sp.row = i
        For j = 1 To sp.MaxCols
            sp.Col = j
            Select Case sp.CellType
                Case CellTypeCheckBox:
                    If sp.Value = "1" Then
                        sp.Col = sp.Col + 1
                        CheckTemp = CheckTemp & sp.Text & ","
                    End If
                Case CellTypeEdit:
                    If sp.Text <> "" Then
                        CheckTemp = CheckTemp & sp.Text & ","
                    End If
            End Select
        Next
    Next
    '/判斷Others
    If CheckTemp = "" Then
        CheckTemp = "Nil"
    Else
        CheckTemp = RejectTxt(Mid(CheckTemp, 1, Len(CheckTemp) - 1))
    End If
    
    GetCheckBox_1 = CheckTemp
End Function

Public Sub UpdateExamOnline(ByVal uni_key As String, ByVal chartno As String, ByVal Item6Report As String)
    Dim SQLString As String

    SQLString = "Update Cris_Exam_Online "
    SQLString = SQLString & "Set Item6 ='" & Replace(Item6Report, "'", "''") & "' "
    SQLString = SQLString & "Where Uni_key ='" & uni_key & "' and Chartno ='" & chartno & "'"
    Connection.Execute (SQLString)
End Sub

Public Sub ReplaceReport(ByVal RKey As String, ByVal RValue As String, ByRef ReportContent As String)
    If RValue = "" Then
        ReportContent = Replace(ReportContent, RKey, "Nil")
    Else
        ReportContent = Replace(ReportContent, RKey, RejectTxt(RValue))
    End If
End Sub


Public Function RejectTxt(ByVal str As String) As String
    Dim StrArr() As String
    Dim RejectCon As String
    If str <> "" Then
        StrArr = Split(str, vbCrLf)
        For i = 0 To UBound(StrArr)
            If i <> 0 Then
                RejectCon = RejectCon & "  " & StrArr(i) & vbCrLf
            Else
                RejectCon = StrArr(i) & vbCrLf
            End If
        Next
    End If
    If Right(RejectCon, 1) = Chr(10) Then
        RejectCon = Mid(RejectCon, 1, Len(RejectCon) - 1)
    End If
    RejectTxt = RejectCon
End Function


'2013年8月26更改前的舊報表。
Public Function Spread2TextEndo(ByVal uni_key As String, ByVal chartno As String, ByRef SpreadForm As fpSpread)
    Dim ReportTemp As String
    ReportTemp = "[Indication]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Indication#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Finding]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Finding#" & vbCrLf
    
    '/跟據randy2013/7/15的信件寫的這段不要
'    ReportTemp = ReportTemp & "[Procedure]" & vbCrLf
'    ReportTemp = ReportTemp & "  " & "#Procedure#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Impression]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Impression#" & vbCrLf
    
    '因尚不知這個內容的來源，所以暫時填空白
    ReportTemp = ReportTemp & "[Comment]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Comment#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Others]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Others#" & vbCrLf
    
    For i = 1 To SpreadForm.MaxRows
        SpreadForm.row = i
        For j = 1 To SpreadForm.MaxCols
        
            If i = 3 Or i = 4 Or i = 5 Or i = 6 Then
                Exit For '/因為3~4~5都有checkbox 要算所以跳過
            End If
            SpreadForm.Col = j
            
            If i = 2 Then
                '#Indication位置
                ReportTemp = Replace(ReportTemp, "#Indication#", GetCheckBox(SpreadForm, 6))
                
            ElseIf i = 8 And SpreadForm.CellType = CellTypeEdit Then
                '#Finding#位置
                Call ReplaceReport("#Finding#", GetEndoFinding(SpreadForm), ReportTemp)

'            ElseIf i = 13 And SpreadForm.CellType = CellTypeCheckBox Then
'                '#Procedure# 位置
'                Call ReplaceReport("#Procedure#", GetEndoProcedure(SpreadForm), ReportTemp)
                
            ElseIf i = 28 And SpreadForm.CellType = CellTypeEdit Then
                '#Impression# 位置
                Call ReplaceReport("#Impression#", SpreadForm.Text, ReportTemp)
                
            ElseIf i = 32 And SpreadForm.CellType = CellTypeEdit Then
                '#Comment# 位置
                Call ReplaceReport("#Comment#", SpreadForm.Text, ReportTemp)
                
            ElseIf i = 36 And SpreadForm.CellType = CellTypeEdit Then
                '#Others# 位置
                Call ReplaceReport("#Others#", SpreadForm.Text, ReportTemp)
            End If
        Next
    Next
    Spread2TextEndo = Trim(ReportTemp)
End Function

'/胃鏡新版報表轉出，2013年8月26更新胃鏡報表。
Public Function Spread2TextEndo_1(ByVal uni_key As String, ByVal chartno As String, ByRef SpreadForm As fpSpread)
    Dim ReportTemp As String
    ReportTemp = "[Indication]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Indication#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Pre-Medication]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Pre-Medication#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Finding]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Finding#" & vbCrLf
    
    '/跟據randy2013/7/15的信件寫的這段不要
'    ReportTemp = ReportTemp & "[Procedure]" & vbCrLf
'    ReportTemp = ReportTemp & "  " & "#Procedure#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Impression]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Impression#" & vbCrLf
    
    '因尚不知這個內容的來源，所以暫時填空白
    ReportTemp = ReportTemp & "[Comment]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Comment#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Others]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Others#" & vbCrLf
    
    For i = 1 To SpreadForm.MaxRows
        SpreadForm.row = i
        For j = 1 To SpreadForm.MaxCols
        
            If i = 3 Or i = 4 Or i = 5 Or i = 6 Then
                Exit For '/因為3~4~5都有checkbox 要算所以跳過
            End If
            SpreadForm.Col = j
            
            If i = 2 Then
                '#Indication位置
                ReportTemp = Replace(ReportTemp, "#Indication#", GetCheckBox(SpreadForm, 6))
            ElseIf i = 8 And SpreadForm.CellType = CellTypeComboBox Then
                '#Pre-Medication#位置
                Dim EndoPre As String
                EndoPre = "1.Gascon drop oral 10-15c.c." & vbCrLf
                EndoPre = EndoPre & "2.Xylocaine(10%) pump spray 1-3puffs"
                If UCase(Trim(SpreadForm.Text)) <> "NIL" And Trim(SpreadForm.Text <> "") Then
                   EndoPre = EndoPre & vbCrLf & "3." & SpreadForm.Text
                End If
                
                Call ReplaceReport("#Pre-Medication#", EndoPre, ReportTemp)
            ElseIf i = 11 And SpreadForm.CellType = CellTypeEdit Then
                '#Finding#位置
                Call ReplaceReport("#Finding#", GetEndoFinding_1(SpreadForm), ReportTemp)

'            ElseIf i = 13 And SpreadForm.CellType = CellTypeCheckBox Then
'                '#Procedure# 位置
'                Call ReplaceReport("#Procedure#", GetEndoProcedure(SpreadForm), ReportTemp)
                
            ElseIf i = 31 And SpreadForm.CellType = CellTypeEdit Then
                '#Impression# 位置
                Call ReplaceReport("#Impression#", SpreadForm.Text, ReportTemp)
                
            ElseIf i = 35 And SpreadForm.CellType = CellTypeEdit Then
                '#Comment# 位置
                Call ReplaceReport("#Comment#", SpreadForm.Text, ReportTemp)
                
            ElseIf i = 39 And SpreadForm.CellType = CellTypeEdit Then
                '#Others# 位置
                Call ReplaceReport("#Others#", SpreadForm.Text, ReportTemp)
            End If
        Next
    Next
    Spread2TextEndo_1 = Trim(ReportTemp)
End Function
'==============================================

'/胃鏡新版報表轉出，2013年8月30更新胃鏡報表。
Public Function Spread2TextEndo_2(ByVal uni_key As String, ByVal chartno As String, ByRef SpreadForm As fpSpread)
    Dim ReportTemp As String
    ReportTemp = "[Indication]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Indication#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Pre-Medication]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Pre-Medication#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Finding]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Finding#" & vbCrLf
    
    '/跟據randy2013/7/15的信件寫的這段不要
'    ReportTemp = ReportTemp & "[Procedure]" & vbCrLf
'    ReportTemp = ReportTemp & "  " & "#Procedure#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Impression]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Impression#" & vbCrLf
    
    '因尚不知這個內容的來源，所以暫時填空白
    ReportTemp = ReportTemp & "[Comment]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Comment#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Others]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Others#" & vbCrLf
    
    '/sheet1
    SpreadForm.sheet = 1
    For i = 1 To SpreadForm.MaxRows
        SpreadForm.row = i
        For j = 1 To SpreadForm.MaxCols
            If i = 3 Or i = 4 Or i = 5 Or i = 6 Then
                Exit For '/因為3~4~5都有checkbox 要算所以跳過
            End If
            SpreadForm.Col = j
            If i = 2 Then
                '#Indication位置
                ReportTemp = Replace(ReportTemp, "#Indication#", GetCheckBox(SpreadForm, 6))
            ElseIf i = 8 And SpreadForm.CellType = CellTypeComboBox Then
                '#Pre-Medication#位置
                Dim EndoPre As String
                EndoPre = "1.Gascon drop oral 10-15c.c." & vbCrLf
                EndoPre = EndoPre & "2.Xylocaine(10%) pump spray 1-3puffs"
                If UCase(Trim(SpreadForm.Text)) <> "NIL" And Trim(SpreadForm.Text <> "") Then
                   EndoPre = EndoPre & vbCrLf & "3." & SpreadForm.Text
                End If
                Call ReplaceReport("#Pre-Medication#", EndoPre, ReportTemp)
            End If
        Next
    Next

    '/sheet2
    SpreadForm.sheet = 2
    For i = 1 To SpreadForm.MaxRows
        SpreadForm.row = i
        For j = 1 To SpreadForm.MaxCols
            SpreadForm.Col = j
            If i = 2 Then
                Call ReplaceReport("#Finding#", GetEndoFinding_2(SpreadForm), ReportTemp)
                Exit For
            End If
        Next
    Next

    '/sheet3
    SpreadForm.sheet = 3
    
    For i = 1 To SpreadForm.MaxRows
        SpreadForm.row = i
        For j = 1 To SpreadForm.MaxCols
            SpreadForm.Col = j
            If i = 1 And j = 5 Then
                '#Impression# 位置
                Call ReplaceReport("#Impression#", SpreadForm.Text, ReportTemp)
            ElseIf i = 5 And j = 5 Then
                '#Comment# 位置
                Call ReplaceReport("#Comment#", SpreadForm.Text, ReportTemp)
            ElseIf i = 9 And j = 5 Then
                '#Others# 位置
                Call ReplaceReport("#Others#", SpreadForm.Text, ReportTemp)
            End If
        Next
    Next
    Spread2TextEndo_2 = Trim(ReportTemp)
End Function
'==============================================

'/胃鏡新版報表轉出，2013年9月13更新胃鏡報表。
Public Function Spread2TextEndo_3(ByVal uni_key As String, ByVal chartno As String, ByRef SpreadForm As fpSpread)
    Dim ReportTemp As String
    ReportTemp = "[Indication]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Indication#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Pre-Medication]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Pre-Medication#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Finding]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Finding#" & vbCrLf
    
    '/跟據randy2013/7/15的信件寫的這段不要
'    ReportTemp = ReportTemp & "[Procedure]" & vbCrLf
'    ReportTemp = ReportTemp & "  " & "#Procedure#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Impression]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Impression#" & vbCrLf
    
    '因尚不知這個內容的來源，所以暫時填空白
    ReportTemp = ReportTemp & "[Comment]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Comment#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Others]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Others#" & vbCrLf
    
    '/sheet1
    SpreadForm.sheet = 1
    For i = 1 To SpreadForm.MaxRows
        SpreadForm.row = i
        For j = 1 To SpreadForm.MaxCols
            If i = 3 Or i = 4 Or i = 5 Or i = 6 Then
                Exit For '/因為3~4~5都有checkbox 要算所以跳過
            End If
            SpreadForm.Col = j
            If i = 2 Then
                '#Indication位置
                ReportTemp = Replace(ReportTemp, "#Indication#", GetCheckBox(SpreadForm, 6))
            ElseIf i = 8 And SpreadForm.CellType = CellTypeComboBox Then
                '#Pre-Medication#位置
                Dim EndoPre As String
                EndoPre = "1.Gascon drop oral 10-15c.c." & vbCrLf
                EndoPre = EndoPre & "2.Xylocaine(10%) pump spray 1-3puffs"
                '/記錄三個pre-medication
                Dim Pr1  As String, Pr2 As String, Pr3 As String
                '/記錄pre_medication 得編號,從3開始
                Dim PrCount As Integer
                PrCount = 3
                Dim l As Integer, m As Integer
                For l = 8 To 10
                    SpreadForm.row = l
                    For m = 1 To SpreadForm.MaxCols
                        SpreadForm.Col = m
                        If SpreadForm.CellType = CellTypeComboBox Then
                            If UCase(Trim(SpreadForm.Text)) <> "NIL" And Trim(SpreadForm.Text <> "") Then
                                EndoPre = EndoPre & vbCrLf & PrCount & "." & SpreadForm.Text
                                PrCount = PrCount + 1
                            End If
                            Exit For
                        End If
                    Next
                Next
                Call ReplaceReport("#Pre-Medication#", EndoPre, ReportTemp)
            End If
        Next
    Next

    '/sheet2
    SpreadForm.sheet = 2
    For i = 1 To SpreadForm.MaxRows
        SpreadForm.row = i
        For j = 1 To SpreadForm.MaxCols
            SpreadForm.Col = j
            If i = 2 Then
                Call ReplaceReport("#Finding#", GetEndoFinding_3(SpreadForm), ReportTemp)
                Exit For
            End If
        Next
    Next

    '/sheet3
    SpreadForm.sheet = 3
    
    For i = 1 To SpreadForm.MaxRows
        SpreadForm.row = i
        For j = 1 To SpreadForm.MaxCols
            SpreadForm.Col = j
            If i = 2 And j = 5 Then
                '#Impression# 位置
                Call ReplaceReport("#Impression#", SpreadForm.Text, ReportTemp)
            ElseIf i = 6 And j = 5 Then
                '#Comment# 位置
                Call ReplaceReport("#Comment#", SpreadForm.Text, ReportTemp)
            ElseIf i = 10 And j = 5 Then
                '#Others# 位置
                Call ReplaceReport("#Others#", SpreadForm.Text, ReportTemp)
            End If
        Next
    Next
    Spread2TextEndo_3 = Trim(ReportTemp)
End Function
'==============================================


''
Public Function GetEndoFinding(ByRef sp As fpSpread) As String
    Dim tmp As String
    
    Dim esophagus As String
    Dim stromach As String
    Dim duodenum As String

    tmp = ""
    '=============Esophagus===========
    sp.row = 8
    sp.Col = 11
    esophagus = sp.Value
    If Trim(esophagus) <> "" Then
        '/如果找到內容是兩行以上
        If InStr(1, esophagus, vbCrLf) <> 0 Then
            tmp = tmp & "Esophagus:" & vbCrLf
            esophagus = "    " & Replace(esophagus, vbCrLf, vbCrLf & "    ")
            tmp = tmp & esophagus & vbCrLf
        Else
            tmp = tmp & "Esophagus:" & esophagus & vbCrLf
        End If
    End If
    
    '=============Stromach============
    sp.row = 12
    sp.Col = 11
    stromach = sp.Value
    '/如果找到內容是兩行以上
    If Trim(stromach) <> "" Then
        If InStr(1, stromach, vbCrLf) <> 0 Then
            tmp = tmp & "Stomach:" & vbCrLf
            stromach = "    " & Replace(stromach, vbCrLf, vbCrLf & "    ")
            tmp = tmp & stromach & vbCrLf
        Else
            tmp = tmp & "Stomach:" & stromach & vbCrLf
        End If
    End If
    '=============Duodenum============
    sp.row = 16
    sp.Col = 11
    duodenum = sp.Value
    '/如果找到內容是兩行以上
    If Trim(duodenum) <> "" Then
        If InStr(1, duodenum, vbCrLf) <> 0 Then
            tmp = tmp & "Duodenum:" & vbCrLf
            duodenum = "    " & Replace(duodenum, vbCrLf, vbCrLf & "    ")
            tmp = tmp & duodenum
        Else
            tmp = tmp & "Duodenum:" & duodenum
        End If
    End If
    GetEndoFinding = tmp
End Function


'/20130826新胃鏡報表。
Public Function GetEndoFinding_1(ByRef sp As fpSpread) As String
    Dim tmp As String
    
    Dim esophagus As String
    Dim stromach As String
    Dim duodenum As String

    tmp = ""
    '=============Esophagus===========
    sp.row = 11
    sp.Col = 11
    esophagus = sp.Value
    If Trim(esophagus) <> "" Then
        '/如果找到內容是兩行以上
        If InStr(1, esophagus, vbCrLf) <> 0 Then
            tmp = tmp & "Esophagus:" & vbCrLf
            esophagus = "    " & Replace(esophagus, vbCrLf, vbCrLf & "    ")
            tmp = tmp & esophagus & vbCrLf
        Else
            tmp = tmp & "Esophagus:" & esophagus & vbCrLf
        End If
    End If
    
    '=============Stromach============
    sp.row = 15
    sp.Col = 11
    stromach = sp.Value
    '/如果找到內容是兩行以上
    If Trim(stromach) <> "" Then
        If InStr(1, stromach, vbCrLf) <> 0 Then
            tmp = tmp & "Stomach:" & vbCrLf
            stromach = "    " & Replace(stromach, vbCrLf, vbCrLf & "    ")
            tmp = tmp & stromach & vbCrLf
        Else
            tmp = tmp & "Stomach:" & stromach & vbCrLf
        End If
    End If
    '=============Duodenum============
    sp.row = 19
    sp.Col = 11
    duodenum = sp.Value
    '/如果找到內容是兩行以上
    If Trim(duodenum) <> "" Then
        If InStr(1, duodenum, vbCrLf) <> 0 Then
            tmp = tmp & "Duodenum:" & vbCrLf
            duodenum = "    " & Replace(duodenum, vbCrLf, vbCrLf & "    ")
            tmp = tmp & duodenum
        Else
            tmp = tmp & "Duodenum:" & duodenum
        End If
    End If
    GetEndoFinding_1 = tmp
End Function

'/20130826新胃鏡報表。
Public Function GetEndoFinding_2(ByRef sp As fpSpread) As String
    Dim tmp As String
    
    Dim esophagus As String
    Dim stromach As String
    Dim duodenum As String
    tmp = ""
    '=============Esophagus===========
    
    sp.row = 2
    sp.Col = 6
    esophagus = sp.Value
    If Trim(esophagus) <> "" Then
        '/如果找到內容是兩行以上
        If InStr(1, esophagus, vbCrLf) <> 0 Then
            tmp = tmp & "Esophagus:" & vbCrLf
            esophagus = "    " & Replace(esophagus, vbCrLf, vbCrLf & "    ")
            tmp = tmp & esophagus & vbCrLf
        Else
            tmp = tmp & "Esophagus:" & esophagus & vbCrLf
        End If
    End If
    
    '=============Stromach============
    sp.row = 6
    sp.Col = 6
    stromach = sp.Value
    '/如果找到內容是兩行以上
    If Trim(stromach) <> "" Then
        If InStr(1, stromach, vbCrLf) <> 0 Then
            tmp = tmp & "Stomach:" & vbCrLf
            stromach = "    " & Replace(stromach, vbCrLf, vbCrLf & "    ")
            tmp = tmp & stromach & vbCrLf
        Else
            tmp = tmp & "Stomach:" & stromach & vbCrLf
        End If
    End If
    '=============Duodenum============
    sp.row = 10
    sp.Col = 6
    duodenum = sp.Value
    '/如果找到內容是兩行以上
    If Trim(duodenum) <> "" Then
        If InStr(1, duodenum, vbCrLf) <> 0 Then
            tmp = tmp & "Duodenum:" & vbCrLf
            duodenum = "    " & Replace(duodenum, vbCrLf, vbCrLf & "    ")
            tmp = tmp & duodenum
        Else
            tmp = tmp & "Duodenum:" & duodenum
        End If
    End If
    GetEndoFinding_2 = tmp
End Function

'/20130913新胃鏡報表。
Public Function GetEndoFinding_3(ByRef sp As fpSpread) As String
    Dim tmp As String
    
    Dim esophagus As String
    Dim stromach As String
    Dim duodenum As String
    tmp = ""
    '=============Esophagus===========
    
    sp.row = 2
    sp.Col = 6
    esophagus = sp.Value
    If Trim(esophagus) <> "" Then
        '/如果找到內容是兩行以上
        If InStr(1, esophagus, vbCrLf) <> 0 Then
            tmp = tmp & "Esophagus:" & vbCrLf
            esophagus = "    " & Replace(esophagus, vbCrLf, vbCrLf & "    ")
            tmp = tmp & esophagus & vbCrLf
        Else
            tmp = tmp & "Esophagus:" & esophagus & vbCrLf
        End If
    End If
    
    '=============Stromach============
    sp.row = 6
    sp.Col = 6
    stromach = sp.Value
    '/如果找到內容是兩行以上
    If Trim(stromach) <> "" Then
        If InStr(1, stromach, vbCrLf) <> 0 Then
            tmp = tmp & "Stomach:" & vbCrLf
            stromach = "    " & Replace(stromach, vbCrLf, vbCrLf & "    ")
            tmp = tmp & stromach & vbCrLf
        Else
            tmp = tmp & "Stomach:" & stromach & vbCrLf
        End If
    End If
    '=============Duodenum============
    sp.row = 10
    sp.Col = 6
    duodenum = sp.Value
    '/如果找到內容是兩行以上
    If Trim(duodenum) <> "" Then
        If InStr(1, duodenum, vbCrLf) <> 0 Then
            tmp = tmp & "Duodenum:" & vbCrLf
            duodenum = "    " & Replace(duodenum, vbCrLf, vbCrLf & "    ")
            tmp = tmp & duodenum
        Else
            tmp = tmp & "Duodenum:" & duodenum
        End If
    End If
    GetEndoFinding_3 = tmp
End Function


'
'Public Function GetEndoFinding(ByRef sp As fpSpread) As String
'    Dim tmp As String
'    Dim finding As String
'    Dim flag As Boolean
'    For i = 8 To 11
'        sp.row = i
'        flag = False
'        tmp = ""
'        For j = 1 To sp.MaxCols
'            sp.Col = j
'            If sp.CellType = CellTypeEdit Or sp.CellType = CellTypeComboBox Then
'                If sp.Text <> "" Then
'                    flag = True
'                    '/只有finding第一格是用打的一開始就要先斷行
'                    If sp.CellType = CellTypeEdit Then
'                        If InStr(1, sp.Text, vbCrLf) Then
'                            tmp = tmp & vbCrLf & "  " & Replace(sp.Text, vbCrLf, vbCrLf & "  ")
'                        Else
'                            tmp = tmp & sp.Text
'                        End If
'                    Else
'                        tmp = tmp & sp.Text
'                    End If
'
'                End If
'            Else
'                If sp.Text <> "" Then
'                    tmp = tmp & sp.Text
'                End If
'            End If
'        Next
'        If flag = True Then
'            finding = finding & tmp & vbCrLf
'        End If
'    Next
'    If Right(finding, 1) = Chr(10) Then
'        finding = Mid(finding, 1, Len(finding) - 1)
'    End If
'    GetEndoFinding = finding
'End Function


Public Function GetEndoProcedure(ByRef sp As fpSpread) As String
    Dim CheckTemp As String
    For i = 13 To 19
        sp.row = i
        For j = 1 To sp.MaxCols
            sp.Col = j
            Select Case sp.CellType
                Case CellTypeCheckBox:
                    If sp.Value = "1" Then
                        sp.Col = sp.Col + 2
                        CheckTemp = CheckTemp & sp.Text & ","
                    End If
                Case CellTypeEdit:
                    If sp.Text <> "" Then
                        CheckTemp = CheckTemp & sp.Text & ","
                    End If
            End Select
        Next
    Next
    GetEndoProcedure = CheckTemp
End Function


Public Function Spread2TextAUTR(ByVal uni_key As String, ByVal chartno As String, ByRef SpreadForm As fpSpread) As String

'=====================Indication 部份: 1~5行===========================
    Dim Ind(9) As String
    '記錄那些Indication存起來了
    Dim IndIdx As Integer
    Dim IndicationContent As String
    
    IndIdx = 0
    
    Dim i As Integer

    For i = 1 To 5
        SpreadForm.row = i
        For j = 1 To SpreadForm.MaxCols
            SpreadForm.Col = j
            Select Case SpreadForm.CellType
                Case CellTypeCheckBox:
                    If SpreadForm.Value = "1" Then
                        '/因為checkbox的內容在下一格要將col 加1
                        SpreadForm.Col = SpreadForm.Col + 1
                        Ind(IndIdx) = SpreadForm.Value
                        IndIdx = IndIdx + 1
                    End If
            End Select
        Next
    Next
    '/判段完所有的Indication 時 將他組起來
    IndicationContent = "INDICATION: "
    For i = 0 To UBound(Ind)
        If Ind(i) <> "" Then
            IndicationContent = IndicationContent & Ind(i) & ";"
        Else
            If i = 0 Then
                IndicationContent = ""
            Else
                IndicationContent = Mid(IndicationContent, 1, Len(IndicationContent) - 1)
            End If
            '/如果沒有資料後面就不做了
            Exit For
        End If
    Next
    If IndicationContent <> "" Then
        IndicationContent = IndicationContent & vbCrLf
    End If
'=====================Indication End ===========================


'=====================Liver==============================
'/Liver 是否為正常的flag
Dim LiverFlag As Boolean
LiverFlag = False
'/7~9行的CheckBox 陣列
Dim ECHOGENICITY() As String, SURFACE() As String, MARGIN() As String, SIZE() As String
Dim LiverContent As String
For i = 7 To 11
    SpreadForm.row = i
    '/用來檢查存在那一個陣列位置的變數
    Dim CheckBoxFlag As Integer
    '/換行時為0
    CheckBoxFlag = 0
    For j = 1 To SpreadForm.MaxCols
        SpreadForm.Col = j
        '/如果Liver是Normal 的話
        If i = 7 And j = 4 And SpreadForm.Value = "1" Then
            LiverFlag = True
            Exit For
        Else
            '/開始分行處理
            Select Case i
                Case 8: '/ECHOGENICITY 一行只處理一次
                    Select Case SpreadForm.CellType
                        Case CellTypeCheckBox:
                            Call GetSelectCheck(i, SpreadForm, ECHOGENICITY)
                            Exit For
                    End Select
                Case 9: '/SURFACE
                    Select Case SpreadForm.CellType
                        Case CellTypeCheckBox:
                            Call GetSelectCheck(i, SpreadForm, SURFACE)
                            Exit For
                    End Select
                Case 10: '/MARGIN
                    Select Case SpreadForm.CellType
                        Case CellTypeCheckBox:
                            Call GetSelectCheck(i, SpreadForm, MARGIN)
                            Exit For
                    End Select
                Case 11: '/size
                    Select Case SpreadForm.CellType
                        Case CellTypeCheckBox:
                            Call GetSelectCheck(i, SpreadForm, SIZE)
                            Exit For
                    End Select
            End Select
        End If
    Next
    If LiverFlag = True Then Exit For
Next

Dim HYPOEHOICTemp As String
HYPOEHOICTemp = ""

For i = 12 To 46 Step 6
    HYPOEHOICTemp = HYPOEHOICTemp & GetLiverCount(i + 1, SpreadForm)
    If i <> 42 Then
        HYPOEHOICTemp = HYPOEHOICTemp & GetCM(i + 2, SpreadForm)
    End If
Next

'/這邊取得CYST的公分
HYPOEHOICTemp = HYPOEHOICTemp & GetCYSTCM(43, SpreadForm)

'/這邊取的 DOPPLLER
HYPOEHOICTemp = HYPOEHOICTemp & GetDOPPLLER(47, SpreadForm)


Dim LiverTemp As String
'/起始值
LiverContent = ""
LiverTemp = ""

If LiverFlag = True Then
    '/正常的情況下
    LiverContent = "[Liver]" & vbCrLf & "NORMAL" & vbCrLf & vbCrLf
Else
    '/=======ECHOGENICITY ========
    LiverTemp = Merge(ECHOGENICITY, "ECHOGENICITY")
    '/=======SURFACE ========
    LiverTemp = LiverTemp & Merge(SURFACE, "SURFACE")
    '/=========MARGIN =======
    LiverTemp = LiverTemp & Merge(MARGIN, "MARGIN")
    '/=========SIZE =======
    LiverTemp = LiverTemp & Merge(SIZE, "SIZE")
    
    
    If LiverTemp & HYPOEHOICTemp <> "" Then
        LiverContent = "[LIVER]" & vbCrLf & LiverTemp & HYPOEHOICTemp & GetOther(52, SpreadForm) & vbCrLf
    End If
End If

'=====================Liver End==============================

'=========================GALLBLADDER========================
Dim GALLBLADDERContent As String
Dim GALLBLADDERCou As String
Dim GALLBLADDERCM As String
Dim GALLBLADDER() As String
SpreadForm.row = 53
SpreadForm.Col = 8
If SpreadForm.Value = "1" Then
    GALLBLADDERContent = "[GALLBLADDER]" & vbCrLf & "NORMAL " & vbCrLf
Else
    Dim GalTitle As String
    '/將有選的值填進陣列
    Call GetSelectCheck(54, SpreadForm, GALLBLADDER)
    '/組出前面的checkbox
    For i = 0 To UBound(GALLBLADDER)
        If GALLBLADDER(i) <> "" Then
            GALLBLADDERContent = GALLBLADDERContent & GALLBLADDER(i) & ","
        End If
    Next
    '/
    SpreadForm.row = 55
    SpreadForm.Col = 8
    If SpreadForm.Value = "1" Then
        SpreadForm.Col = SpreadForm.Col + 1
        GALLBLADDERContent = GALLBLADDERContent & SpreadForm.Value & ","
    End If
    
    
    If GALLBLADDERContent <> "" Then
        GALLBLADDERContent = Mid(GALLBLADDERContent, 1, Len(GALLBLADDERContent) - 1) & ";"
    End If
    
    For i = 56 To 69 Step 7
        GalTitle = GetGALLBLADDERTitle(i, SpreadForm)
        '/取得顆數
        GALLBLADDERCou = GetGALLBLADDERCount(i + 2, SpreadForm)
        '/取的公分或Other
        GALLBLADDERCM = GetGALLBLADDERCM(i + 3, SpreadForm)
        
        GALLBLADDERContent = GALLBLADDERContent & GALLBLADDERCou & GalTitle & GALLBLADDERCM
    Next
    If GetOther(70, SpreadForm) <> "" Then
        GALLBLADDERContent = GALLBLADDERContent & vbCrLf & GetOther(70, SpreadForm)
    End If
    
    If GALLBLADDERContent <> "" Then
        GALLBLADDERContent = "[GALLBLADDER] " & vbCrLf & GALLBLADDERContent
    End If
End If
    If GALLBLADDERContent <> "" Then
        GALLBLADDERContent = GALLBLADDERContent & vbCrLf
        Call Check2VbCrlf(GALLBLADDERContent)
    End If
    
    
'=========================INTRAHEPATIC=======================
Dim INTYRAHEPATICContent As String
INTYRAHEPATICContent = GetINTRAHEPATIC(SpreadForm)
If INTYRAHEPATICContent <> "" Then
    INTYRAHEPATICContent = INTYRAHEPATICContent & vbCrLf
    Call Check2VbCrlf(INTYRAHEPATICContent)
End If
'=========================COMMON BILE====================

Dim COMMONBILEContent As String
COMMONBILEContent = GetCOMMONBILE(SpreadForm)
If COMMONBILEContent <> "" Then
    COMMONBILEContent = COMMONBILEContent & vbCrLf
    Call Check2VbCrlf(COMMONBILEContent)
End If
'=========================PORTAL VEIN====================
Dim PORTALContent As String
PORTALContent = GetPORTAL(SpreadForm)

'/這邊接完DOPLLER 再換行
If PORTALContent <> "" Then
    PORTALContent = PORTALContent & vbCrLf
    Call Check2VbCrlf(PORTALContent)
End If

'=========================SPLEEN====================
Dim SPLEENContent As String
SPLEENContent = GetSPLEEN(SpreadForm)
If SPLEENContent <> "" Then
    SPLEENContent = SPLEENContent & vbCrLf
    Call Check2VbCrlf(SPLEENContent)
End If

'=======================PANCREAS====================
Dim PANCREASContent As String
PANCREASContent = GetPANCREAS(SpreadForm)
If PANCREASContent <> "" Then
    PANCREASContent = PANCREASContent & vbCrLf
    Call Check2VbCrlf(PANCREASContent)
End If
'=======================KIDNEY=====================
Dim KIDNEYContent As String
KIDNEYContent = GetKIDNEY(SpreadForm)
If KIDNEYContent <> "" Then
    KIDNEYContent = KIDNEYContent & vbCrLf
    Call Check2VbCrlf(KIDNEYContent)
End If
'=======================ASCITES=====================
Dim ASCITESContent  As String
ASCITESContent = GetASCITES(SpreadForm)

If ASCITESContent <> "" Then
    ASCITESContent = ASCITESContent & vbCrLf
    Call Check2VbCrlf(ASCITESContent)
End If
'=======================OTHER=====================
'Dim OtherContent As String
'OtherContent = GetOther(SpreadForm)

'=======================DIAGNOSIS=====================
Dim DIAGNOSISContent As String
DIAGNOSISContent = GetDIAGNOSIS(SpreadForm)
If DIAGNOSISContent <> "" Then
    DIAGNOSISContent = "[DIAGNOSIS]" & vbCrLf & DIAGNOSISContent & vbCrLf
    Call Check2VbCrlf(DIAGNOSISContent)
End If


'/===================SUGGEST======================
Dim SUGGESTContent As String
SUGGESTContent = GetSUGGEST(SpreadForm)

'Indication不組
'Spread2TextAUTR = Spread2TextAUTR & IndicationContent

Spread2TextAUTR = Spread2TextAUTR & LiverContent
Spread2TextAUTR = Spread2TextAUTR & GALLBLADDERContent
Spread2TextAUTR = Spread2TextAUTR & INTYRAHEPATICContent
Spread2TextAUTR = Spread2TextAUTR & COMMONBILEContent
Spread2TextAUTR = Spread2TextAUTR & PORTALContent
Spread2TextAUTR = Spread2TextAUTR & SPLEENContent
Spread2TextAUTR = Spread2TextAUTR & PANCREASContent
Spread2TextAUTR = Spread2TextAUTR & KIDNEYContent
Spread2TextAUTR = Spread2TextAUTR & ASCITESContent
Spread2TextAUTR = Spread2TextAUTR & OtherContent
Spread2TextAUTR = Spread2TextAUTR & DIAGNOSISContent
Spread2TextAUTR = Spread2TextAUTR & SUGGESTContent


Spread2TextAUTR = Replace(Spread2TextAUTR, ",", ", ")
Spread2TextAUTR = Replace(Spread2TextAUTR, ";", "; ")

End Function

'/此涵式是用來做句子的組成 => 格式 [ Arr(1),Arr(2),Arr(..) Title ;  ]
Public Function Merge(ByRef Arr() As String, ByVal Title As String) As String
    Dim temp As String
    For i = 0 To UBound(Arr)
        If Arr(i) <> "" Then
            temp = temp & Arr(i) & ","
        Else
            If i = 0 Then
                Exit For
            End If
        End If
    Next
    If temp <> "" Then
        temp = Mid(temp, 1, Len(temp) - 1) & " " & Title & ";"
    Else
        temp = ""
    End If
    Merge = temp
End Function


'/這個涵式是用來將某一行的checkbox塞入陣列
Public Sub GetSelectCheck(ByVal row As Integer, ByRef sp As fpSpread, ByRef ReturnArr() As String)
    sp.row = row
    Dim CheckBoxPoint As Integer
    CheckBoxPoint = 0
    For j = 1 To sp.MaxCols
        sp.Col = j
        Select Case sp.CellType
            Case CellTypeCheckBox:
                If sp.Value = "1" Then
                    sp.Col = sp.Col + 1
                    ReDim Preserve ReturnArr(CheckBoxPoint)
                    ReturnArr(CheckBoxPoint) = sp.Value
                    CheckBoxPoint = CheckBoxPoint + 1
                End If
        End Select
    Next
    '/結束後在宣告一次
    If CheckBoxPoint <> 0 Then
        ReDim Preserve ReturnArr(CheckBoxPoint - 1)
    Else
        ReDim Preserve ReturnArr(CheckBoxPoint)
    End If
End Sub

Public Function GetLiverCount(ByVal i As Integer, ByRef sp As fpSpread) As String
    '/先取得數量然後組出字串
    Dim TitleName As String
    sp.row = i
    Select Case i
        Case 13:
            TitleName = "HYPOEHOIC"
        Case 19:
            TitleName = "ISOECHOIC"
        Case 25:
            TitleName = "HYPERECHOIC"
        Case 31:
            TitleName = "ANECHOIC"
        Case 37:
            TitleName = "HETEROECHOIC"
        Case 43:
            TitleName = "CYST"
            sp.row = 42
    End Select
    For j = 1 To sp.MaxCols
        sp.Col = j
        Select Case sp.CellType
            Case CellTypeCheckBox
                If sp.Value = "1" Then
                    sp.Col = sp.Col + 1
                    Select Case sp.Value
                        Case "1"
                            If InStr(1, TitleName, "CYST") > 0 Then
                                GetLiverCount = "ONE " & TitleName & " "
                            Else
                                GetLiverCount = "ONE " & TitleName & " NODULE "
                            End If
                        Case "2"
                            If InStr(1, TitleName, "CYST") > 0 Then
                                GetLiverCount = "TWO " & TitleName & "S "
                            Else
                                GetLiverCount = "TWO " & TitleName & " NODULES "
                            End If
                        Case "3"
                            If InStr(1, TitleName, "CYST") > 0 Then
                                GetLiverCount = "THREE " & TitleName & "S "
                            Else
                                GetLiverCount = "THREE " & TitleName & " NODULES "
                            End If
                        Case "MORE"
                            If InStr(1, TitleName, "CYST") > 0 Then
                                GetLiverCount = "SEVERAL " & TitleName & "S "
                            Else
                                GetLiverCount = "SEVERAL " & TitleName & " NODULES "
                            End If
                    End Select
                End If
        End Select
    Next
End Function

Public Function GetCM(ByVal k As String, ByRef sp As fpSpread) As String
    '/取得幾個CM
    Dim Location As String
    Dim CMTemp As String
    Dim Cmflag As Integer
    
    For i = k To k + 3
        Cmflag = 0
        sp.row = i
        Select Case i
            Case 17, 23, 29, 35, 41:
                '/如果有找到有打Other 時直接離開
                sp.Col = 10
                If Trim(sp.Value) <> "" Then
                    CMTemp = sp.Value & ";"
                    Exit For
                End If
            Case Else
                
                For j = 1 To sp.MaxCols
                    sp.Col = j
                    
                    Select Case j
                        Case 10:
                            If Trim(sp.Value) <> "" Then
                                CMTemp = CMTemp & sp.Value & " "
                                Cmflag = Cmflag + 1
                            End If
                        Case 16:
                            If Trim(sp.Value) <> "" Then
                                If Cmflag = 1 Then
                                    CMTemp = CMTemp & "X " & sp.Value & " CM "
                                Else
                                    CMTemp = CMTemp & "" & sp.Value & " CM "
                                End If
                                Cmflag = Cmflag + 2
                            End If
                        Case 27:
                            If Cmflag < 2 And Cmflag <> 0 Then
                                CMTemp = CMTemp & "CM "
                            End If
                            If Trim(sp.Value) <> "" Then
       
                                CMTemp = CMTemp & sp.Value & ","
                            Else

                                If Cmflag <> 0 Then
                                    CMTemp = CMTemp & ","
                                End If
                            End If
                    End Select
                    
                Next
        End Select
    Next
    If CMTemp <> "" Then
        CMTemp = Mid(CMTemp, 1, Len(CMTemp) - 1) & ";" & vbCrLf
    End If

    GetCM = CMTemp
End Function

Public Function GetCYSTCM(ByVal k As Integer, ByRef sp As fpSpread) As String
    Dim Location As String
    Dim CYSTCm As String
    For i = k To k + 3
        sp.row = i
        For j = 1 To sp.MaxCols
            sp.Col = j
            Select Case j
                Case 10:
                    If Trim(sp.Value) <> "" Then
                        CYSTCm = CYSTCm & sp.Value & " CM "
                    End If
                Case 21
                    If Trim(sp.Value) <> "" Then
                        CYSTCm = CYSTCm & sp.Value & ","
                    End If
            End Select
        Next
    Next
    
    If CYSTCm <> "" Then
        CYSTCm = Mid(CYSTCm, 1, Len(CYSTCm) - 1) & Location & ";"
    End If
    GetCYSTCM = CYSTCm & Location
End Function

Public Function GetDOPPLLER(ByVal k As Integer, ByRef sp As fpSpread) As String
    Dim DOPPLLERflag As Boolean
    Dim DOPPLLERTemp As String
    
    DOPPLLERflag = False
    For i = k To k + 3
        sp.row = i
        For j = 1 To sp.MaxCols
            sp.Col = j
            If i = k And j = 3 And Trim(sp.Value) <> "" Then
                DOPPLLERflag = True
            End If
            '/如果有下拉才要找有勾的帶出
            If DOPPLLERflag = True Then
                Select Case sp.CellType
                    Case CellTypeCheckBox:
                        If sp.Value = "1" Then
                            Select Case i
                                Case k + 1, k + 3:
                                    sp.Col = sp.Col + 1
                                    DOPPLLERTemp = DOPPLLERTemp & sp.Value
                                    sp.row = sp.row + 1
                                    DOPPLLERTemp = DOPPLLERTemp & sp.Value
                            End Select
                        End If
                End Select
            End If
        Next
    Next
    If DOPPLLERTemp <> "" Then
        GetDOPPLLER = "DOPLLER:" & DOPPLLERTemp & vbCrLf
    Else
        GetDOPPLLER = "" & vbCrLf
    End If
    
End Function


Public Function GetGALLBLADDERTitle(ByVal k As Integer, ByRef sp As fpSpread) As String
    For i = k To k + 1
         sp.row = i
         For j = 1 To sp.MaxCols
            sp.Col = j
            Select Case sp.CellType
                Case CellTypeCheckBox:
                    If sp.Value = "1" Then
                        sp.Col = sp.Col + 1
                        GetGALLBLADDERTitle = GetGALLBLADDERTitle & sp.Value & ","
                        Exit For
                    End If
            End Select
         Next
    Next
    If Trim(GetGALLBLADDERTitle <> "") Then
        GetGALLBLADDERTitle = Mid(GetGALLBLADDERTitle, 1, Len(GetGALLBLADDERTitle) - 1) & " "
    Else
        GetGALLBLADDERTitle = ""
    End If
End Function

Public Function GetGALLBLADDERCount(ByVal k As Integer, ByRef sp As fpSpread) As String
    sp.row = k
    For i = 1 To sp.MaxCols
        sp.Col = i
        Select Case sp.CellType
            Case CellTypeCheckBox:
                If Trim(sp.Value) = "1" Then
                    sp.Col = sp.Col + 1
                    Select Case sp.Value
                        Case "1":
                            GetGALLBLADDERCount = "ONE "
                        Case "2":
                            GetGALLBLADDERCount = "TWO "
                        Case "3":
                            GetGALLBLADDERCount = "THREE "
                        Case "MORE":
                            GetGALLBLADDERCount = "SEVERAL "
                    End Select
                End If
        End Select
    Next
End Function

Public Function GetGALLBLADDERCM(ByVal k As Integer, ByRef sp As fpSpread) As String
    For i = k To k + 3
        sp.row = i
        For j = 1 To sp.MaxCols
            sp.Col = j
            Select Case j
                Case 14
                    If i = k + 3 Then '/other
                        If Trim(sp.Value) <> "" Then
                            GetGALLBLADDERCM = sp.Value
                        End If
                    Else '/不是other
                        If Trim(sp.Value) <> "" Then
                            GetGALLBLADDERCM = GetGALLBLADDERCM & sp.Value & " CM,"
                        End If
                    End If
            End Select
        Next
    Next
    If Trim(GetGALLBLADDERCM) <> "" Then
        GetGALLBLADDERCM = Mid(GetGALLBLADDERCM, 1, Len(GetGALLBLADDERCM) - 1) & ";" & vbCrLf
    Else
        GetGALLBLADDERCM = ""
    End If
    
End Function


Public Function GetINTRAHEPATIC(ByRef sp As fpSpread) As String
    '/先判斷是不是NORMAL
    sp.row = 71
    sp.Col = 10
    If sp.Value = "1" Then
        GetINTRAHEPATIC = "[INTRAHEPATIC DUCT]" & vbCrLf & "NORMAL "
    Else
        sp.row = 72
        For j = 1 To sp.MaxCols
            sp.Col = j
            Select Case sp.CellType
                Case CellTypeCheckBox:
                    If sp.Value = "1" Then
                        sp.Col = sp.Col + 1
                        GetINTRAHEPATIC = GetINTRAHEPATIC & sp.Value & " "
                    End If
            End Select
        Next
        GetINTRAHEPATIC = Replace(GetINTRAHEPATIC, "LOBE(S)", "")
        If InStr(1, GetINTRAHEPATIC, "RT") > 0 Or InStr(1, GetINTRAHEPATIC, "LT") > 0 Then
            GetINTRAHEPATIC = GetINTRAHEPATIC & " LOBE "
        Else
            GetINTRAHEPATIC = GetINTRAHEPATIC & " LOBES "
        End If
        sp.row = 73
        sp.Col = 10
        If sp.Value = "1" Then
            sp.Col = sp.Col + 1
            GetINTRAHEPATIC = GetINTRAHEPATIC & sp.Value & " "
        End If
        
        If GetINTRAHEPATIC <> "" Then
            GetINTRAHEPATIC = "[INTRAHEPATIC DUCT]" & vbCrLf & GetINTRAHEPATIC
        End If
    End If
    If Trim(GetINTRAHEPATIC) <> "" Then
        GetINTRAHEPATIC = Mid(GetINTRAHEPATIC, 1, Len(GetINTRAHEPATIC) - 1) & vbCrLf
    End If
End Function

Public Function GetCOMMONBILE(ByRef sp As fpSpread) As String
    sp.row = 74
    sp.Col = 10
    If sp.Value = "1" Then
        GetCOMMONBILE = "[COMMON BILE DUCT]" & vbCrLf & "NORMAL " & vbCrLf
    Else
        For j = 1 To sp.MaxCols
            sp.Col = j
            Select Case sp.CellType
                Case CellTypeCheckBox:
                    If sp.Value = "1" Then
                        sp.Col = sp.Col + 1
                        GetCOMMONBILE = GetCOMMONBILE & sp.Value & " "
                    End If
                Case CellTypeEdit:
                    If Trim(sp.Value) <> "" Then
                        GetCOMMONBILE = GetCOMMONBILE & sp.Value & " "
                    End If
                    
            End Select
        Next
        
        sp.row = 75
        sp.Col = 10
        If sp.Value = "1" Then
            sp.Col = sp.Col + 1
            If GetCOMMONBILE <> "" Then
                GetCOMMONBILE = GetCOMMONBILE & "," & sp.Value & " "
            Else
                GetCOMMONBILE = GetCOMMONBILE & sp.Value & " "
            End If
        End If
        
        If Trim(GetCOMMONBILE) <> "" Then
            GetCOMMONBILE = "[COMMON BILE DUCT]" & vbCrLf & Mid(GetCOMMONBILE, 1, Len(GetCOMMONBILE) - 1) & vbCrLf
        End If
        
    End If
End Function

Public Function GetPORTAL(ByRef sp As fpSpread) As String
    sp.row = 76
    sp.Col = 10
    If sp.Value = "1" Then
        GetPORTAL = "[PORTAL VEIN]" & vbCrLf & "NORMAL " & vbCrLf
    Else
        sp.row = 77
        For j = 1 To sp.MaxCols
            sp.Col = j
            Select Case sp.CellType
                Case CellTypeCheckBox:
                    If sp.Value = "1" Then
                        sp.Col = sp.Col + 1
                        GetPORTAL = GetPORTAL & sp.Value & " "
                    End If
                Case CellTypeEdit:
                    If Trim(sp.Value) <> "" Then
                        GetPORTAL = GetPORTAL & sp.Value & " "
                    End If
                    
            End Select
        Next
        
        GetPORTAL = GetPORTAL & GetDOPPLLER(78, sp)
        If Trim(GetPORTAL) <> "" Then
            '/後面還有DOPLLER 所以先接分號
            GetPORTAL = "[PORTAL VEIN]" & vbCrLf & Mid(GetPORTAL, 1, Len(GetPORTAL) - 1) & ";" & vbCrLf
        End If
    End If

End Function

Public Function GetSPLEEN(ByRef sp As fpSpread) As String
    sp.row = 84
    sp.Col = 5
    If sp.Value = "1" Then
        GetSPLEEN = "[SPLEEN]" & vbCrLf & "NORMAL" & vbCrLf
    Else
        sp.row = 85
        For j = 1 To sp.MaxCols
            sp.Col = j
            Select Case sp.CellType
                Case CellTypeCheckBox:
                    If sp.Value = "1" Then
                        sp.Col = sp.Col + 1
                        GetSPLEEN = GetSPLEEN & sp.Value & " "
                    End If
                Case CellTypeEdit:
                    If Trim(sp.Value) <> "" Then
                        '/公分
                        GetSPLEEN = GetSPLEEN & sp.Value & " "
                        '/公分後的字
                        sp.Col = sp.Col + 3
                        GetSPLEEN = GetSPLEEN & sp.Value & " "
                    End If
            End Select
        Next
        If GetOther(86, sp) <> "" Then
            GetSPLEEN = GetSPLEEN & vbCrLf & GetOther(86, sp)
        End If
        
        If Trim(GetSPLEEN) <> "" Then
            GetSPLEEN = "[SPLEEN]" & vbCrLf & Mid(GetSPLEEN, 1, Len(GetSPLEEN) - 1) & vbCrLf
        End If
    End If
End Function


Public Function GetPANCREAS(ByRef sp As fpSpread) As String
    sp.row = 87
    sp.Col = 5
    If sp.Value = "1" Then
        GetPANCREAS = "[PANCREAS]" & vbCrLf & "NORMAL " & vbCrLf
    Else
        sp.row = 88
        For j = 1 To sp.MaxCols
            sp.Col = j
            Select Case sp.CellType
                Case CellTypeCheckBox:
                    If sp.Value = "1" Then
                        sp.Col = sp.Col + 1
                        GetPANCREAS = GetPANCREAS & sp.Value & " "
                    End If
            End Select
        Next
        If GetOther(89, sp) <> "" Then
            GetPANCREAS = GetPANCREAS & vbCrLf & GetOther(89, sp)
        End If
        If Trim(GetPANCREAS) <> "" Then
            GetPANCREAS = "[PANCREAS]" & vbCrLf & Mid(GetPANCREAS, 1, Len(GetPANCREAS) - 1) & vbCrLf
        End If
    End If

End Function

Public Function GetKIDNEY(ByRef sp As fpSpread) As String
    Dim V1 As String, V2 As String
    
    sp.row = 90
    sp.Col = 13
    If sp.Value = "1" Then
        sp.Col = sp.Col + 1
        GetKIDNEY = GetKIDNEY & sp.Value & " "
    End If
    
    sp.Col = 20
    If sp.Value = "1" Then
        sp.Col = sp.Col + 1
        GetKIDNEY = GetKIDNEY & sp.Value & " "
    End If
    
    If GetKIDNEY <> "" Then
        GetKIDNEY = GetKIDNEY & vbCrLf
    End If
    
    '/~2
    Dim KIDNEYTitle As String
    Dim KIDNEYSizeTemp As String
    For i = 91 To 93
        V1 = ""
        V2 = ""
        sp.row = i
        KIDNEYTitle = ""
        KIDNEYSizeTemp = ""
        For j = 1 To sp.MaxCols
            sp.Col = j
            Select Case j
                Case 11
                    If sp.Value <> "" Then
                        V1 = sp.Value
                        KIDNEYSizeTemp = KIDNEYSizeTemp & V1 & " CM"
                        
                    End If
                Case 20
                    If sp.Value <> "" Then
                        V2 = sp.Value
                        sp.Col = 16
                        If V1 <> "" Then
                            KIDNEYSizeTemp = KIDNEYSizeTemp & ",CORTEX " & V2 & " CM"
                        Else
                            KIDNEYSizeTemp = KIDNEYSizeTemp & "CORTEX " & V2 & " CM"
                        End If
                    End If
            End Select
        Next
        '///
        sp.Col = 5
        If InStr(1, sp.Value, "LEFT") > 0 Or InStr(1, sp.Value, "RIGHT") > 0 Then
            If InStr(1, sp.Value, "LEFT") > 0 Then
                KIDNEYTitle = KIDNEYTitle & Replace(sp.Value, "LEFT", "LEFT SIDE") & " "
            ElseIf InStr(1, sp.Value, "RIGHT") > 0 Then
                KIDNEYTitle = KIDNEYTitle & Replace(sp.Value, "RIGHT", "RIGHT SIDE") & " "
            End If
        Else
            KIDNEYTitle = KIDNEYTitle & sp.Value
        End If
        
        If V1 <> "" Or V2 <> "" Then
            GetKIDNEY = GetKIDNEY & KIDNEYTitle & " " & KIDNEYSizeTemp & ";" & vbCrLf
        End If
    Next
    '/ANECHOIC AREA
    sp.row = 94
    sp.Col = 5
    Dim ANECHOICFlag As Boolean
    ANECHOICFlag = False
    If sp.Value = "1" Then
        sp.Col = sp.Col + 1
        GetKIDNEY = GetKIDNEY & sp.Value & " "
        ANECHOICFlag = True
        
    End If
    sp.Col = 20
    If sp.Value = "1" Then
        sp.Col = sp.Col + 1
        GetKIDNEY = GetKIDNEY & "NUMBER: " & sp.Value & " "
    End If
    sp.Col = 23
    If sp.Value = "1" Then
        sp.Col = sp.Col + 1
        GetKIDNEY = GetKIDNEY & "NUMBER: " & sp.Value & " "
    End If
    sp.Col = 26
    If sp.Value = "1" Then
        sp.Col = sp.Col + 1
        GetKIDNEY = GetKIDNEY & "NUMBER: " & sp.Value & " "
    End If
    sp.Col = 28
    If sp.Value = "1" Then
        sp.Col = sp.Col + 1
        GetKIDNEY = GetKIDNEY & "NUMBER: " & sp.Value & " "
    End If
    
    Dim temp As String
    For i = 95 To 98
        sp.row = i
        For j = 1 To sp.MaxCols
            sp.Col = j
            Select Case sp.CellType
                Case CellTypeEdit
                    If sp.Value <> "" Then
                        If i <> 98 Then
                            temp = temp & sp.Value & " CM,"
                        Else
                            temp = temp & sp.Value & " "
                        End If
                    End If
            End Select
        Next
    Next
    
    If temp <> "" Then
        GetKIDNEY = GetKIDNEY & "DIAMETER " & Mid(temp, 1, Len(temp) - 1) & ";"
    ElseIf ANECHOICFlag = True Then
        GetKIDNEY = GetKIDNEY & ";"
    End If
    
    
    
    '/HYPOECHOIC
    Dim HYPOECHOICFlag As Boolean
    HYPOECHOICFlag = False
    sp.row = 99
    sp.Col = 5
    If sp.Value = "1" Then
        sp.Col = sp.Col + 1
        GetKIDNEY = GetKIDNEY & sp.Value & " "
        HYPOECHOICFlag = True
    End If
    sp.Col = 20
    If sp.Value = "1" Then
        sp.Col = sp.Col + 1
        GetKIDNEY = GetKIDNEY & "NUMBER: " & sp.Value & " "
    End If
    sp.Col = 23
    If sp.Value = "1" Then
        sp.Col = sp.Col + 1
        GetKIDNEY = GetKIDNEY & "NUMBER: " & sp.Value & " "
    End If
    sp.Col = 26
    If sp.Value = "1" Then
        sp.Col = sp.Col + 1
        GetKIDNEY = GetKIDNEY & "NUMBER: " & sp.Value & " "
    End If
    sp.Col = 28
    If sp.Value = "1" Then
        sp.Col = sp.Col + 1
        GetKIDNEY = GetKIDNEY & "NUMBER: " & sp.Value & " "
    End If
    temp = ""
    For i = 100 To 103
        sp.row = i
        For j = 1 To sp.MaxCols
            sp.Col = j
            Select Case sp.CellType
                Case CellTypeEdit
                    If sp.Value <> "" Then
                        If i <> 103 Then
                            temp = temp & sp.Value & " CM,"
                        Else
                            temp = temp & sp.Value & " "
                        End If
                    End If
            End Select
        Next
    Next
    If temp <> "" Then
        GetKIDNEY = GetKIDNEY & "DIAMETER " & Mid(temp, 1, Len(temp) - 1) & ";"
    ElseIf HYPOECHOICFlag = True Then
        GetKIDNEY = GetKIDNEY & ";"
    End If
    
'===========================

    temp = ""
    sp.row = 104
    sp.Col = 5
    If sp.Value = "1" Then
        sp.Col = sp.Col + 1
        GetKIDNEY = GetKIDNEY & sp.Value & ";"
    End If
    temp = RowCheckBoxForOne(105, sp, " ")
    If Trim(temp) <> "" Then
        GetKIDNEY = GetKIDNEY & "LOCATION :" & temp & ";"
    End If
    If GetOther(106, sp) <> "" Then
        GetKIDNEY = GetKIDNEY & vbCrLf & GetOther(106, sp)
    End If
    If GetKIDNEY <> "" Then
        GetKIDNEY = "[KIDNEY ECHOGENICITY]" & vbCrLf & Mid(GetKIDNEY, 1, Len(GetKIDNEY) - 1) & vbCrLf
    End If
End Function

Public Function GetASCITES(ByRef sp As fpSpread) As String
    sp.row = 107
    For j = 1 To sp.MaxCols
        sp.Col = j
        Select Case sp.CellType
            Case CellTypeCheckBox:
            If sp.Value = "1" Then
                sp.Col = sp.Col + 1
                Select Case UCase(Trim(sp.Value))
                    Case "YES":
                        GetASCITES = "ASCITES WAS NOTED"
                                            
                    Case "NO":
                        GetASCITES = "NO ASCITES WAS NOTED"
                End Select
            End If
        End Select
    Next
    
    If GetASCITES <> "" Then
        GetASCITES = "[ASCITES]" & vbCrLf & GetASCITES & vbCrLf
    End If
End Function


Public Function GetOther(ByVal i As Integer, ByRef sp As fpSpread) As String
    sp.row = i
    sp.Col = 6
    If Trim(sp.Value) <> "" Then
        GetOther = sp.Value & vbCrLf
    Else
        GetOther = ""
    End If
End Function



Public Function GetDIAGNOSIS(ByRef sp As fpSpread) As String
    Dim Liver() As String, LiverFlag As Boolean
    Dim GALLBLADDER() As String, GALLBLADDERFlag As Boolean
    Dim PANCREAS() As String, PANCREASFlag As Boolean
    Dim SPLEEN() As String, SPLEENFlag As Boolean
    Dim KIDNEY() As String, KIDNEYFlag As Boolean
    Dim temp As String
    LiverFlag = False
    Dim CheckPoint As Integer
    CheckPoint = 0
    temp = ""
    LiverFlag = False
    Dim DiagnosisCount As Integer
    DiagnosisCount = 1
    '=======Liver=========
    For i = 110 To 117
        sp.row = i
        For j = 1 To sp.MaxCols
            sp.Col = j
            Select Case sp.CellType
                Case CellTypeCheckBox:
                    If sp.Value = "1" Then
                        sp.Col = sp.Col + 1
                        
                        If i = 111 Then
                            ReDim Preserve Liver(CheckPoint)
                            '將這行全部組出來
                            Liver(CheckPoint) = RowCheckBoxForOne(i, sp, " ")
                            '/如果最後面的字不是 ")"
                            If Right(Trim(Liver(CheckPoint)), 1) <> ")" Then Liver(CheckPoint) = Liver(CheckPoint) & ")"
                            CheckPoint = CheckPoint + 1
                            '/因為這行已經特殊處理完畢了不需要再處理。
                            Exit For '/ j
                        Else
                            ReDim Preserve Liver(CheckPoint)
                            Liver(CheckPoint) = sp.Value
                            CheckPoint = CheckPoint + 1
                        End If
                    End If
            End Select
        Next
    Next
    '如果沒有定義到會是空的在ubound()時會發生錯誤。
    If CheckPoint <> 0 Then
        ReDim Preserve Liver(CheckPoint - 1)
    Else
        ReDim Preserve Liver(CheckPoint)
    End If
    
    For i = 0 To UBound(Liver)
        If Trim(Liver(i)) <> "" Then
            LiverFlag = True
            temp = temp & "    " & DiagnosisCount & "." & Liver(i) & vbCrLf
            DiagnosisCount = DiagnosisCount + 1
        End If
    Next
    
    If LiverFlag = True Then
        GetDIAGNOSIS = GetDIAGNOSIS & temp
    End If

    
    '=======GALLBLADDER========
    CheckPoint = 0
    temp = ""
    GALLBLADDERFlag = False
     For i = 119 To 122
        sp.row = i
        For j = 1 To sp.MaxCols
            sp.Col = j
            Select Case sp.CellType
                Case CellTypeCheckBox:
                    If sp.Value = "1" Then
                        sp.Col = sp.Col + 1
                        ReDim Preserve GALLBLADDER(CheckPoint)
                        GALLBLADDER(CheckPoint) = sp.Value
                        CheckPoint = CheckPoint + 1
                    End If
            End Select
        Next
    Next
    
    If CheckPoint <> 0 Then
        ReDim Preserve GALLBLADDER(CheckPoint - 1)
    Else
        ReDim Preserve GALLBLADDER(CheckPoint)
    End If
    
    For i = 0 To UBound(GALLBLADDER)
        If Trim(GALLBLADDER(i)) <> "" Then
            GALLBLADDERFlag = True
            temp = temp & "    " & DiagnosisCount & "." & GALLBLADDER(i) & vbCrLf
            DiagnosisCount = DiagnosisCount + 1
        End If
    Next

    If GALLBLADDERFlag = True Then
        GetDIAGNOSIS = GetDIAGNOSIS & temp
    End If

    '==================PANCREAS ==============
    CheckPoint = 0
    temp = ""
    PANCREASFlag = FALE
    
     For i = 124 To 126
        sp.row = i
        For j = 1 To sp.MaxCols
            sp.Col = j
            Select Case sp.CellType
                Case CellTypeCheckBox:
                    If sp.Value = "1" Then
                        sp.Col = sp.Col + 1
                        If i = 125 Or i = 126 Then
                            ReDim Preserve PANCREAS(CheckPoint)
                            '將這行全部組出來
                            PANCREAS(CheckPoint) = RowCheckBoxForOne(i, sp, " ")
                            CheckPoint = CheckPoint + 1
                            '/因為這行已經特殊處理完畢了不需要再處理。
                            Exit For '/ j
                        Else
                            ReDim Preserve PANCREAS(CheckPoint)
                            PANCREAS(CheckPoint) = sp.Value
                            CheckPoint = CheckPoint + 1
                        End If
                    End If
            End Select
        Next
    Next
    
    If CheckPoint <> 0 Then
        ReDim Preserve PANCREAS(CheckPoint - 1)
    Else
        ReDim Preserve PANCREAS(CheckPoint)
    End If
    
    For i = 0 To UBound(PANCREAS)
        If Trim(PANCREAS(i)) <> "" Then
            PANCREASFlag = True
            temp = temp & "    " & DiagnosisCount & "." & PANCREAS(i) & vbCrLf
            DiagnosisCount = DiagnosisCount + 1
        End If
    Next
    
    If PANCREASFlag = True Then
        GetDIAGNOSIS = GetDIAGNOSIS & temp
    End If
    
    '==========SPLEEN ============
    CheckPoint = 0
    temp = ""
    SPLEENFlag = False
     For i = 128 To 130
        sp.row = i
        For j = 1 To sp.MaxCols
            sp.Col = j
            Select Case sp.CellType
                Case CellTypeCheckBox:
                    If sp.Value = "1" Then
                        sp.Col = sp.Col + 1
                        If i = 128 Then
                            ReDim Preserve SPLEEN(CheckPoint)
                            '將這行全部組出來
                            SPLEEN(CheckPoint) = RowCheckBoxForOne(i, sp, " ")
                            CheckPoint = CheckPoint + 1
                            '/因為這行已經特殊處理完畢了不需要再處理。
                            Exit For '/ j
                        Else
                            ReDim Preserve SPLEEN(CheckPoint)
                            SPLEEN(CheckPoint) = sp.Value
                            CheckPoint = CheckPoint + 1
                        End If
                    End If
            End Select
        Next
    Next
    
    If CheckPoint <> 0 Then
        ReDim Preserve SPLEEN(CheckPoint - 1)
    Else
        ReDim Preserve SPLEEN(CheckPoint)
    End If
    
    For i = 0 To UBound(SPLEEN)
        If Trim(SPLEEN(i)) <> "" Then
            SPLEENFlag = True
            temp = temp & "    " & DiagnosisCount & "." & SPLEEN(i) & vbCrLf
            DiagnosisCount = DiagnosisCount + 1
        End If
    Next
    If SPLEENFlag = True Then
        GetDIAGNOSIS = GetDIAGNOSIS & temp
    End If
    '===============================
    
    CheckPoint = 0
    temp = ""
    KIDNEYFlag = False
     For i = 132 To 136
        sp.row = i
        For j = 1 To sp.MaxCols
            sp.Col = j
            Select Case sp.CellType
                Case CellTypeCheckBox:
                    If sp.Value = "1" Then
                        sp.Col = sp.Col + 1
                        If i = 134 Or i = 135 Or i = 136 Then
                            ReDim Preserve KIDNEY(CheckPoint)
                            '將這行全部組出來
                            KIDNEY(CheckPoint) = RowCheckBoxForOne(i, sp, " ")
                            CheckPoint = CheckPoint + 1
                            '/因為這行已經特殊處理完畢了不需要再處理。
                            Exit For '/ j
                        Else
                            ReDim Preserve KIDNEY(CheckPoint)
                            KIDNEY(CheckPoint) = sp.Value
                            CheckPoint = CheckPoint + 1
                        End If
                    End If
            End Select
        Next
    Next
    
    If CheckPoint <> 0 Then
        ReDim Preserve KIDNEY(CheckPoint - 1)
    Else
        ReDim Preserve KIDNEY(CheckPoint)
    End If
    
    
    For i = 0 To UBound(KIDNEY)
        If Trim(KIDNEY(i)) <> "" Then
            If InStr(1, UCase(Trim(KIDNEY(i))), "RIGHT") > 0 Or InStr(1, UCase(Trim(KIDNEY(i))), "LEFT") > 0 Then
                KIDNEY(i) = Replace(UCase(KIDNEY(i)), "(S)", "") & " KIDENEY "
            ElseIf InStr(1, UCase(Trim(KIDNEY(i))), "BOTH") > 0 Then
                KIDNEY(i) = Replace(UCase(KIDNEY(i)), ")", "")
                KIDNEY(i) = Replace(UCase(KIDNEY(i)), "(", "") & " KIDENYS "
            End If
            KIDNEYFlag = True
            temp = temp & "    " & DiagnosisCount & "." & KIDNEY(i) & vbCrLf
            DiagnosisCount = DiagnosisCount + 1
        End If
    Next
    
    '/如果有值在最前面加上KIDNEY
    If KIDNEYFlag = True Then
        GetDIAGNOSIS = GetDIAGNOSIS & temp
    End If
    
    Dim other As String
    other = GetOther(137, sp)
    
    If other <> "" Then
        GetDIAGNOSIS = GetDIAGNOSIS & "    " & DiagnosisCount & "." & other
    End If
    
    
    
End Function

'/將某一行的CheckBox 組成一個字串
Public Function RowCheckBoxForOne(ByVal i As Integer, ByRef sp As fpSpread, ByVal SpStr As String) As String
    sp.row = i
    For j = 1 To sp.MaxCols
        sp.Col = j
        If sp.Value = "1" Then
            sp.Col = sp.Col + 1
            RowCheckBoxForOne = RowCheckBoxForOne & sp.Value & SpStr
        End If
    Next
End Function


Public Function GetSUGGEST(ByRef sp As fpSpread) As String
    Dim SUGGEST() As String, SUGGESTFlag As Boolean
    Dim CheckPoint As Integer
    Dim temp As String
    SUGGESTFlag = False
    CheckPoint = 0
    Dim SUGGESTCount As Integer
    temp = ""
    SUGGESTCount = 1
    SUGGESTFlag = False
     For i = 138 To 140
        sp.row = i
        For j = 1 To sp.MaxCols
            sp.Col = j
            Select Case sp.CellType
                Case CellTypeCheckBox:
                    If sp.Value = "1" Then
                        sp.Col = sp.Col + 1
                        If i = 138 Or i = 139 Then
                            ReDim Preserve SUGGEST(CheckPoint)
                            '將這行全部組出來
                            SUGGEST(CheckPoint) = RowCheckBoxForOne(i, sp, ",")
                            CheckPoint = CheckPoint + 1
                            '/因為這行已經特殊處理完畢了不需要再處理。
                            Exit For '/ j
                        Else
                            ReDim Preserve SUGGEST(CheckPoint)
                            SUGGEST(CheckPoint) = sp.Value
                            CheckPoint = CheckPoint + 1
                        End If
                    End If
            End Select
        Next
    Next
    
    If CheckPoint <> 0 Then
        ReDim Preserve SUGGEST(CheckPoint - 1)
    Else
        ReDim Preserve SUGGEST(CheckPoint)
    End If
    
    
    For i = 0 To UBound(SUGGEST)
        If Trim(SUGGEST(i)) <> "" Then
            SUGGESTFlag = True
            
            If InStr(1, SUGGEST(i), "FOLLOW UP") > 0 Then
                If InStr(1, SUGGEST(i), "MONTHS") = 0 Then
                    If Right(SUGGEST(i), 1) = "," Then
                        SUGGEST(i) = Mid(SUGGEST(i), 1, Len(SUGGEST(i)) - 1) & " MONTHS LATER "
                    End If
                End If
            End If
            temp = temp & "    " & SUGGESTCount & "." & SUGGEST(i)
            If Right(temp, 1) = "," Then
                temp = Mid(temp, 1, Len(temp) - 1)
            End If
            SUGGESTCount = SUGGESTCount + 1
            temp = temp & vbCrLf
        End If
    Next
    Dim other As String
    other = GetOther(141, sp)
    If other <> "" Then
        temp = temp & "    " & SUGGESTCount & "." & other
    End If
    
    
    '/如果有值在最前面加上SUGGEST
    If SUGGESTFlag = True Or other <> "" Then
        GetSUGGEST = GetSUGGEST & "[SUGGEST]" & vbCrLf & temp
    End If
    
End Function

Public Sub Check2VbCrlf(ByRef Content As String)
    If Right(Content, 4) <> vbCrLf & vbCrLf Then
        If Right(Content, 2) = vbCrLf Then
            Content = Content & vbCrLf
        Else
            Content = Content & vbCrLf & vbCrLf
        End If
    End If
End Sub

'======================20130708新增EUS.rps 和 ERCP.rps輸出================================

Public Function Spread2TextEUS(ByVal uni_key As String, ByVal chartno As String, ByRef SpreadForm As fpSpread) As String
    Dim ReportTemp As String
    '================模板======================
    ReportTemp = "[Indication]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Indication#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Finding]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Finding#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Pre-medication]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Pre-medication#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Impression]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Impression#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Comment]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Comment#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Suggestion]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Suggestion#"
    '===============取值========================
    
    Dim Indication As String '/延用胃鏡的
    Indication = GetCheckBox(SpreadForm, 6)
    
    Dim finding As String
    SpreadForm.row = 8
    SpreadForm.Col = 11
    finding = SpreadForm.Value
    
    Dim pre_medication As String
    SpreadForm.row = 11
    SpreadForm.Col = 12
    Dim EUSPre As String
    EUSPre = "1.Gascon drop oral 10-15c.c." & vbCrLf
    EUSPre = EUSPre & "2.Xylocaine(10%) pump spray 1-3puffs"
    If UCase(Trim(SpreadForm.Text)) <> "NIL" And Trim(SpreadForm.Text <> "") Then
       EUSPre = EUSPre & vbCrLf & "3." & SpreadForm.Text
    End If
    
    pre_medication = EUSPre
    
    Dim impression As String
    SpreadForm.row = 13
    SpreadForm.Col = 11
    impression = SpreadForm.Value
    
    Dim comment As String
    SpreadForm.row = 17
    SpreadForm.Col = 11
    comment = SpreadForm.Value
    
    Dim suggestion As String
    SpreadForm.row = 21
    SpreadForm.Col = 11
    suggestion = SpreadForm.Value
    
    Call ReplaceReport("#Indication#", Indication, ReportTemp)
    Call ReplaceReport("#Finding#", finding, ReportTemp)
    Call ReplaceReport("#Pre-medication#", pre_medication, ReportTemp)
    Call ReplaceReport("#Impression#", impression, ReportTemp)
    Call ReplaceReport("#Comment#", comment, ReportTemp)
    Call ReplaceReport("#Suggestion#", suggestion, ReportTemp)
    
    Spread2TextEUS = ReportTemp
End Function

    
'======================20130708新增EUS.rps 和 ERCP.rps輸出================================

Public Function Spread2TextEUS_1(ByVal uni_key As String, ByVal chartno As String, ByRef SpreadForm As fpSpread) As String
    Dim ReportTemp As String
    '================模板======================
    ReportTemp = "[Indication]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Indication#" & vbCrLf
    
    ReportTemp = ReportTemp & "[EndoScopy Finding]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#EndoScopyFinding#" & vbCrLf
    
    ReportTemp = ReportTemp & "[EUSFinding]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#EUSFinding#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Pre-medication]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Pre-medication#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Impression]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Impression#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Comment]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Comment#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Suggestion]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Suggestion#"
    '===============取值========================
    
    Dim Indication As String '/延用胃鏡的
    Indication = GetCheckBox(SpreadForm, 5)
    
    Dim EndoScopyfinding As String
    SpreadForm.row = 7
    SpreadForm.Col = 11
    EndoScopyfinding = SpreadForm.Value
    
    Dim EUSfinding As String
    SpreadForm.row = 10
    SpreadForm.Col = 11
    EUSfinding = SpreadForm.Value
    
    Dim pre_medication As String
'    SpreadForm.row = 14
'    SpreadForm.Col = 12
    Dim EUSPre As String
    EUSPre = "1.Gascon drop oral 10-15c.c." & vbCrLf
    EUSPre = EUSPre & "2.Xylocaine(10%) pump spray 1-3puffs"
    
    '/記錄三個pre-medication
    Dim Pr1  As String, Pr2 As String, Pr3 As String
    '/記錄pre_medication 得編號,從3開始
    Dim PrCount As Integer
    PrCount = 3
    Dim l As Integer, m As Integer
    For l = 14 To 17
        SpreadForm.row = l
        For m = 1 To SpreadForm.MaxCols
            SpreadForm.Col = m
            If SpreadForm.CellType = CellTypeComboBox Then
                If UCase(Trim(SpreadForm.Text)) <> "NIL" And Trim(SpreadForm.Text <> "") Then
                    EUSPre = EUSPre & vbCrLf & PrCount & "." & SpreadForm.Text
                    PrCount = PrCount + 1
                End If
                Exit For
            End If
        Next
    Next
'    If UCase(Trim(SpreadForm.Text)) <> "NIL" And Trim(SpreadForm.Text <> "") Then
'       EUSPre = EUSPre & vbCrLf & "3." & SpreadForm.Text
'    End If
'
    pre_medication = EUSPre
    
    Dim impression As String
    SpreadForm.row = 18
    SpreadForm.Col = 11
    impression = SpreadForm.Value
    
    Dim comment As String
    SpreadForm.row = 22
    SpreadForm.Col = 11
    comment = SpreadForm.Value
    
    Dim suggestion As String
    SpreadForm.row = 26
    SpreadForm.Col = 11
    suggestion = SpreadForm.Value
    
    Call ReplaceReport("#Indication#", Indication, ReportTemp)
    Call ReplaceReport("#EndoScopyFinding#", EndoScopyfinding, ReportTemp)
    Call ReplaceReport("#EUSFinding#", EUSfinding, ReportTemp)
    Call ReplaceReport("#Pre-medication#", pre_medication, ReportTemp)
    Call ReplaceReport("#Impression#", impression, ReportTemp)
    Call ReplaceReport("#Comment#", comment, ReportTemp)
    Call ReplaceReport("#Suggestion#", suggestion, ReportTemp)
    
    Spread2TextEUS_1 = ReportTemp
End Function


Public Function Spread2TextEUS_2(ByVal uni_key As String, ByVal chartno As String, ByRef SpreadForm As fpSpread) As String
    Dim ReportTemp As String
    '================模板======================
    ReportTemp = "[Indication]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Indication#" & vbCrLf
    
    ReportTemp = ReportTemp & "[EndoScopy Finding]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#EndoScopyFinding#" & vbCrLf
    
    ReportTemp = ReportTemp & "[EUSFinding]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#EUSFinding#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Pre-medication]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Pre-medication#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Impression]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Impression#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Comment]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Comment#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Suggestion]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Suggestion#"
    '===============取值========================
    
    Dim Indication As String '/延用胃鏡的
    Indication = GetCheckBox(SpreadForm, 5)
    
    Dim pre_medication As String
'    SpreadForm.row = 14
'    SpreadForm.Col = 12
    Dim EUSPre As String
    EUSPre = "1.Gascon drop oral 10-15c.c." & vbCrLf
    EUSPre = EUSPre & "2.Xylocaine(10%) pump spray 1-3puffs"
    
    '/記錄三個pre-medication
    Dim Pr1  As String, Pr2 As String, Pr3 As String
    '/記錄pre_medication 得編號,從3開始
    Dim PrCount As Integer
    PrCount = 3
    Dim l As Integer, m As Integer
    For l = 7 To 9
        SpreadForm.row = l
        For m = 1 To SpreadForm.MaxCols
            SpreadForm.Col = m
            If SpreadForm.CellType = CellTypeComboBox Then
                If UCase(Trim(SpreadForm.Text)) <> "NIL" And Trim(SpreadForm.Text <> "") Then
                    EUSPre = EUSPre & vbCrLf & PrCount & "." & SpreadForm.Text
                    PrCount = PrCount + 1
                End If
                Exit For
            End If
        Next
    Next


'    If UCase(Trim(SpreadForm.Text)) <> "NIL" And Trim(SpreadForm.Text <> "") Then
'       EUSPre = EUSPre & vbCrLf & "3." & SpreadForm.Text
'    End If
    pre_medication = EUSPre
    
    Dim EndoScopyfinding As String
    SpreadForm.row = 11
    SpreadForm.Col = 11
    EndoScopyfinding = SpreadForm.Value
    
    
    Dim EUSfinding As String
    SpreadForm.row = 14
    SpreadForm.Col = 11
    EUSfinding = SpreadForm.Value
     
     
    Dim impression As String
    SpreadForm.row = 18
    SpreadForm.Col = 11
    impression = SpreadForm.Value
    
    Dim comment As String
    SpreadForm.row = 22
    SpreadForm.Col = 11
    comment = SpreadForm.Value
    
    Dim suggestion As String
    SpreadForm.row = 26
    SpreadForm.Col = 11
    suggestion = SpreadForm.Value
    
    Call ReplaceReport("#Indication#", Indication, ReportTemp)
    Call ReplaceReport("#EndoScopyFinding#", EndoScopyfinding, ReportTemp)
    Call ReplaceReport("#EUSFinding#", EUSfinding, ReportTemp)
    Call ReplaceReport("#Pre-medication#", pre_medication, ReportTemp)
    Call ReplaceReport("#Impression#", impression, ReportTemp)
    Call ReplaceReport("#Comment#", comment, ReportTemp)
    Call ReplaceReport("#Suggestion#", suggestion, ReportTemp)
    
    Spread2TextEUS_2 = ReportTemp
End Function




Public Function Spread2TextERCP(ByVal uni_key As String, ByVal chartno As String, ByRef SpreadForm As fpSpread) As String
    Dim ReportTemp As String
    '=======================模板================================
    ReportTemp = "[Indication]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Indication#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Pre-medication]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Pre-medication#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Finding]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#finding#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Diagnosis]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Diagnosis#" & vbCrLf

    ReportTemp = ReportTemp & "[Exam Limitation]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#exam_limitation#" & vbCrLf

    ReportTemp = ReportTemp & "[Complication]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Complication#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Suggestion]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Suggestion#" & vbCrLf
    
    '===================取值===================================
    Dim Indication As String
    Indication = GetCheckBox(SpreadForm, 6)
    
    Dim pre_medication As String
    SpreadForm.row = 8
    SpreadForm.Col = 13
    
'    pre_medication = "1.Gascon drop oral 10-15c.c." & vbCrLf
'    pre_medication = pre_medication & "2.Xylocaine(10%) pump spray 1-3puffs"
'    If UCase(Trim(SpreadForm.Text)) <> "NIL" And Trim(SpreadForm.Text <> "") Then
'       pre_medication = pre_medication & vbCrLf & "3." & SpreadForm.Text
'    End If
    pre_medication = SpreadForm.Text
    
    
    '/2013年8月26必須加上特定的字串
    
    
    
    Dim finding As String
    SpreadForm.row = 10
    SpreadForm.Col = 13
    finding = SpreadForm.Value
    
    Dim diagnosis As String
    SpreadForm.row = 13
    SpreadForm.Col = 13
    diagnosis = SpreadForm.Value
    
    Dim exam_limitation As String
    SpreadForm.row = 16
    SpreadForm.Col = 13
    exam_limitation = SpreadForm.Value
    
    Dim complication As String
    SpreadForm.row = 17
    SpreadForm.Col = 13
    complication = SpreadForm.Value
    
    Dim suggestion As String
    SpreadForm.row = 18
    SpreadForm.Col = 13
    suggestion = SpreadForm.Text
    
    Call ReplaceReport("#Indication#", Indication, ReportTemp)
    Call ReplaceReport("#Pre-medication#", pre_medication, ReportTemp)
    Call ReplaceReport("#finding#", finding, ReportTemp)
    Call ReplaceReport("#Diagnosis#", diagnosis, ReportTemp)
    Call ReplaceReport("#exam_limitation#", exam_limitation, ReportTemp)
    Call ReplaceReport("#Complication#", complication, ReportTemp)
    Call ReplaceReport("#Suggestion#", suggestion, ReportTemp)
    Spread2TextERCP = ReportTemp
    
End Function

Public Function Spread2TextERCP_1(ByVal uni_key As String, ByVal chartno As String, ByRef SpreadForm As fpSpread) As String
    Dim ReportTemp As String
    Dim temp$
    
    '=======================模板================================
    ReportTemp = "[Indication]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Indication#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Pre-medication]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Pre-medication#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Finding]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#finding#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Diagnosis]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Diagnosis#" & vbCrLf

    ReportTemp = ReportTemp & "[Exam Limitation]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#exam_limitation#" & vbCrLf

    ReportTemp = ReportTemp & "[Complication]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Complication#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Suggestion]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Suggestion#" & vbCrLf
    
    ReportTemp = ReportTemp & "[Other]" & vbCrLf
    ReportTemp = ReportTemp & "  " & "#Other#" & vbCrLf
    '===================取值===================================
    Dim Indication As String
    Indication = GetCheckBox_1(SpreadForm, 2, 4)
    
    Dim pre_medication As String, i As Integer
    
    SpreadForm.Col = 13
    pre_medication = ""
    For i = 6 To 8
        SpreadForm.row = i
        temp$ = Trim(SpreadForm.Text)
        If Len(temp$) > 0 And UCase(temp$) <> "NIL" Then
            If Len(pre_medication) > 0 Then
                pre_medication = pre_medication & ", "
            End If
            pre_medication = pre_medication & temp$
        End If
    Next
    If pre_medication = "" Then
        pre_medication = "Nil"
    End If
    
    Dim finding As String
    SpreadForm.row = 11
    SpreadForm.Col = 13
    finding = SpreadForm.Value
    
    Dim diagnosis As String
    SpreadForm.row = 14
    SpreadForm.Col = 13
    diagnosis = SpreadForm.Value
    
    Dim exam_limitation As String
    SpreadForm.row = 17
    SpreadForm.Col = 13
    exam_limitation = SpreadForm.Value
    
    Dim complication As String
    SpreadForm.row = 18
    SpreadForm.Col = 13
    complication = SpreadForm.Text
    
    Dim suggestion As String
    SpreadForm.row = 19
    SpreadForm.Col = 13
    suggestion = SpreadForm.Text
    
    Dim other As String
    SpreadForm.row = 20
    SpreadForm.Col = 13
    other = SpreadForm.Text
    
    Call ReplaceReport("#Indication#", Indication, ReportTemp)
    Call ReplaceReport("#Pre-medication#", pre_medication, ReportTemp)
    Call ReplaceReport("#finding#", finding, ReportTemp)
    Call ReplaceReport("#Diagnosis#", diagnosis, ReportTemp)
    Call ReplaceReport("#exam_limitation#", exam_limitation, ReportTemp)
    Call ReplaceReport("#Complication#", complication, ReportTemp)
    Call ReplaceReport("#Suggestion#", suggestion, ReportTemp)
    Call ReplaceReport("#Other#", other, ReportTemp)
    Spread2TextERCP_1 = ReportTemp
    
End Function


