VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Begin VB.Form frmCanSpread 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  '單線固定
   Caption         =   "套餐範本編輯/選用"
   ClientHeight    =   11175
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   15270
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11175
   ScaleWidth      =   15270
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "刪除"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6840
      Style           =   1  '圖片外觀
      TabIndex        =   13
      ToolTipText     =   "刪除目前選取的套餐範本"
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "重整"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      Style           =   1  '圖片外觀
      TabIndex        =   12
      ToolTipText     =   "根據目前醫師代號重新讀取套餐範本資料"
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtDr 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   11
      ToolTipText     =   "變更醫師為00000時，可另存新檔為通用範本"
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "傳回目前的套餐範本"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11760
      Style           =   1  '圖片外觀
      TabIndex        =   10
      ToolTipText     =   "將選取的套餐範本傳回"
      Top             =   10200
      Width           =   3375
   End
   Begin MSComDlg.CommonDialog CmnDlg 
      Left            =   0
      Top             =   8520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox lstCanTemplate 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6780
      Left            =   11760
      TabIndex        =   9
      ToolTipText     =   "雙擊以載入選取的套餐範本"
      Top             =   3360
      Width           =   3375
   End
   Begin VB.CommandButton cmdAddCanTemplate 
      BackColor       =   &H00FFC0C0&
      Caption         =   "另存新檔"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9960
      Style           =   1  '圖片外觀
      TabIndex        =   7
      ToolTipText     =   "將目前編輯中的範本存入新的套餐範本檔"
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdSaveCanTemplate 
      BackColor       =   &H00FFC0C0&
      Caption         =   "儲存"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8280
      Style           =   1  '圖片外觀
      TabIndex        =   6
      ToolTipText     =   "將目前編輯中的範本存回選取的套餐範本檔內"
      Top             =   120
      Width           =   1575
   End
   Begin VB.ListBox lstSourceTemplate 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      Left            =   11760
      TabIndex        =   5
      ToolTipText     =   "雙擊以載入選取的空白範本"
      Top             =   480
      Width           =   3375
   End
   Begin FPSpreadADO.fpSpread fpSpread1 
      Height          =   10125
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   11520
      _Version        =   458752
      _ExtentX        =   20320
      _ExtentY        =   17859
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frmCanSpread.frx":0000
   End
   Begin VB.Label lblSpread 
      Height          =   375
      Index           =   3
      Left            =   10800
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblSpread 
      Height          =   375
      Index           =   2
      Left            =   10800
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblSpread 
      Height          =   375
      Index           =   1
      Left            =   10800
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblSpread 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   14
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "套餐範本"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   11760
      TabIndex        =   8
      Top             =   3000
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "原始空白範本"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   11760
      TabIndex        =   4
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label lbType 
      BackColor       =   &H00808080&
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   1200
      TabIndex        =   3
      Top             =   480
      Width           =   4485
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "類別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1125
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "醫師"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1125
   End
End
Attribute VB_Name = "frmCanSpread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xSpreadCanName As String     '用於傳回新增範本名稱
Public txtExamType As String        '用於片語功能使用
Public is_Open As Boolean           '用於片語功能使用
Dim xTemplateFile() As String       '用於記錄原始範本資料
Dim currTemplate(8) As String       '用於記錄目前使用的範本資料
'0  :Source_filePath        :原來的範本的路徑
'1  :SOURCE_FILENAME        :原來的範本的檔名
'2  :SOURCE_EXAMNAME        :原來的範本ExamName
'3  :SOURCE_EXAMID          :原來的範本ExamID
'4  :SOURCE_EXAMDESCRIPTION :原來的範本EXAMDescription
'5  :SPREAD_FILENAME        :套餐範本的檔案名稱
'6  :SPREAD_NAME            :套餐範本名稱
'7  :USERID                 :醫師代碼
'8  :template_type          :目前範本來源，0為原始範本，1為套餐範本
Dim FcurrTemplate As Boolean        '用於記錄目前範本是否修改後已存
Dim yTemplateFile() As String       '用於紀錄USER套餐範本資料
Dim has_yTemplate As Boolean        '用於記錄是否有USER套餐範本資料
Dim xCurrForm
'Public lblSpread(3) As String          '用於記錄片語的回傳位置

'設定spread值
Public Sub setSPtext(ByVal sheet As Long, ByVal Col As Long, ByVal row As Long, ByVal tvalue As String)
    fpSpread1.sheet = sheet
    fpSpread1.row = row
    fpSpread1.Col = Col
    If Trim(tvalue) = "" Then
        fpSpread1.Text = ""
    Else
        fpSpread1.Text = tvalue
    End If
End Sub

'取得spread內的文字
Public Function getSPtext(ByVal sheet As Long, ByVal Col As Long, ByVal row As Long) As String
    fpSpread1.sheet = sheet
    fpSpread1.Col = Col
    fpSpread1.row = row
    getSPtext = Trim(fpSpread1.Text)
End Function

Private Sub cmdPhrase_Click()
    Dim t$
    Dim i As Integer
    
    If Val(lblSpread(0)) + Val(lblSpread(1)) > 0 Then
        Set xCurrForm = currForm
        Set currForm = Me
        Load frmPhraseSpread
        Me.Enabled = False
        frmPhraseSpread.Show
        
    End If

End Sub

'新增範本名稱
Private Sub cmdAddCanTemplate_Click()
    Dim xflag As Boolean
    Dim sql$
    
    cmdAddCanTemplate.Enabled = False
    xSpreadCanName = ""
    Me.Hide
    frmCanSpreadDialog.Show 1
    Me.Show

    If xSpreadCanName <> "" Then
        xflag = True
        '先檢查是否有USER套餐範本資料
        If has_yTemplate Then
            '檢查是否有重複名稱的範本
            For i = 0 To UBound(yTemplateFile)
                If yTemplateFile(i, 5) = xSpreadCanName Then
                    xflag = False
                    Exit For
                End If
            Next
        End If
        
        '若無重複名稱，新增USER範本
        If xflag Then
            currTemplate(6) = xSpreadCanName
            '新的套餐範本檔名為USER命名的名稱+年月日.rps
            currTemplate(5) = xSpreadCanName & Format(Now, "YYYYMMDDhhmmss") & ".rps"
            currTemplate(7) = Trim(txtDr.Text)
            If Not FSO.FolderExists(currTemplate(0) & currTemplate(7)) Then
                Call FSO.CreateFolder(currTemplate(0) & currTemplate(7))
            End If
            '套餐範本是儲存在source_filepath & userid\的目錄之下
            Call fpSpread1.SaveToFile(currTemplate(0) & currTemplate(7) & "\" & currTemplate(5), False)
            
            sql$ = "Insert into cris_SpreadForm_List ( "
            sql$ = sql$ & " USERID, SPREAD_NAME, SPREAD_FILENAME, EXAM_TYPE, DIVISION_NAME, SOURCE_FILENAME, "
            sql$ = sql$ & " SOURCE_FILEPATH, SOURCE_EXAMNAME, SOURCE_EXAMID, SOURCE_EXAMDESCRIPTION "
            sql$ = sql$ & ") values ( "
            sql$ = sql$ & " '" & currTemplate(7) & "', '" & currTemplate(6) & "', '" & currTemplate(5) & "', "
            sql$ = sql$ & " '" & lbType.Caption & "', '" & Trim(curr_Record.Division_on) & "', "
            sql$ = sql$ & " '" & currTemplate(1) & "', '" & currTemplate(0) & "', '" & currTemplate(2) & "', "
            sql$ = sql$ & " '" & currTemplate(3) & "', '" & currTemplate(4) & "' )"
            Call Connection.Execute(sql$)
            currTemplate(8) = "1"
            '重新載入套餐範本
            Call Load_Can_Tempalte
            frmCanSpread.Caption = "套餐範本編輯/選用 - 套餐範本: " & currTemplate(6)
        Else
            MsgBox "檔名重複，請另外取名"
        End If
        
    End If
    cmdAddCanTemplate.Enabled = True
End Sub

'存回USER套餐範本
Private Sub cmdSaveCanTemplate_Click()
    cmdSaveCanTemplate.Enabled = False
    '先檢查目前狀態是否有載入USER套餐範本；若為原始範本則跳新增套餐範本名稱
    If currTemplate(8) = "0" Then
        Call cmdAddCanTemplate_Click
    Else
        '是否有修改過，有修改過時才需匯出檔案，不需修改資料表
        If Not FcurrTemplate Then
            '套餐範本是儲存在source_filepath & userid\的目錄之下
            Call fpSpread1.SaveToFile(currTemplate(0) & currTemplate(7) & "\" & currTemplate(5), False)
        End If
    End If
    cmdSaveCanTemplate.Enabled = True
End Sub

'重新載入套餐範本資料
Private Sub Command1_Click()
    Command1.Enabled = False
    Call Load_Can_Tempalte
    
    Command1.Enabled = True
End Sub

'刪除目前選取的範本記錄
Private Sub Command2_Click()
    Dim sql$
    
    Command2.Enabled = False
    '只有套餐範本才可以刪除，原始範本不可刪除
    If currTemplate(8) = "1" Then
        sql$ = "delete from cris_SpreadForm_List "
        sql$ = sql$ & " where division_name='" & Trim(curr_Record.Division_on) & "' "
        sql$ = sql$ & " and userID = '" & currTemplate(7) & "' "
        sql$ = sql$ & " and exam_type = '" & lbType.Caption & "' "
        sql$ = sql$ & " and Spread_Name = '" & currTemplate(6) & "' "
        Call Connection.Execute(sql$)
        Call Load_Can_Tempalte
        If lstCanTemplate.ListCount > 0 Then
            lstCanTemplate.ListIndex = 0
            Call lstCanTemplate_DblClick
        End If
    End If
    Command2.Enabled = True
End Sub

'傳回選取的套餐範本
Private Sub Command3_Click()
    Dim tPath$
    
    Command3.Enabled = False
    '將目前開啟的範本存入curr_record.uni_key+.rps檔案內
    tPath$ = App.Path
    If Right(tPath$, 1) <> "\" Then
        tPath$ = tPath$ & "\"
    End If
    tPath$ = tPath$ & "temp"
    If Not FSO.FolderExists(tPath$) Then
        Call FSO.CreateFolder(tPath$)
    End If
    Call fpSpread1.SaveToFile(tPath$ & "\" & Trim(curr_Record.uni_key) & ".rps", False)
    Spread_ID = currTemplate(3)
    Spread_Name = currTemplate(2)
    curr_Record.TemplateFile = currTemplate(1)
    curr_Record.TemplateName = currTemplate(4)
    Call currForm.fpSpread1.LoadFromFile(tPath$ & "\" & Trim(curr_Record.uni_key) & ".rps")
    '設定curr_form的狀態與curr_record的設定值
    Command3.Enabled = True
    currForm.fpSpread1.ChangeMade = True
    currForm.Enabled = True
    Unload Me
End Sub

'載入原始範本
Private Sub Load_Source_Tempalte()
    Dim sql$
    Dim rcount As Integer
    Dim i As Integer
    
    '撈入原始範本
    lstSourceTemplate.Clear
    sql$ = "SELECT * FROM CRIS_ReportTemplate "
    sql$ = sql$ & " where divisionname='" & Trim(curr_Record.Division_on) & "' "
    sql$ = sql$ & " ORDER BY DivisionName, ExamID"
    Call OpenRecordset(sql$, Connection, Recordset)
    rcount = 0
    If Not Recordset.EOF Then
        While Not Recordset.EOF
            rcount = rcount + 1
            Recordset.MoveNext
        Wend
        Recordset.MoveFirst
        ReDim xTemplateFile(rcount - 1, 4)
        i = 0
        While Not Recordset.EOF
            If FSO.FileExists(Recordset("templatefilesource") & Recordset("templatefilename")) Then
                lstSourceTemplate.AddItem Recordset("examID") & " - " & Recordset("examname")
                '原來的範本的路徑
                xTemplateFile(i, 0) = Recordset("templatefilesource")
                '原來的範本ExamID
                xTemplateFile(i, 1) = Recordset("ExamID")
                '原來的範本ExamName
                xTemplateFile(i, 2) = Recordset("ExamName")
                '原來的範本EXAMDescription
                xTemplateFile(i, 3) = Recordset("ExamDescription")
                '原來的範本的檔名
                xTemplateFile(i, 4) = Recordset("templatefilename")
                i = i + 1
            End If
            Recordset.MoveNext
        Wend
    End If
End Sub

'載入套餐範本
Private Sub Load_Can_Tempalte()
    Dim sql$
    Dim rcount As Integer
    Dim i As Integer
    
    lstCanTemplate.Clear
    has_yTemplate = False
    sql$ = "SELECT * FROM cris_SpreadForm_List "
    sql$ = sql$ & " where division_name='" & Trim(curr_Record.Division_on) & "' "
    sql$ = sql$ & " and (userID = '" & Trim(txtDr.Text) & "' or userID = '00000') "
    sql$ = sql$ & " and exam_type = '" & lbType.Caption & "' "
    sql$ = sql$ & " ORDER BY Division_Name, userID, Spread_Name"
    Call OpenRecordset(sql$, Connection, Recordset)
    rcount = 0
    If Not Recordset.EOF Then
        While Not Recordset.EOF
            rcount = rcount + 1
            Recordset.MoveNext
        Wend
        Recordset.MoveFirst
        ReDim yTemplateFile(rcount - 1, 9)
        i = 0
        While Not Recordset.EOF
            'USER代碼
            yTemplateFile(i, 0) = NoNull(Recordset("UserID"))
            '原來的範本EXAMDescription
            yTemplateFile(i, 1) = NoNull(Recordset("Source_ExamDescription"))
            '原來的範本ExamName
            yTemplateFile(i, 2) = NoNull(Recordset("Source_ExamName"))
            '原來的範本ExamID
            yTemplateFile(i, 3) = NoNull(Recordset("Source_ExamID"))
            '原來的範本的檔名
            yTemplateFile(i, 4) = NoNull(Recordset("Source_FileName"))
            '原來的範本的檔案路徑
            yTemplateFile(i, 5) = NoNull(Recordset("Source_FilePath"))
            '科室名稱
            yTemplateFile(i, 6) = NoNull(Recordset("Division_Name"))
            '檢查類別
            yTemplateFile(i, 7) = NoNull(Recordset("Exam_Type"))
            '套餐範本名稱
            yTemplateFile(i, 8) = NoNull(Recordset("Spread_Name"))
            '套餐範本的檔案名稱
            yTemplateFile(i, 9) = NoNull(Recordset("Spread_FileName"))
            lstCanTemplate.AddItem yTemplateFile(i, 0) & " - " & yTemplateFile(i, 8)
            i = i + 1
            Recordset.MoveNext
        Wend
        has_yTemplate = True
    End If
End Sub

Private Sub Form_Activate()
    Set currForm = xCurrForm
    
End Sub

Private Sub Form_Load()
    Dim sql$
    Dim rcount As Integer
    Dim i As Integer
    
    Set xCurrForm = currForm
    txtDr.Text = frmQueue.lblUser.Caption
    lbType.Caption = curr_Record.Type
    txtExamType = curr_Record.Type
    is_Open = False
    
    '撈入原始範本
    Call Load_Source_Tempalte
    
    '撈入USER的套餐範本
    Call Load_Can_Tempalte
            
    '載入預設的原始範本與設定初值
    If lstSourceTemplate.ListCount > 0 Then
        lstSourceTemplate.ListIndex = 0
        Call lstSourceTemplate_DblClick
    ElseIf lstCanTemplate.ListCount > 0 Then
        lstCanTemplate.ListIndex = 0
        Call lstCanTemplate_DblClick
    End If

    FcurrTemplate = True
    Command3.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    currForm.Enabled = True
    currForm.Show
    Unload Me
End Sub

Private Sub fpSpread1_Change(ByVal Col As Long, ByVal row As Long)
    FcurrTemplate = False
End Sub

Private Sub fpSpread1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    Dim tmpTxtRowCol$, txtCol&, txtRow&
           
    fpSpread1.Col = Col
    fpSpread1.row = row
    fpSpread1.sheet = fpSpread1.ActiveSheet
    tmpTxtRowCol$ = fpSpread1.CellTag
    If Len(tmpTxtRowCol$) = 4 Then
       txtRow& = Val(Left(tmpTxtRowCol$, 2))
       txtCol& = Val(Right(tmpTxtRowCol$, 2))
        
       fpSpread1.Col = txtCol&
       fpSpread1.row = txtRow&
       
       lblSpread(0) = fpSpread1.row
       lblSpread(1) = fpSpread1.Col
       lblSpread(2) = fpSpread1.CellNote
       'lblSpread(3) = fpSpread1.SelStart
    Else
       lblSpread(0) = ""
       lblSpread(1) = ""
       lblSpread(2) = ""
       'lblSpread(3) = ""
    
    End If
        '--------------------------------------------------------------

    If fpSpread1.CellType = CellTypeEdit Then
       DoEvents
       
       Call cmdPhrase_Click
    End If
End Sub

'USER雙擊套餐範本
Private Sub lstCanTemplate_DblClick()
    Dim temp$
    
    lstCanTemplate.Enabled = False
    '載入套餐範本
    If lstCanTemplate.ListIndex >= 0 Then
        If FSO.FileExists(yTemplateFile(lstCanTemplate.ListIndex, 5) & yTemplateFile(lstCanTemplate.ListIndex, 0) & "\" & yTemplateFile(lstCanTemplate.ListIndex, 9)) Then
            Call fpSpread1.LoadFromFile(yTemplateFile(lstCanTemplate.ListIndex, 5) & yTemplateFile(lstCanTemplate.ListIndex, 0) & "\" & yTemplateFile(lstCanTemplate.ListIndex, 9))
            fpSpread1.ChangeMade = True
            fpSpread1.ColHeadersShow = False
            fpSpread1.RowHeadersShow = False
            '設定初值
            currTemplate(0) = yTemplateFile(lstCanTemplate.ListIndex, 5)
            currTemplate(1) = yTemplateFile(lstCanTemplate.ListIndex, 4)
            currTemplate(2) = yTemplateFile(lstCanTemplate.ListIndex, 2)
            currTemplate(3) = yTemplateFile(lstCanTemplate.ListIndex, 3)
            currTemplate(4) = yTemplateFile(lstCanTemplate.ListIndex, 1)
            currTemplate(5) = yTemplateFile(lstCanTemplate.ListIndex, 9)
            currTemplate(6) = yTemplateFile(lstCanTemplate.ListIndex, 8)
            currTemplate(7) = yTemplateFile(lstCanTemplate.ListIndex, 0)
            currTemplate(8) = "1"
            frmCanSpread.Caption = "套餐範本編輯/選用 - 套餐範本: " & currTemplate(6)
            Command2.Enabled = True
            FcurrTemplate = True
        Else
            MsgBox "找不到套餐範本檔案，請聯絡軒崴工程師!"
            temp$ = "遺失套餐範本檔案: "
            temp$ = temp$ & yTemplateFile(lstCanTemplate.ListIndex, 5) & yTemplateFile(lstCanTemplate.ListIndex, 0) & "\" & yTemplateFile(lstCanTemplate.ListIndex, 9)
            Log (temp$)
        End If
    End If
    lstCanTemplate.Enabled = True
End Sub

'USER雙擊原始範本
Private Sub lstSourceTemplate_DblClick()
    lstSourceTemplate.Enabled = False
    '載入原始範本
    If lstSourceTemplate.ListIndex >= 0 Then
        If FSO.FileExists(xTemplateFile(lstSourceTemplate.ListIndex, 0) & xTemplateFile(lstSourceTemplate.ListIndex, 4)) Then
            Call fpSpread1.LoadFromFile(xTemplateFile(lstSourceTemplate.ListIndex, 0) & xTemplateFile(lstSourceTemplate.ListIndex, 4))
            fpSpread1.ChangeMade = True
            fpSpread1.ColHeadersShow = False
            fpSpread1.RowHeadersShow = False
            '設定初值
            currTemplate(0) = xTemplateFile(lstSourceTemplate.ListIndex, 0)
            currTemplate(1) = xTemplateFile(lstSourceTemplate.ListIndex, 4)
            currTemplate(2) = xTemplateFile(lstSourceTemplate.ListIndex, 2)
            currTemplate(3) = xTemplateFile(lstSourceTemplate.ListIndex, 1)
            currTemplate(4) = xTemplateFile(lstSourceTemplate.ListIndex, 3)
            currTemplate(5) = ""
            currTemplate(6) = ""
            currTemplate(7) = Trim(txtDr.Text)
            currTemplate(8) = "0"
            frmCanSpread.Caption = "套餐範本編輯/選用 - 原始範本: " & currTemplate(3) & " - " & currTemplate(2)
            FcurrTemplate = True
            '原始範本不可刪除
            Command2.Enabled = False
        Else
            MsgBox "找不到原始範本檔案，請聯絡軒崴工程師!"
            temp$ = "遺失原始範本檔案: "
            temp$ = temp$ & xTemplateFile(lstSourceTemplate.ListIndex, 0) & xTemplateFile(lstSourceTemplate.ListIndex, 4)
            Log (temp$)
        End If
    End If
    lstSourceTemplate.Enabled = True
End Sub
