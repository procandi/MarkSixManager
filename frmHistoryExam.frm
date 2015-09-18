VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Begin VB.Form frmHistoryExam 
   BorderStyle     =   1  '單線固定
   Caption         =   "歷史報告"
   ClientHeight    =   10965
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   15780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10965
   ScaleWidth      =   15780
   StartUpPosition =   3  '系統預設值
   Begin VSFlex7Ctl.VSFlexGrid VSFlexGrid1 
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   14535
      _cx             =   25638
      _cy             =   4471
      _ConvInfo       =   -1
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   16744576
      ForeColorFixed  =   16777215
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   "(Format)"
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   0   'False
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "帶入"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   14760
      Style           =   1  '圖片外觀
      TabIndex        =   3
      ToolTipText     =   "帶入舊記錄的資料"
      Top             =   120
      Width           =   855
   End
   Begin FPSpreadADO.fpSpread fpSpreadx 
      Height          =   8055
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   15495
      _Version        =   458752
      _ExtentX        =   27331
      _ExtentY        =   14208
      _StockProps     =   64
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
      SpreadDesigner  =   "frmHistoryExam.frx":0000
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8055
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   15495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "此記錄尚無報告"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   48
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   4200
      TabIndex        =   2
      Top             =   5520
      Width           =   6720
   End
End
Attribute VB_Name = "frmHistoryExam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const fields_max = 8
Dim aryRecord() As String
Dim aryResult()
Dim is_cloase As Boolean

Private Sub Command1_Click()
    Dim sql$, i As Integer
    
    nowUni_key$ = VSFlexGrid1.Cell(flexcpText, VSFlexGrid1.row, 2, VSFlexGrid1.row, 2)
    tmp$ = VSFlexGrid1.Cell(flexcpText, VSFlexGrid1.row, 8, VSFlexGrid1.row, 8)
    If UCase(tmp$) = UCase(curr_Record.TemplateFile) Then
        sql$ = "select * from CRIS_ABDOMINAL_CONTENT where uni_key = '" & nowUni_key$ & "' "
        sql$ = sql$ & " and chartno = '" & curr_Record.chartno & "' "
        Call OpenRecordset(sql$, Connection, Recordset)
        If Not Recordset.EOF Then
            For i = 0 To 11
                AContent(i) = NoNull(Recordset(Aindex(i, 1)))
            Next
            curr_Record.Item6 = Label2.Caption
            Unload Me
        Else
            MsgBox "查無此筆記錄的資料，無法帶入"
        End If
    Else
        MsgBox "所選的歷史記錄範本不同，不可帶入"
    End If
End Sub

Private Sub Form_Activate()
    If is_cloase Then
        Unload Me
    End If
End Sub

'Private Sub List1_Click()
'    Dim tmpPath$, tmp$, tmpSpread$
'
'    tmpPath$ = path_Images & "\Img" & Format(curr_Record.Date, "yyMM")
'    tmpPath$ = tmpPath$ & "\" & Trim(txtUni_key.Text) & ".rpt"
'
'    If isFileExist(tmpPath$, vbNormal) Then
'        ret = fpSpread1.LoadFromFile(tmpPath$)
'    End If
'End Sub

Private Sub Form_Load()
    Dim RecordsNo&, queueCaption$
    Dim dbgControl As Variant, adoControl As Variant
    
    Dim adoDB As New adoDB.Connection
    Dim adoOnline1 As New adoDB.Recordset
    Dim adoRecord1 As New adoDB.Recordset
    Dim sMode$, sqlSource$, tmpChartNo$
    Dim conn$, tmp$
    ReDim aryRecord(fields_max, 1000)
    
    Screen.MousePointer = vbHourglass
    DoEvents
    is_cloase = False
    
    If UCase(Trim(curr_Record.TemplateFile)) = "AU03261.RPS" Then
        Command1.Visible = True
        VSFlexGrid1.width = 14535
    Else
        Command1.Visible = False
        VSFlexGrid1.width = 15495
    End If
    
    adoDB.Open dbConnection$
    Set adoControl = adoOnline1
    '這邊抓取資料時，多抓取了一個欄位的資料(templatefile)，主要是保留若有需要再抓不到報告檔時，而可以根據這資料去抓取範本檔
    sqlSource$ = "select examdate, examtime,uni_key,  type, division_on, dr_report, dr_from, status, templatefile from cris_exam_online  where chartno = '" & curr_Record.chartno & "' "
    sqlSource$ = sqlSource$ & " and uni_key <> '" & curr_Record.uni_key & "' order by examdate DESC, examtime DESC"
    adoControl.Open sqlSource$, adoDB, adOpenForwardOnly, adLockReadOnly
    If adoControl.EOF Then
        MsgBox "查無歷史記錄，程式關閉"
        Command1.Visible = False
        is_cloase = True
    Else
        RecordsNo& = 0
        Do While Not adoControl.EOF
           If RecordsNo& > 1000 Then
              Exit Do
           End If
           
           For i% = 0 To adoControl.Fields.Count - 1
               aryRecord(i%, RecordsNo&) = NoNull(adoControl.Fields(i%))
           Next
           RecordsNo& = RecordsNo& + 1
           adoControl.MoveNext
        Loop
        
        adoControl.Close
        adoDB.Close
        Set adoControl = Nothing
        Set adoDB = Nothing
        
        ReDim aryResult(fields_max, RecordsNo& - 1)
        For j% = 0 To RecordsNo& - 1
                For i% = 0 To fields_max
                    aryResult(i%, j%) = aryRecord(i%, j%)
                Next
            Next
        DoEvents
        queueCaption$ = "檢查日期 　|時 間  |檢查編號        　|檢查類別        |檢查科室　   　|檢查醫師　   　|來源　   　|狀態  　  　|報告範本名稱 "
'        VSFlexGrid1.Rows = RecordsNo& - 1
'        VSFlexGrid1.Cols = fields_max
        VSFlexGrid1.FormatString = queueCaption$
        VSFlexGrid1.BindToArray aryResult
        
        VSFlexGrid1.Select 1, 1
    End If
    Me.ZOrder 0
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmHistoryExam.Visible = False
    currForm.Enabled = True
    currForm.Show
    Unload frmHistoryExam
End Sub

Private Function getdata(ByVal UniKey As String) As String
    Dim sql$, i As Integer, tmp$
    Dim BContent(11) As String
    
    sql$ = "select * from CRIS_ABDOMINAL_CONTENT where uni_key = '" & UniKey & "' "
    sql$ = sql$ & " and chartno = '" & curr_Record.chartno & "' "
    Call OpenRecordset(sql$, Connection, Recordset)
    If Not Recordset.EOF Then
        For i = 0 To 11
            BContent(i) = NoNull(Recordset(Aindex(i, 1)))
        Next
    Else
        For i = 0 To 11
            BContent(i) = ""
        Next
    End If
    tmp$ = ""
    For i = 0 To 11
        If Len(Trim(BContent(i))) > 0 Then
            If tmp$ <> "" And Right(tmp$, 1) <> Chr(10) Then
                tmp$ = tmp$ & vbCrLf
            End If
            tmp$ = tmp$ & "[" & Aindex(i, 0) & "]" & vbCrLf & BContent(i)
        ElseIf Len(Aindex(i, 2)) > 0 Then
            If tmp$ <> "" And Right(tmp$, 1) <> Chr(10) Then
                tmp$ = tmp$ & vbCrLf
            End If
            tmp$ = tmp$ & "[" & Aindex(i, 0) & "]" & vbCrLf & Aindex(i, 2)
        End If
    Next
    
    getdata = tmp$
    
End Function

Private Sub VSFlexGrid1_SelChange()
    Dim tmpSS7File$, tmp$, sql$
    
    nowUni_key$ = VSFlexGrid1.Cell(flexcpText, VSFlexGrid1.row, 2, VSFlexGrid1.row, 2)
    tmp$ = VSFlexGrid1.Cell(flexcpText, VSFlexGrid1.row, 8, VSFlexGrid1.row, 8)
    If UCase(tmp$) = "AU03261.RPS" Then
        fpSpreadx.Visible = False
        Label2.Caption = getdata(nowUni_key$)
        Command1.Enabled = True
        Label2.Visible = True
        Label2.ZOrder 0
    Else
        Command1.Enabled = False
        Label2.Caption = ""
        Label2.Visible = False
        tmpSS7File$ = path_Images & "\Img" & Format(VSFlexGrid1.Cell(flexcpText, VSFlexGrid1.row, 0, VSFlexGrid1.row, 0), "yyMM") & "\" & nowUni_key$ & ".rpt"
        If isFileExist(tmpSS7File$, vbNormal) Then
            ret = fpSpreadx.LoadFromFile(tmpSS7File$)
            fpSpreadx.Visible = True
        Else
            '無spread資料時，檢查是否有item6資料可顯示
            fpSpreadx.Visible = False
            sql$ = "select * from cris_exam_online where uni_key ='" & VSFlexGrid1.Cell(flexcpText, VSFlexGrid1.row, 2, VSFlexGrid1.row, 2) & "' and chartno = '" & curr_Record.chartno & "' "
            Call OpenRecordset(sql$, Connection, Recordset)
            If Not Recordset.EOF Then
                tmp$ = NoNull(Recordset("item6"))
                If Len(tmp$) > 0 Then
                    Label2.Visible = True
                    Label2.Caption = tmp$
                    Label2.ZOrder 0
                End If
            End If
        End If
        fpSpreadx.ColHeadersShow = False
        fpSpreadx.RowHeadersShow = False
    End If
End Sub
