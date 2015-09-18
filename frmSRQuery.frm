VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSRQuery 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Structed Report Query"
   ClientHeight    =   10860
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10860
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkBasic 
      BackColor       =   &H00E0E0E0&
      Caption         =   "只查詢基本資料"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   17
      ToolTipText     =   "勾選時不查詢或顯示任何Structed Report資料"
      Top             =   120
      Width           =   5415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "刪除條件"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   4
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "加入條件"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   3
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   120
      Width           =   975
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      ItemData        =   "frmSRQuery.frx":0000
      Left            =   6360
      List            =   "frmSRQuery.frx":0002
      TabIndex        =   14
      Top             =   480
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   1800
      TabIndex        =   10
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   1800
      TabIndex        =   11
      Top             =   1920
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      ItemData        =   "frmSRQuery.frx":0004
      Left            =   1800
      List            =   "frmSRQuery.frx":0020
      TabIndex        =   7
      Top             =   720
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "離開"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   2
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "查詢"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Index           =   1
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "輸出Excel檔"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   0
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      ItemData        =   "frmSRQuery.frx":0062
      Left            =   1800
      List            =   "frmSRQuery.frx":00F0
      TabIndex        =   2
      Top             =   120
      Width           =   3375
   End
   Begin VSFlex7Ctl.VSFlexGrid VSFlexGrid1 
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   15015
      _cx             =   26485
      _cy             =   14631
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   25
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Index           =   0
      Left            =   1800
      TabIndex        =   12
      Top             =   1320
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   26083329
      CurrentDate     =   41159
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   13
      Top             =   1920
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   26083329
      CurrentDate     =   41159
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "判斷值2："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   1980
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "判斷值1："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   1380
      Width           =   1410
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "邏輯運算："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   780
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "查詢欄位："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   1575
   End
End
Attribute VB_Name = "frmSRQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mcombo1_text As String
Private Const fields_max = 48
Dim aryRecord()
Dim aryResult()

Private Sub Combo1_Click()
    If Mcombo1_text <> Combo1.Text Then
        Text1(0).Text = ""
        Text1(0).Visible = True
        Text1(1).Text = ""
        Text1(1).Visible = False
        Label1(3).Visible = False
        Combo2.Text = ""
        DTPicker1(0).Visible = False
        DTPicker1(1).Visible = False
    
        Select Case Combo1.Text
            Case "檢查日期"
                DTPicker1(0).Visible = True
                Text1(0).Text = Format(DTPicker1(0).Value, "yyyy/MM/dd")
                Text1(0).Enabled = False
                Text1(1).Text = Format(DTPicker1(1).Value, "yyyy/MM/dd")
                Text1(1).Enabled = False
            Case Else
                Text1(0).Enabled = True
                Text1(1).Enabled = True
        End Select
    End If
End Sub

Private Sub Combo2_Click()
    If Combo2.Text = "介於之間" Then
        Text1(1).Visible = True
        Label1(3).Visible = True
        If Combo1.Text = "檢查日期" Then
            DTPicker1(1).Visible = True
            Text1(1).Enabled = False
        Else
            DTPicker1(1).Visible = False
            Text1(1).Enabled = True
        End If
    Else
        Text1(1).Visible = False
        Label1(3).Visible = False
        DTPicker1(1).Visible = False
    End If
End Sub

Private Sub ADDcondition()
    Dim i%, j%
    Dim fieldName$, fcondition$, tmpp1$, tmpp2$
    
    i% = 0
    Select Case Combo1.Text
        Case "病歷號"
            i% = 1
            fieldName$ = "a.chartno"
        Case "姓名"
            fieldName$ = "b.name"
            i% = 1
        Case "年齡"
            fieldName$ = "a.age"
            i% = 2
        Case "檢查日期"
            fieldName$ = "a.examdate"
            i% = 1
        Case "性別"
            fieldName$ = "b.sex"
            i% = 1
        Case "醫師別"
            fieldName$ = "a.Dr_report"
            i% = 1
        Case "檢查類別"
            fieldName$ = "a.type"
            i% = 1
        Case "Conclusion"
            fieldName$ = "c.Conclusion"
            i% = 1
        Case "Impression"
            fieldName$ = "c.Impression"
            i% = 1
        Case "IVSd"
            fieldName$ = "c.IVSD"
            i% = 2
        Case "LVIDd"
            fieldName$ = "c.LVIDD"
            i% = 2
        Case "LVPWd"
            fieldName$ = "c.LVPWD"
            i% = 2
        Case "IVSs"
            fieldName$ = "c.IVSS"
            i% = 2
        Case "LVIDs"
            fieldName$ = "c.LVIDS"
            i% = 2
        Case "LVPWs"
            fieldName$ = "c.LVPWS"
            i% = 2
        Case "LVOT diam"
            fieldName$ = "c.LVOT_DIAM"
            i% = 2
        Case "Ao root diam"
            fieldName$ = "c.AO_ROOT_DIAM"
            i% = 2
        Case "ACS"
            fieldName$ = "c.ACS"
            i% = 2
        Case "LA dimension"
            fieldName$ = "c.LA_DIMENSION"
            i% = 2
        Case "EF(Teich)"
            fieldName$ = "c.EF_TEICH"
            i% = 2
        Case "EF(Simpson's)"
            fieldName$ = "c.EF_SIMPOSNS"
            i% = 2
        Case "LVOT max"
            fieldName$ = "c.LVOT_MAX"
            i% = 2
        Case "AV max"
            fieldName$ = "c.AV_MAX"
            i% = 2
        Case "AV max PG"
            fieldName$ = "c.AV_MAX_PG"
            i% = 2
        Case "AV mean PG"
            fieldName$ = "c.AVMEANPG"
            i% = 2
        Case "AVA"
            fieldName$ = "c.AVA"
            i% = 2
        Case "AI Vmax"
            fieldName$ = "c.AI_VMAX"
            i% = 2
        Case "MV E point"
            fieldName$ = "c.MV_E_POINT"
            i% = 2
        Case "MV A point"
            fieldName$ = "c.MV_A_POINT"
            i% = 2
        Case "MV E/A"
            fieldName$ = "c.MV_EA"
            i% = 2
        Case "MVA(P1/2t)"
            fieldName$ = "c.MVA_P21"
            i% = 2
        Case "Max vel(TR)"
            fieldName$ = "c.MAX_VELTR"
            i% = 2
        Case "Max PG(TR)"
            fieldName$ = "c.MAX_PGTR"
            i% = 2
        Case "RA pressure"
            fieldName$ = "c.RA_PRESSURE"
            i% = 2
        Case "RVSP(TR)"
            fieldName$ = "c.RVSP_TR"
            i% = 2
        Case "PA Vmax"
            fieldName$ = "c.PA_VMAX"
            i% = 2
        Case "PA Max PG"
            fieldName$ = "c.PA_MAX_PG"
            i% = 2
        Case "PA Mean PG"
            fieldName$ = "c.PA_MEAN_PG"
            i% = 2
        Case "PR Vmax"
            fieldName$ = "c.PR_VMAX"
            i% = 2
        Case "RVOT Diam"
            fieldName$ = "c.RVOT_DIAM"
            i% = 2
        Case "TDI E/A"
            fieldName$ = "c.TDI_EA"
            i% = 2
        Case "QP/QS"
            fieldName$ = "c.QP_QS"
            i% = 2
        Case "LA Volume(BP)"
            fieldName$ = "c.LA_VOLUMEBP"
            i% = 2
        Case "MPI"
            fieldName$ = "c.MPI"
            i% = 2
        Case "LV mass"
            fieldName$ = "c.LV_MAX"
            i% = 2
        Case "RA dimension"
            fieldName$ = "c.ITEM1_VALUE"
            i% = 2
        Case "RV dimension"
            fieldName$ = "c.ITEM2_VALUE"
            i% = 2
        Case Else
            MsgBox "查詢欄位名稱有誤，無法查詢，請排除後再試"
            Exit Sub
    End Select
    
    fcondition$ = ""
    If i% = 1 Then
        tmpp1$ = "'" & Text1(0).Text & "' "
        tmpp2$ = "'" & Text1(1).Text & "' "
    ElseIf i% = 2 Then
        tmpp1$ = Text1(0).Text
        tmpp2$ = Text1(1).Text
    End If
    Select Case Combo2.Text
        Case "大於"
            fcondition$ = fieldName$ & " > " & tmpp1$
        Case "等於"
            fcondition$ = fieldName$ & " = " & tmpp1$
        Case "小於"
            fcondition$ = fieldName$ & " < " & tmpp1$
        Case "不等於"
            fcondition$ = fieldName$ & " <> " & tmpp1$
        Case "大於等於"
            fcondition$ = fieldName$ & " >= " & tmpp1$
        Case "小於等於"
            fcondition$ = fieldName$ & " <= " & tmpp1$
        Case "包含"
            fcondition$ = fieldName$ & " like '%" & Text1(0).Text & "%' "
        Case "介於之間"
            fcondition$ = fieldName$ & " >= " & tmpp1$ & " and "
            fcondition$ = fcondition$ & fieldName$ & " <= " & tmpp2$
        Case Else
            MsgBox "邏輯運算名稱有誤，無法查詢，請排除後再試"
            Exit Sub
    End Select
    List1.AddItem fcondition$
End Sub

Private Sub Query_SR()
    Dim dbgControl As Variant, adoControl As Variant
    Dim RecordsNo&
    Dim i%, j%
    Dim adoDB As New adoDB.Connection
    Dim adoOnline1 As New adoDB.Recordset
    Dim fieldName$, fcondition$, SQL$, tmpp1$, tmpp2$
    
    On Error GoTo ttt
    If List1.ListCount > 0 Then
        fcondition$ = ""
        For i% = 0 To List1.ListCount - 1
            If chkBasic.Value = 0 Then      '查詢全部
                fcondition$ = fcondition$ & List1.List(i%) & " and "
            ElseIf Mid(List1.List(i%), 1, 1) <> "c" Then    '只查基本資料，所以c開頭為cris_smart_save之資料不列入條件
                fcondition$ = fcondition$ & List1.List(i%) & " and "
            End If
        Next
        fcondition$ = Left(fcondition$, Len(fcondition$) - 4)
    Else
        MsgBox "查無條件，查詢取消"
        Exit Sub
    End If
    
    '/**/
    'SQL$ = "select a.chartno, b.name,a.age, b.sex, a.type, a.examdate, a.examtime, a.dr_report "
    'If chkBasic.Value = 0 Then
    '    SQL$ = SQL$ & ", c.conclusion, c.impression "
    '    SQL$ = SQL$ & ", c.IVSD, c.LVIDD, c.LVPWD, c.IVSS, c.LVIDS, c.LVPWS, c.LVOT_DIAM, c.AO_ROOT_DIAM, c.ACS, c.LA_DIMENSION "
    '    SQL$ = SQL$ & ", c.EF_TEICH, c.EF_SIMPOSNS, c.LVOT_MAX, c.AV_MAX, c.AV_MAX_PG, c.AVMEANPG, c.AVA, c.AI_VMAX "
    '    SQL$ = SQL$ & ", c.MV_E_POINT, c.MV_A_POINT, c.MV_EA, c.MVA_P21 "
    '    SQL$ = SQL$ & ", c.MAX_VELTR, c.MAX_PGTR, c.RA_PRESSURE, c.RVSP_TR "
    '    SQL$ = SQL$ & ", c.PA_VMAX, c.PA_MAX_PG, c.PA_MEAN_PG, c.PR_VMAX, c.RVOT_DIAM "
    '    SQL$ = SQL$ & ", c.TDI_EA, c.QP_QS, c.LA_VOLUMEBP, c.MPI, c.LV_MAX, c.ITEM1_VALUE, c.ITEM2_VALUE "
    '    SQL$ = SQL$ & " from cris_exam_online a, cris_patient_info b, cris_smart_save c"
    '    SQL$ = SQL$ & " where a.status<>'已刪除' and a.chartno = b.chartno and a.uni_key = c.uni_key and " & fcondition$
    'Else
    '    SQL$ = SQL$ & " from cris_exam_online a, cris_patient_info b"
    '    SQL$ = SQL$ & " where a.status<>'已刪除' and a.chartno = b.chartno and " & fcondition$
    'End If
    '/**/
    Dim SQLCount As String, RecordCount As Long
    
    SQL$ = "select a.chartno, b.name,a.age, b.sex, a.type, a.examdate, a.examtime, a.dr_report "
    If chkBasic.Value = 0 Then
        SQL$ = SQL$ & ", c.conclusion, c.impression "
        SQL$ = SQL$ & ", c.IVSD, c.LVIDD, c.LVPWD, c.IVSS, c.LVIDS, c.LVPWS, c.LVOT_DIAM, c.AO_ROOT_DIAM, c.ACS, c.LA_DIMENSION "
        SQL$ = SQL$ & ", c.EF_TEICH, c.EF_SIMPOSNS, c.LVOT_MAX, c.AV_MAX, c.AV_MAX_PG, c.AVMEANPG, c.AVA, c.AI_VMAX "
        SQL$ = SQL$ & ", c.MV_E_POINT, c.MV_A_POINT, c.MV_EA, c.MVA_P21 "
        SQL$ = SQL$ & ", c.MAX_VELTR, c.MAX_PGTR, c.RA_PRESSURE, c.RVSP_TR "
        SQL$ = SQL$ & ", c.PA_VMAX, c.PA_MAX_PG, c.PA_MEAN_PG, c.PR_VMAX, c.RVOT_DIAM "
        SQL$ = SQL$ & ", c.TDI_EA, c.QP_QS, c.LA_VOLUMEBP, c.MPI, c.LV_MAX, c.ITEM1_VALUE, c.ITEM2_VALUE "
        SQL$ = SQL$ & " from cris_exam_online a, cris_patient_info b, cris_smart_save c"
        SQL$ = SQL$ & " where a.status<>'已刪除' and a.chartno = b.chartno and a.uni_key = c.uni_key and " & fcondition$
        

        SQLCount = "select count(*) "
        SQLCount = SQLCount & " from cris_exam_online a, cris_patient_info b, cris_smart_save c"
        SQLCount = SQLCount & " where a.status<>'已刪除' and a.chartno = b.chartno and a.uni_key = c.uni_key and " & fcondition$
    Else
        SQL$ = SQL$ & " from cris_exam_online a, cris_patient_info b"
        SQL$ = SQL$ & " where a.status<>'已刪除' and a.chartno = b.chartno and " & fcondition$
        
        
        SQLCount = "select count(*) "
        SQLCount = SQLCount & " from cris_exam_online a, cris_patient_info b"
        SQLCount = SQLCount & " where a.status<>'已刪除' and a.chartno = b.chartno and " & fcondition$
    End If
    '/**/
    Screen.MousePointer = vbHourglass
'    DoEvents

    adoDB.Open dbConnection$
'    Set adoControl = adoOnline1
    
    '/**/
    'adoOnline1.Open SQL$, adoDB, adOpenForwardOnly, adLockReadOnly
    'RecordsNo& = 0
    'Do While Not adoOnline1.EOF
    '    RecordsNo& = RecordsNo& + 1
    '    adoOnline1.MoveNext
    'Loop
    '/**/
    adoOnline1.Open SQLCount, adoDB, adOpenForwardOnly, adLockReadOnly
    RecordsNo& = Val(adoOnline1(0)) + 1
    adoOnline1.Close
    
    
    adoOnline1.Open SQL$, adoDB, adOpenForwardOnly, adLockReadOnly
    '/**/
    
    
    If RecordsNo& > 0 Then
        '/**/
        'adoOnline1.MoveFirst
        '/**/
        ReDim aryRecord(fields_max - 1, RecordsNo&)
        
        j% = 0
        Do While Not adoOnline1.EOF
           For i% = 0 To adoOnline1.Fields.Count - 1
               aryRecord(i%, j%) = Replace(NoNull(adoOnline1.Fields(i%)), vbCrLf, ";")
           Next
           j% = j% + 1
           adoOnline1.MoveNext
        Loop
        adoOnline1.Close
        adoDB.Close
        Set adoOnline1 = Nothing
        Set adoDB = Nothing

        ReDim aryResult(fields_max, RecordsNo&)
        For j% = 0 To RecordsNo& - 1
            aryResult(0, j%) = j% + 1
            For i% = 0 To fields_max - 1
                aryResult(i% + 1, j%) = aryRecord(i%, j%)
            Next
        Next
'        DoEvents
'        VSFlexGrid1.Rows = RecordsNo&
'        VSFlexGrid1.Cols = fields_max + 1
    '    VSFlexGrid1.FormatString = queueCaption$
        VSFlexGrid1.BindToArray aryResult

        VSFlexGrid1.Select 1, 1
        
        Command1(0).Enabled = True
    Else
        MsgBox "查無任何符合條件的記錄!"
    End If
'    Me.ZOrder 0
    Screen.MousePointer = vbDefault
    Exit Sub
    
ttt:
    Screen.MousePointer = vbDefault
    MsgBox "錯誤的查詢條件，請重新輸入條件"
    PrintLog ("查詢錯誤" & vbCrLf & SQL$)
    List1.Clear
    
End Sub

Private Sub Command1_Click(Index As Integer)
    
    Select Case Index
        Case 0
        '轉Excel
            Call CreatePath(App.Path & "\SR\")
            Call VSFlexGrid1.SaveGrid(App.Path & "\SR\" & Format(Now, "yyyyMMddhhmmss") & ".xls", flexFileTabText)
        Case 1
        '查詢
            Call Query_SR
'            Command1(0).Enabled = True
        Case 2
        '離開
            Unload Me
        Case 3
        '加入條件
            Call ADDcondition
        Case 4
        '刪除條件
            Call Delcondition
    End Select
End Sub

Private Sub Delcondition()
    If List1.ListIndex >= 0 Then
        List1.RemoveItem (List1.ListIndex)
    End If
End Sub

Private Sub DTPicker1_CloseUp(Index As Integer)
    Text1(Index).Text = Format(DTPicker1(Index).Value, "yyyy/MM/dd")
End Sub

Private Sub Form_Load()
    Dim queueCaption$
    
    Mcombo1_text = ""
    queueCaption$ = "編號|病歷號碼　|姓名　　|年齡|性別|檢查類別　　|檢查日期　|時 間|報告醫師　　|Conclusion　　|Impression  "
    queueCaption$ = queueCaption$ & "|IVSd |LVIDd |LVPWd|IVSs |LVIDs|LVPWs|LVOT diam|Ao root diam|ACS  |LA dimension |EF(Teich)|EF(Simpson's)"
    queueCaption$ = queueCaption$ & "|LVOT max|AV max|AV maxPG|AV meanPG|AVA  |AIVmax"
    queueCaption$ = queueCaption$ & "|MV E point|MV A point|MV E/A|MVA(P1/2t)"
    queueCaption$ = queueCaption$ & "|Max vel(TR)|Max PG(TR)|RA pressure|RVSP(TR)"
    queueCaption$ = queueCaption$ & "|PA Vmax|PA maxPG|PA MeanPG|PR Vmax|RVOT Diam"
    queueCaption$ = queueCaption$ & "|TDI E/A|QP/QS|LA Volume(BP)|MPI  |LV mass|RA dimension|RV dimension"
    VSFlexGrid1.FormatString = queueCaption$
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmQueue.Enabled = True
    frmQueue.SetFocus
End Sub

Private Sub Text1_DblClick(Index As Integer)
    Load frmPhraseSpread
    frmPhraseSpread.Show 1
End Sub
