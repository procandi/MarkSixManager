VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBought 
   Caption         =   "產品購買明細"
   ClientHeight    =   10560
   ClientLeft      =   7425
   ClientTop       =   3090
   ClientWidth     =   14895
   Icon            =   "frmBought.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   10560
   ScaleWidth      =   14895
   Begin Threed.SSPanel pnlFilter 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      _Version        =   65536
      _ExtentX        =   26485
      _ExtentY        =   1931
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      Outline         =   -1  'True
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         TabIndex        =   8
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtPrice 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4200
         TabIndex        =   7
         Top             =   600
         Width           =   1935
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00FFC0C0&
         Caption         =   "更新清單"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13560
         Style           =   1  '圖片外觀
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00FFC0C0&
         Caption         =   "清除條件"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12240
         Style           =   1  '圖片外觀
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&X 關　　閉"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12960
         Style           =   1  '圖片外觀
         TabIndex        =   2
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmdModify 
         BackColor       =   &H00FFC0C0&
         Caption         =   "加購"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         Style           =   1  '圖片外觀
         TabIndex        =   1
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label lblName 
         Alignment       =   2  '置中對齊
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   "王小明"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   5295
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "數量"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   23
         Left            =   3120
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "產品名稱"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   9255
      Left            =   0
      TabIndex        =   10
      Top             =   1320
      Width           =   15015
      _Version        =   65536
      _ExtentX        =   26485
      _ExtentY        =   16325
      _StockProps     =   15
      BackColor       =   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      Outline         =   -1  'True
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   9015
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   15901
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   17
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmBought"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRefresh_Click()
    'RefreshDataGridHeader
End Sub

Private Sub Form_Paint()
    Call cmdRefresh_Click
End Sub

Private Sub cmdClose_Click()
    Call Form_Unload(0)
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Adodc1.Recordset.RecordCount > 0 Then
        If DataGrid1.SelBookmarks.Count <> 0 Then Call DataGrid1.SelBookmarks.Remove(0)
        Call DataGrid1.SelBookmarks.Add(DataGrid1.Bookmark)
    End If
End Sub

Private Sub Form_Load()
    Adodc1.ConnectionString = basDataBase.Connection_String
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from bought;"
    Set DataGrid1.DataSource = Adodc1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmCustom.Show
    Unload Me
End Sub

Sub RefreshDataGridHeader()
    DataGrid1.Columns("CID").Caption = "客戶編號"
    DataGrid1.Columns("CName").Caption = "姓名"
    DataGrid1.Columns("CType").Caption = "客別註記"
    DataGrid1.Columns("Address").Caption = "地址"
    DataGrid1.Columns("OpenDate").Caption = "開戶日期"
    DataGrid1.Columns("BankID").Caption = "銀行帳號"
    DataGrid1.Columns("Proportion").Caption = "成數"
    DataGrid1.Columns("BonusTarget").Caption = "退水"
    DataGrid1.Columns("Phone1").Caption = "電話1"
    DataGrid1.Columns("Phone2").Caption = "電話2"
    DataGrid1.Columns("Phone3").Caption = "電話3"
    DataGrid1.Columns("Phone4").Caption = "電話4"
    DataGrid1.Columns("Phone5").Caption = "電話5"
    DataGrid1.Columns("Phone6").Caption = "電話6"
    DataGrid1.Columns("Note").Caption = "備註"
End Sub

