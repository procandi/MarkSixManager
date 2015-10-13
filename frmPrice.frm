VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPrice 
   Caption         =   "客戶產品價格表"
   ClientHeight    =   10560
   ClientLeft      =   7425
   ClientTop       =   3285
   ClientWidth     =   14895
   Icon            =   "frmPrice.frx":0000
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
      Begin VB.TextBox txtUpset 
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
         Left            =   10440
         TabIndex        =   19
         Top             =   120
         Width           =   1935
      End
      Begin VB.TextBox txtCurrentDate 
         Enabled         =   0   'False
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
         Left            =   7320
         MaxLength       =   10
         TabIndex        =   16
         Top             =   120
         Width           =   1650
      End
      Begin VB.TextBox txtWinningPrice 
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
         Left            =   7320
         TabIndex        =   15
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtPName 
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
         TabIndex        =   10
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtCurrentPrice 
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
         TabIndex        =   9
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
         TabIndex        =   5
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtAccessionNo 
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
         TabIndex        =   4
         Top             =   1680
         Width           =   1815
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
      Begin VB.CommandButton cmdModifyPrice 
         BackColor       =   &H00FFC0C0&
         Caption         =   "修改客戶產品價格"
         Enabled         =   0   'False
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
         Left            =   9720
         Style           =   1  '圖片外觀
         TabIndex        =   1
         Top             =   600
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker dtpCurrentDate 
         Height          =   360
         Left            =   7320
         TabIndex        =   17
         Top             =   120
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy/MM/dd"
         Format          =   3866627
         CurrentDate     =   37058
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "底價"
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
         Index           =   2
         Left            =   9360
         TabIndex        =   20
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "交易日期"
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
         Index           =   1
         Left            =   6240
         TabIndex        =   18
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "購買價格"
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
         Index           =   0
         Left            =   3120
         TabIndex        =   14
         Top             =   600
         Width           =   1095
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
         TabIndex        =   11
         Top             =   120
         Width           =   5295
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "中獎金額"
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
         Left            =   6240
         TabIndex        =   8
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "檢查編號"
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
         Index           =   19
         Left            =   120
         TabIndex        =   7
         Top             =   1680
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
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   9255
      Left            =   0
      TabIndex        =   12
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
         TabIndex        =   13
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
Attribute VB_Name = "frmPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim selectFields As String

Private Sub cmdClear_Click()
    txtPName.Text = ""
    txtCurrentDate.Text = ""
    txtCurrentPrice.Text = ""
    txtWinningPrice.Text = ""
    txtUpset.Text = ""
End Sub

'add function to refresh database and datagrid
Private Sub cmdRefresh_Click()
    Dim condition As String
    
    condition = ""

    If txtPName.Text <> "" Then
        condition = condition & IIf(condition = "", "", "and ") & "PName='" & txtPName.Text & "' "
    End If
    If txtCurrentDate.Text <> "" Then
        condition = condition & IIf(condition = "", "", "and ") & "CurrentDate='" & txtCurrentDate.Text & "' "
    End If
    If txtCurrentPrice.Text <> "" Then
        condition = condition & IIf(condition = "", "", "and ") & "CurrentPrice='" & txtCurrentPrice.Text & "' "
    End If
    If txtWinningPrice.Text <> "" Then
        condition = condition & IIf(condition = "", "", "and ") & "WinningPrice='" & txtWinningPrice.Text & "' "
    End If
    If txtUpset.Text <> "" Then
        condition = condition & IIf(condition = "", "", "and ") & "Upset='" & txtUpset.Text & "' "
    End If

    If condition = "" Then
        Adodc1.RecordSource = "select " & selectFields & " from price,product where price.PID=product.PID and CID='" & basVariable.SelectCID & "' order by price.PID,CurrentDate desc;"
    Else
        Adodc1.RecordSource = "select " & selectFields & " from price,product where price.PID=product.PID and CID='" & basVariable.SelectCID & "' and " & condition & " order by price.PID,CurrentDate desc;"
    End If
    Adodc1.Refresh
    RefreshDataGridHeader
End Sub

Private Sub dtpCurrentDate_CloseUp()
    txtCurrentDate.Text = Format(dtpCurrentDate.Value, "yyyy/MM/dd")
End Sub

'do refresh database and datagrid when form paint
Private Sub Form_Paint()
    Call cmdRefresh_Click
End Sub

Private Sub cmdModifyPrice_Click()
    basVariable.Action = "ModifyPrice"
    frmPriceUpdate.Show
    Me.Hide
End Sub

Private Sub cmdClose_Click()
    Call Form_Unload(0)
End Sub

'get something system needed when user click datagrid row
Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Adodc1.Recordset.RecordCount > 0 Then
        If DataGrid1.Columns("交易流水號") <> "" Then
            cmdModifyPrice.Enabled = True
            basVariable.Parameter = DataGrid1.Row
            basVariable.CurrentSwiftCode = DataGrid1.Columns("交易流水號")
            basVariable.SelectPID = DataGrid1.Columns("產品編號")
        End If
        
        If DataGrid1.SelBookmarks.Count <> 0 Then Call DataGrid1.SelBookmarks.Remove(0)
        Call DataGrid1.SelBookmarks.Add(DataGrid1.Bookmark)
    End If
End Sub

'import database and export to datagrid when form load
Private Sub Form_Load()
    DataGrid1.AllowAddNew = False
    DataGrid1.AllowUpdate = False
    
    lblName(0).Caption = basVariable.SelectCName
    selectFields = "SwiftCode,CID,price.PID,PName,CurrentDate,CurrentPrice,WinningPrice,Upset"

    Adodc1.ConnectionString = basDataBase.Connection_String
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select " & selectFields & " from price,product where price.PID=product.PID and CID='" & basVariable.SelectCID & "' order by price.PID,CurrentDate desc;"
    Set DataGrid1.DataSource = Adodc1
    RefreshDataGridHeader
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmCustom.Show
    Unload Me
End Sub

'a function to batch rename datagrid header
Sub RefreshDataGridHeader()
    DataGrid1.Columns("SwiftCode").Caption = "交易流水號"
    DataGrid1.Columns("CID").Caption = "客戶編號"
    DataGrid1.Columns("PID").Caption = "產品編號"
    DataGrid1.Columns("CurrentDate").Caption = "交易日期"
    DataGrid1.Columns("CurrentPrice").Caption = "購買價格"
    DataGrid1.Columns("WinningPrice").Caption = "中獎金額"
    DataGrid1.Columns("Upset").Caption = "價格底線"
    
    DataGrid1.Columns("PName").Caption = "產品名稱"
End Sub

