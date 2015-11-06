VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCustom 
   Caption         =   "客戶資料表"
   ClientHeight    =   10665
   ClientLeft      =   1965
   ClientTop       =   3240
   ClientWidth     =   14985
   Icon            =   "frmCustom.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   10665
   ScaleWidth      =   14985
   Begin VB.Timer Timer1 
      Left            =   3840
      Top             =   1560
   End
   Begin Threed.SSPanel pnlFilter 
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      _Version        =   65536
      _ExtentX        =   26485
      _ExtentY        =   2778
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
      Begin VB.TextBox txtBonusTarget 
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
         Left            =   13320
         TabIndex        =   27
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtProportion 
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
         Left            =   10320
         TabIndex        =   26
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtCType 
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
         TabIndex        =   25
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton cmdOrder 
         BackColor       =   &H00FFC0C0&
         Caption         =   "客戶購買明細表"
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
         Left            =   7320
         Style           =   1  '圖片外觀
         TabIndex        =   24
         Top             =   120
         Width           =   2295
      End
      Begin VB.TextBox txtNote 
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
         Left            =   10320
         TabIndex        =   23
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtAddress 
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
         TabIndex        =   21
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton cmdPrice 
         BackColor       =   &H00FFC0C0&
         Caption         =   "客戶產品價格表"
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
         TabIndex        =   20
         Top             =   120
         Width           =   2295
      End
      Begin VB.CommandButton cmdModify 
         BackColor       =   &H00FFC0C0&
         Caption         =   "編輯客戶"
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
         Left            =   120
         Style           =   1  '圖片外觀
         TabIndex        =   11
         Top             =   120
         Width           =   2295
      End
      Begin VB.TextBox txtBankID 
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
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CommandButton cmdAppend 
         BackColor       =   &H00FFC0C0&
         Caption         =   "新增客戶"
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
         Left            =   2520
         Style           =   1  '圖片外觀
         TabIndex        =   8
         Top             =   120
         Width           =   2295
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
         Left            =   13560
         Style           =   1  '圖片外觀
         TabIndex        =   7
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox txtCName 
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
         TabIndex        =   6
         Top             =   600
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
         TabIndex        =   5
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFC0C0&
         Caption         =   "刪除客戶"
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
         Left            =   4920
         Style           =   1  '圖片外觀
         TabIndex        =   4
         Top             =   120
         Width           =   2295
      End
      Begin VB.TextBox txtPhone 
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
         TabIndex        =   3
         Top             =   1080
         Width           =   1815
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
         TabIndex        =   2
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtOpenDate 
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
         Left            =   4200
         MaxLength       =   10
         TabIndex        =   1
         Top             =   600
         Width           =   1650
      End
      Begin MSComCtl2.DTPicker dtpOpenDate 
         Height          =   360
         Left            =   4200
         TabIndex        =   10
         Top             =   600
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
         Format          =   103612419
         CurrentDate     =   37058
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "備註"
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
         Left            =   9240
         TabIndex        =   22
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "客別註記"
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
         Left            =   6240
         TabIndex        =   19
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "成數"
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
         Left            =   9240
         TabIndex        =   18
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "銀行帳號"
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
         Index           =   9
         Left            =   3120
         TabIndex        =   17
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "地址"
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
         Index           =   18
         Left            =   6240
         TabIndex        =   16
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "姓名"
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
         Index           =   20
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "電話"
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
         Index           =   21
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "退水"
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
         Index           =   22
         Left            =   12240
         TabIndex        =   13
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "開戶日期"
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
         TabIndex        =   12
         Top             =   600
         Width           =   1095
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   8895
      Left            =   0
      TabIndex        =   28
      Top             =   1800
      Width           =   15015
      _Version        =   65536
      _ExtentX        =   26485
      _ExtentY        =   15690
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
         Height          =   8655
         Left            =   120
         TabIndex        =   29
         Top             =   120
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   15266
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
      Top             =   1560
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
Attribute VB_Name = "frmCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAppend_Click()
    basVariable.Action = "AddNewCustom"
    frmCustomAddNew.Show
    Me.Hide
End Sub

Private Sub cmdOrder_Click()
    basVariable.Action = "OrderDetail"
    frmOrder.Show
    Me.Hide
End Sub

Private Sub cmdClear_Click()
    txtCType.Text = ""
    txtOpenDate.Text = ""
    txtAddress.Text = ""
    txtProportion.Text = ""
    txtBonusTarget.Text = ""
    txtPhone.Text = ""
    txtBankID.Text = ""
    txtCName.Text = ""
    txtNote.Text = ""
End Sub

Private Sub cmdDelete_Click()
    Call Adodc1.Recordset.Delete
    Call Adodc1.Recordset.Update
End Sub

Private Sub cmdModify_Click()
    basVariable.Action = "ModifyCustom"
    frmCustomAddNew.Show
    Me.Hide
End Sub

Private Sub cmdPrice_Click()
    basVariable.Action = "PriceDetail"
    frmPrice.Show
    Me.Hide
End Sub

'add function to refresh database and datagrid
Private Sub cmdRefresh_Click()
    cmdModify.Enabled = False
    cmdDelete.Enabled = False
    cmdOrder.Enabled = False
    cmdPrice.Enabled = False
    

    Dim condition As String
    
    condition = ""

    If txtCName.Text <> "" Then
        condition = condition & IIf(condition = "", "", "and ") & "CName='" & txtCName.Text & "' "
    End If
    If txtOpenDate.Text <> "" Then
        condition = condition & IIf(condition = "", "", "and ") & "OpenDate='" & txtOpenDate.Text & "' "
    End If
    If txtAddress.Text <> "" Then
        condition = condition & IIf(condition = "", "", "and ") & "Address='" & txtAddress.Text & "' "
    End If
    If txtProportion.Text <> "" Then
        condition = condition & IIf(condition = "", "", "and ") & "Proportion='" & txtProportion.Text & "' "
    End If
    If txtBonusTarget.Text <> "" Then
        condition = condition & IIf(condition = "", "", "and ") & "BonusTarget='" & txtBonusTarget.Text & "' "
    End If
    If txtPhone.Text <> "" Then
        condition = condition & IIf(condition = "", "", "and ")
        
        condition = condition & "(Phone1='" & txtPhone.Text & "' or "
        condition = condition & "Phone2='" & txtPhone.Text & "' or "
        condition = condition & "Phone3='" & txtPhone.Text & "' or "
        condition = condition & "Phone4='" & txtPhone.Text & "' or "
        condition = condition & "Phone5='" & txtPhone.Text & "' or "
        condition = condition & "Phone6='" & txtPhone.Text & "') "
    End If
    If txtBankID.Text <> "" Then
        condition = condition & IIf(condition = "", "", "and ") & "BankID='" & txtBankID.Text & "' "
    End If
    If txtCType.Text <> "" Then
        condition = condition & IIf(condition = "", "", "and ") & "CType='" & txtCType.Text & "' "
    End If
    If txtNote.Text <> "" Then
        condition = condition & IIf(condition = "", "", "and ") & "Note='" & txtNote.Text & "' "
    End If
    
    
    If condition = "" Then
        Adodc1.RecordSource = "select * from custom;"
    Else
        Adodc1.RecordSource = "select * from custom where " & condition & ";"
    End If
    Adodc1.Refresh
    RefreshDataGridHeader
End Sub

'get something system needed when user click datagrid row
Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Adodc1.Recordset.RecordCount > 0 Then
        cmdModify.Enabled = True
        cmdDelete.Enabled = True
        cmdOrder.Enabled = True
        cmdPrice.Enabled = True
    
        basVariable.SelectCID = DataGrid1.Columns("客戶編號")
        basVariable.SelectCName = DataGrid1.Columns("姓名")
        If DataGrid1.SelBookmarks.Count <> 0 Then Call DataGrid1.SelBookmarks.Remove(0)
        Call DataGrid1.SelBookmarks.Add(DataGrid1.Bookmark)
    End If
End Sub

Private Sub dtpOpenDate_CloseUp()
    txtOpenDate.Text = Format(dtpOpenDate.Value, "yyyy/MM/dd")
End Sub

'import database and export to datagrid when form load
Private Sub Form_Load()
    DataGrid1.AllowAddNew = False
    DataGrid1.AllowUpdate = False
    
    Adodc1.ConnectionString = basDataBase.Connection_String
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from custom;"
    Set DataGrid1.DataSource = Adodc1
    
    
    dtpOpenDate.Value = Format(DateTime.Now, "yyyy/MM/dd")
End Sub

Private Sub cmdClose_Click()
    Call Form_Unload(0)
End Sub

'do refresh database and datagrid when form paint
Private Sub Form_Paint()
    Call cmdRefresh_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmProve.Show
    Unload Me
End Sub

'a function to batch rename datagrid header
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
