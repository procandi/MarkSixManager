VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCustomAddNew 
   BackColor       =   &H00808080&
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶資料明細"
   ClientHeight    =   6270
   ClientLeft      =   10800
   ClientTop       =   4425
   ClientWidth     =   8775
   Icon            =   "frmCustomAddNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   8775
   Begin Threed.SSPanel pnlRegist 
      Height          =   5535
      Left            =   120
      TabIndex        =   18
      Top             =   600
      Width           =   8415
      _Version        =   65536
      _ExtentX        =   14843
      _ExtentY        =   9763
      _StockProps     =   15
      ForeColor       =   0
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Outline         =   -1  'True
      FloodColor      =   0
      Alignment       =   6
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
         Left            =   1320
         MaxLength       =   256
         TabIndex        =   0
         Top             =   240
         Width           =   3015
      End
      Begin VB.ComboBox cmbProportion 
         Height          =   300
         Left            =   1320
         TabIndex        =   5
         Top             =   1920
         Width           =   3015
      End
      Begin VB.TextBox txtOpenDate 
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
         Left            =   2760
         MaxLength       =   256
         TabIndex        =   30
         Top             =   1200
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
         Left            =   1320
         MaxLength       =   256
         TabIndex        =   2
         Top             =   720
         Width           =   3015
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
         Height          =   795
         Left            =   240
         MaxLength       =   256
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   4080
         Width           =   7935
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
         Height          =   435
         Left            =   240
         MaxLength       =   256
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   3240
         Width           =   7935
      End
      Begin VB.TextBox txtPhone6 
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
         Left            =   5520
         MaxLength       =   256
         TabIndex        =   12
         Top             =   2520
         Width           =   2655
      End
      Begin VB.TextBox txtPhone5 
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
         Left            =   5520
         MaxLength       =   256
         TabIndex        =   11
         Top             =   2160
         Width           =   2655
      End
      Begin VB.TextBox txtPhone4 
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
         Left            =   5520
         MaxLength       =   256
         TabIndex        =   10
         Top             =   1800
         Width           =   2655
      End
      Begin VB.TextBox txtPhone3 
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
         Left            =   5520
         MaxLength       =   256
         TabIndex        =   9
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtPhone2 
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
         Left            =   5520
         MaxLength       =   256
         TabIndex        =   8
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox txtPhone1 
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
         Left            =   5520
         MaxLength       =   256
         TabIndex        =   7
         Top             =   720
         Width           =   2655
      End
      Begin VB.ComboBox cmbBonusTarget 
         Height          =   300
         Left            =   1320
         TabIndex        =   6
         Top             =   2280
         Width           =   3015
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
         Left            =   1320
         MaxLength       =   256
         TabIndex        =   4
         Top             =   1560
         Width           =   3015
      End
      Begin VB.TextBox txtCID 
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
         Left            =   5520
         MaxLength       =   256
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&U 確定修改"
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
         Height          =   495
         Left            =   2880
         Style           =   1  '圖片外觀
         TabIndex        =   16
         Tag             =   "Edit"
         Top             =   4920
         Width           =   2535
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&X 關閉"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5520
         Style           =   1  '圖片外觀
         TabIndex        =   17
         Top             =   4920
         Width           =   2655
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&R 確定新增"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         Style           =   1  '圖片外觀
         TabIndex        =   15
         Tag             =   "Insert"
         Top             =   4920
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker dtpOpenDate 
         Height          =   375
         Left            =   2760
         TabIndex        =   3
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
         Format          =   88670211
         CurrentDate     =   42267
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00808080&
         BorderStyle     =   1  '單線固定
         Caption         =   "客戶編號"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   0
         Left            =   4440
         TabIndex        =   29
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00808080&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   10
         Left            =   240
         TabIndex        =   27
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00808080&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   2
         Left            =   4440
         TabIndex        =   19
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00808080&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblRegist 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00808080&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   25
         Top             =   1200
         Width           =   2505
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00808080&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Index           =   8
         Left            =   240
         TabIndex        =   24
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00808080&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   6
         Left            =   240
         TabIndex        =   23
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00808080&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   5
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00808080&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   1
         Left            =   240
         TabIndex        =   21
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00808080&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   4
         Left            =   240
         TabIndex        =   20
         Top             =   2280
         Width           =   1095
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   240
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
   Begin VB.Label Label1 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   "客戶資料明細"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   24
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   1
      Left            =   5280
      TabIndex        =   28
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmCustomAddNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Call Adodc1.Recordset.Cancel
    Call Form_Unload(0)
End Sub

Private Sub cmdOK_Click()
    Adodc1.Recordset.Fields.Item("CID").Value = txtCID.Text
    Adodc1.Recordset.Fields.Item("CName").Value = txtCName.Text
    Call Adodc1.Recordset.Update
    Call Form_Unload(0)
End Sub

Private Sub cmdUpdate_Click()
    Call Adodc1.Recordset.Update
    Call Form_Unload(0)
End Sub

Private Sub dtpOpenDate_CloseUp()
    txtOpenDate.Text = Format(dtpOpenDate.Value, "yyyy/MM/dd")
End Sub

'import database and export to datagrid when form load
Private Sub Form_Load()
    Adodc1.ConnectionString = basDataBase.Connection_String
    Adodc1.CommandType = adCmdText
    If basVariable.Action = "AddNewCustom" Then
        Adodc1.RecordSource = "select * from custom;"
    Else
        Adodc1.RecordSource = "select * from custom where CID='" & basVariable.SelectCID & "';"
    End If
    Adodc1.LockType = adLockOptimistic
    
    
    Set txtCType.DataSource = Adodc1
    Set txtCName.DataSource = Adodc1
    Set dtpOpenDate.DataSource = Adodc1
    Set txtBankID.DataSource = Adodc1
    Set cmbProportion.DataSource = Adodc1
    Set cmbBonusTarget.DataSource = Adodc1
    Set txtCID.DataSource = Adodc1
    Set txtPhone1.DataSource = Adodc1
    Set txtPhone2.DataSource = Adodc1
    Set txtPhone3.DataSource = Adodc1
    Set txtPhone4.DataSource = Adodc1
    Set txtPhone5.DataSource = Adodc1
    Set txtPhone6.DataSource = Adodc1
    Set txtAddress.DataSource = Adodc1
    Set txtNote.DataSource = Adodc1
    

    
    
    
    If basVariable.Action = "AddNewCustom" Then
        'add new
        
        Call Adodc1.Recordset.MoveLast
        CID = Adodc1.Recordset.Fields.Item("CID").Value
        
        Call Adodc1.Recordset.AddNew
        txtCID.Text = Val(CID) + 1

        cmdOK.Enabled = True
        cmdUpdate.Enabled = False
    Else
        'modify
        
        txtCType.DataField = "CType"
        txtCName.DataField = "CName"
        txtOpenDate.DataField = "OpenDate"
        txtBankID.DataField = "BankID"
        cmbProportion.DataField = "Proportion"
        cmbBonusTarget.DataField = "BonusTarget"
        txtCID.DataField = "CID"
        txtPhone1.DataField = "Phone1"
        txtPhone2.DataField = "Phone2"
        txtPhone3.DataField = "Phone3"
        txtPhone4.DataField = "Phone4"
        txtPhone5.DataField = "Phone5"
        txtPhone6.DataField = "Phone6"
        txtAddress.DataField = "Address"
        txtNote.DataField = "Note"
        txtCID.Text = basVariable.SelectCID
    
        cmdOK.Enabled = False
        cmdUpdate.Enabled = True
    End If
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmCustom.Show
    Unload Me
End Sub
