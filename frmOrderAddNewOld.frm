VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmOrderAddNewOld 
   BackColor       =   &H00808080&
   BorderStyle     =   1  '單線固定
   Caption         =   "產品價格變更"
   ClientHeight    =   5865
   ClientLeft      =   1290
   ClientTop       =   2655
   ClientWidth     =   6495
   Icon            =   "frmOrderAddNewOld.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   6495
   Begin Threed.SSPanel pnlBasic 
      Height          =   4935
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   5895
      _Version        =   65536
      _ExtentX        =   10398
      _ExtentY        =   8705
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
      BorderWidth     =   1
      BevelInner      =   1
      Outline         =   -1  'True
      Alignment       =   6
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
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Top             =   3600
         Width           =   4095
      End
      Begin VB.TextBox txtBonusMoney 
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
         Left            =   1560
         TabIndex        =   7
         Top             =   3120
         Width           =   4095
      End
      Begin VB.TextBox txtAddMoney 
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
         Left            =   1560
         TabIndex        =   6
         Top             =   2640
         Width           =   4095
      End
      Begin VB.ComboBox cmbPName 
         Height          =   300
         Left            =   1560
         TabIndex        =   1
         Top             =   600
         Width           =   4095
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
         Left            =   1560
         MaxLength       =   256
         TabIndex        =   15
         Top             =   1200
         Width           =   3855
      End
      Begin VB.TextBox txtWinningCount 
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
         Left            =   1560
         TabIndex        =   5
         Top             =   2160
         Width           =   4095
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
         Left            =   3120
         Style           =   1  '圖片外觀
         TabIndex        =   10
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&U 確定加購"
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
         Left            =   1560
         Style           =   1  '圖片外觀
         TabIndex        =   9
         Tag             =   "Edit"
         Top             =   4200
         Width           =   1335
      End
      Begin VB.TextBox txtCurrentCount 
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
         Left            =   1560
         TabIndex        =   4
         Top             =   1680
         Width           =   4095
      End
      Begin MSComCtl2.DTPicker dtpCurrentDate 
         Height          =   375
         Left            =   1560
         TabIndex        =   16
         Top             =   1200
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
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
         Format          =   103481347
         CurrentDate     =   42267
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   "交易備註"
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
         Index           =   6
         Left            =   360
         TabIndex        =   20
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   "退水金額"
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
         Index           =   5
         Left            =   360
         TabIndex        =   19
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   "漲價"
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
         Index           =   2
         Left            =   360
         TabIndex        =   18
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
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
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   14
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   "中獎數量"
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
         Left            =   360
         TabIndex        =   13
         Top             =   2160
         Width           =   1215
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
         Left            =   360
         TabIndex        =   12
         Top             =   120
         Width           =   5295
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   "交易數量"
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
         Index           =   1
         Left            =   360
         TabIndex        =   11
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
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
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   15
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3720
      Top             =   5520
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
   Begin VB.Label lblAddCount 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   "己新增0筆"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label1 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   "產品價格變更"
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
      Index           =   0
      Left            =   3120
      TabIndex        =   2
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "frmOrderAddNewOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim addCount As Integer
Dim selectFields As String

Private Sub cmdClose_Click()
    Call Form_Unload(0)
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo errout:
    Dim PID As String, LastSwiftCode As String
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim BonusTarget As String
    
    SQL = "select * from [order] order by SwiftCode desc;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, order_rec)
    If order_rec.EOF Then
        LastSwiftCode = "0"
    Else
        LastSwiftCode = order_rec("SwiftCode")
    End If
    order_rec.Close
    
    SQL = "select * from product where PName='" & cmbPName.Text & "';"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    PID = product_rec("PID")
    product_rec.Close
    
    
    Adodc1.Recordset.Fields.Item("SwiftCode").Value = Val(LastSwiftCode) + 1
    Adodc1.Recordset.Fields.Item("PID").Value = PID
    'Adodc1.Recordset.Fields.Item("PName").Value = txtPName.Text
    Adodc1.Recordset.Fields.Item("CID").Value = basVariable.SelectCID
    Adodc1.Recordset.Fields.Item("CurrentDate").Value = txtCurrentDate.Text
    Adodc1.Recordset.Fields.Item("CurrentCount").Value = txtCurrentCount.Text
    If txtWinningCount.Text = "" Then
        Adodc1.Recordset.Fields.Item("WinningCount").Value = "0"
    Else
        Adodc1.Recordset.Fields.Item("WinningCount").Value = txtWinningCount.Text
    End If
    Adodc1.Recordset.Fields.Item("AddMoney").Value = txtAddMoney.Text
    Adodc1.Recordset.Fields.Item("BonusMoney").Value = txtBonusMoney.Text
    Adodc1.Recordset.Fields.Item("Note").Value = txtNote.Text
    Call Adodc1.Recordset.Update
    Call Adodc1.Recordset.AddNew    'Call Form_Unload(0)
    
    
    cmbPName.Text = ""
    dtpCurrentDate.Value = Format(DateTime.Now, "yyyy/MM/dd")
    txtCurrentDate.Text = Format(DateTime.Now, "yyyy/MM/dd")
    txtCurrentCount.Text = ""
    txtWinningCount.Text = ""
    txtAddMoney.Text = ""
    txtBonusMoney.Text = ""
    txtNote.Text = ""
    addCount = addCount + 1
    lblAddCount.Caption = "已新增" & addCount & "筆"
    
    If False Then
errout:
        MsgBox "輸入的資料有問題，或產品名稱、交易日期、交易數量、中獎數量未填寫！"
    End If
End Sub

Private Sub dtpCurrentDate_CloseUp()
    txtCurrentDate.Text = Format(dtpCurrentDate.Value, "yyyy/MM/dd")
End Sub

'import database and export to datagrid when form load
Private Sub Form_Load()
    cmbPName.Enabled = True
    txtCurrentDate.Enabled = True
    dtpCurrentDate.Enabled = True
    

    lblName(0).Caption = basVariable.SelectCName
    selectFields = "SwiftCode,CID,[order].PID,PName,CurrentDate,CurrentCount,WinningCount"
    
    Adodc1.ConnectionString = basDataBase.Connection_String
    Adodc1.CommandType = adCmdText
    'Adodc1.RecordSource = "select " & selectFields & " from [order] where [order].CID='" & basVariable.SelectCID & "';"
    Adodc1.RecordSource = "select * from [order] where [order].CID='" & basVariable.SelectCID & "';"
    Adodc1.LockType = adLockOptimistic
    
    
    'Set txtPName.DataSource = Adodc1
    Set txtCurrentDate.DataSource = Adodc1
    Set txtCurrentCount.DataSource = Adodc1
    Set txtWinningCount.DataSource = Adodc1
    Set txtAddMoney.DataSource = Adodc1


    'add new
    
    Call Adodc1.Recordset.AddNew
    
    
    Call ComboBox_LoadFrom_DataBase_ByFile(cmbPName, "PName", "product", "", "", "")
    
    dtpCurrentDate.Value = Format(DateTime.Now, "yyyy/MM/dd")
    txtCurrentDate.Text = Format(DateTime.Now, "yyyy/MM/dd")
    
    
    addCount = 0
    lblAddCount.Caption = "已新增" & addCount & "筆"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmOrder.Show
    Unload Me
End Sub


Private Sub cmbPName_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtCurrentDate.SetFocus
    End If
End Sub

Private Sub txtCurrentDate_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtCurrentCount.SetFocus
    End If
End Sub

Private Sub txtCurrentCount_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtWinningCount.SetFocus
    End If
End Sub

Private Sub txtWinningCount_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtAddMoney.SetFocus
    End If
End Sub

Private Sub txtAddMoney_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtBonusMoney.SetFocus
    End If
End Sub

Private Sub txtBonusMoney_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtNote.SetFocus
    End If
End Sub

Private Sub txtNote_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call cmdUpdate_Click
    End If
End Sub

