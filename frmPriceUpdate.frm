VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPriceUpdate 
   BackColor       =   &H00808080&
   BorderStyle     =   1  '單線固定
   Caption         =   "產品價格變更"
   ClientHeight    =   4755
   ClientLeft      =   6135
   ClientTop       =   5940
   ClientWidth     =   6495
   Icon            =   "frmPriceUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   6495
   Begin Threed.SSPanel pnlBasic 
      Height          =   3855
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   5895
      _Version        =   65536
      _ExtentX        =   10398
      _ExtentY        =   6800
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
      Begin VB.TextBox txtCurrentDate 
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
         TabIndex        =   16
         Top             =   1200
         Width           =   3855
      End
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
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   2640
         Width           =   4095
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00FFC0C0&
         Caption         =   "下一筆"
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
         Left            =   4320
         Style           =   1  '圖片外觀
         TabIndex        =   10
         Tag             =   "Edit"
         Top             =   3120
         Width           =   1335
      End
      Begin VB.CommandButton cmdPrev 
         BackColor       =   &H00FFC0C0&
         Caption         =   "上一筆"
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
         Left            =   120
         Style           =   1  '圖片外觀
         TabIndex        =   9
         Tag             =   "Edit"
         Top             =   3120
         Width           =   1335
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
         Left            =   3000
         Style           =   1  '圖片外觀
         TabIndex        =   8
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&U 確定修改"
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
         TabIndex        =   7
         Tag             =   "Edit"
         Top             =   3120
         Width           =   1335
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
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   1680
         Width           =   4095
      End
      Begin VB.TextBox txtPName 
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
         Left            =   1560
         TabIndex        =   2
         Top             =   600
         Width           =   4095
      End
      Begin MSComCtl2.DTPicker dtpCurrentDate 
         Height          =   375
         Left            =   1560
         TabIndex        =   17
         Top             =   1200
         Width           =   4095
         _ExtentX        =   7223
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
         Format          =   88604675
         CurrentDate     =   42267
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
         TabIndex        =   15
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   "價格底線"
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
         TabIndex        =   14
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
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
      Left            =   240
      Top             =   360
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
      TabIndex        =   1
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "frmPriceUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim selectFields As String

Private Sub cmdClose_Click()
    Call Form_Unload(0)
End Sub

Private Sub cmdNext_Click()
    If Val(basVariable.Parameter) < Adodc1.Recordset.RecordCount Then
        basVariable.Parameter = Val(basVariable.Parameter) + 1
        Call Adodc1.Recordset.MoveNext
     End If
    
    
    If Val(basVariable.Parameter) > 0 Then
        cmdPrev.Enabled = True
    Else
        cmdPrev.Enabled = False
    End If
    If Val(basVariable.Parameter) < Adodc1.Recordset.RecordCount - 1 Then
        cmdNext.Enabled = True
    Else
        cmdNext.Enabled = False
    End If
End Sub

Private Sub cmdPrev_Click()
    If Val(basVariable.Parameter) > 0 Then
        basVariable.Parameter = Val(basVariable.Parameter) - 1
        Call Adodc1.Recordset.MovePrevious
    End If
    
    
    If Val(basVariable.Parameter) > 0 Then
        cmdPrev.Enabled = True
    Else
        cmdPrev.Enabled = False
    End If
    If Val(basVariable.Parameter) < Adodc1.Recordset.RecordCount - 1 Then
        cmdNext.Enabled = True
    Else
        cmdNext.Enabled = False
    End If
End Sub

Private Sub cmdUpdate_Click()
    Call Adodc1.Recordset.Update
    Call Form_Unload(0)
End Sub

Private Sub dtpCurrentDate_CloseUp()
    txtCurrentDate.Text = Format(dtpCurrentDate.Value, "yyyy/MM/dd")
End Sub

'import database and export to datagrid when form load
Private Sub Form_Load()
    txtCurrentDate.Enabled = False
    dtpCurrentDate.Enabled = False
    

    lblName(0).Caption = basVariable.SelectCName
    selectFields = "SwiftCode,CID,price.PID,PName,CurrentDate,CurrentPrice,WinningPrice,Upset"
    
    Adodc1.ConnectionString = basDataBase.Connection_String
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from price,product where price.PID=product.PID and CID='" & basVariable.SelectCID & "' order by price.PID,CurrentDate desc;"
    Adodc1.LockType = adLockOptimistic
    
    
    Set txtPName.DataSource = Adodc1
    Set txtCurrentDate.DataSource = Adodc1
    Set txtCurrentPrice.DataSource = Adodc1
    Set txtWinningPrice.DataSource = Adodc1
    Set txtUpset.DataSource = Adodc1
 

    'modify
    
    txtPName.DataField = "PName"
    txtCurrentDate.DataField = "CurrentDate"
    txtCurrentPrice.DataField = "CurrentPrice"
    txtWinningPrice.DataField = "WinningPrice"
    txtUpset.DataField = "Upset"
    
    
    Call Adodc1.Recordset.Move(basVariable.Parameter)
    If Val(basVariable.Parameter) > 0 Then
        cmdPrev.Enabled = True
    Else
        cmdPrev.Enabled = False
    End If
    If Val(basVariable.Parameter) < Adodc1.Recordset.RecordCount - 1 Then
        cmdNext.Enabled = True
    Else
        cmdNext.Enabled = False
    End If
    
    
    txtCurrentDate.Text = Format(DateTime.Now, "yyyy/MM/dd")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmPrice.Show
    Unload Me
End Sub
