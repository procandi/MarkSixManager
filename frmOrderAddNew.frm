VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmOrderAddNew 
   BackColor       =   &H00808080&
   BorderStyle     =   1  '單線固定
   Caption         =   "產品價格變更"
   ClientHeight    =   10215
   ClientLeft      =   615
   ClientTop       =   840
   ClientWidth     =   5010
   Icon            =   "frmOrderAddNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10215
   ScaleWidth      =   5010
   Begin Threed.SSPanel pnlBasic 
      Height          =   9375
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   4575
      _Version        =   65536
      _ExtentX        =   8070
      _ExtentY        =   16536
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
      Begin VB.ComboBox cmbPName 
         Height          =   300
         Left            =   2640
         TabIndex        =   36
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtWinningCount_Special 
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
         Left            =   2640
         TabIndex        =   12
         Top             =   6000
         Width           =   1455
      End
      Begin VB.TextBox txtCurrentCount_Special 
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
         Left            =   2640
         TabIndex        =   11
         Top             =   5520
         Width           =   1455
      End
      Begin VB.TextBox txtWinningCount_4K 
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
         Left            =   2640
         TabIndex        =   10
         Top             =   5040
         Width           =   1455
      End
      Begin VB.TextBox txtCurrentCount_4K 
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
         Left            =   2640
         TabIndex        =   9
         Top             =   4560
         Width           =   1455
      End
      Begin VB.TextBox txtWinningCount_3K 
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
         Left            =   2640
         TabIndex        =   8
         Top             =   4080
         Width           =   1455
      End
      Begin VB.TextBox txtCurrentCount_3K 
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
         Left            =   2640
         TabIndex        =   7
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox txtWinningCount_2K 
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
         Left            =   2640
         TabIndex        =   6
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox txtCurrentCount_2K 
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
         Left            =   2640
         TabIndex        =   5
         Top             =   2640
         Width           =   1455
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
         Height          =   855
         Left            =   2640
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   7440
         Width           =   1455
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
         Left            =   2640
         TabIndex        =   14
         Top             =   6960
         Width           =   1455
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
         Left            =   2640
         TabIndex        =   13
         Top             =   6480
         Width           =   1455
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
         Left            =   2640
         MaxLength       =   256
         TabIndex        =   22
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtWinningCount_Car 
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
         Left            =   2640
         TabIndex        =   4
         Top             =   2160
         Width           =   1455
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
         Left            =   2880
         Style           =   1  '圖片外觀
         TabIndex        =   17
         Top             =   8520
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
         Left            =   720
         Style           =   1  '圖片外觀
         TabIndex        =   16
         Tag             =   "Edit"
         Top             =   8520
         Width           =   1335
      End
      Begin VB.TextBox txtCurrentCount_Car 
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
         Left            =   2640
         TabIndex        =   3
         Top             =   1680
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpCurrentDate 
         Height          =   375
         Left            =   2640
         TabIndex        =   23
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   96141315
         CurrentDate     =   42267
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   "3包或特交易數量"
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
         Index           =   18
         Left            =   600
         TabIndex        =   35
         Top             =   5520
         Width           =   2055
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   "3包或特中獎數量"
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
         Index           =   17
         Left            =   600
         TabIndex        =   34
         Top             =   6000
         Width           =   2055
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   "4K交易數量"
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
         Index           =   16
         Left            =   600
         TabIndex        =   33
         Top             =   4560
         Width           =   2055
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   "4K中獎數量"
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
         Index           =   14
         Left            =   600
         TabIndex        =   32
         Top             =   5040
         Width           =   2055
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   "3K交易數量"
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
         Index           =   13
         Left            =   600
         TabIndex        =   31
         Top             =   3600
         Width           =   2055
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   "3K中獎數量"
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
         Index           =   12
         Left            =   600
         TabIndex        =   30
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   "2K交易數量"
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
         Index           =   11
         Left            =   600
         TabIndex        =   29
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   "2K中獎數量"
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
         Index           =   10
         Left            =   600
         TabIndex        =   28
         Top             =   3120
         Width           =   2055
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
         Left            =   600
         TabIndex        =   27
         Top             =   7440
         Width           =   2055
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
         Left            =   600
         TabIndex        =   26
         Top             =   6960
         Width           =   2055
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
         Left            =   600
         TabIndex        =   25
         Top             =   6480
         Width           =   2055
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
         Left            =   600
         TabIndex        =   21
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   "車中獎數量"
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
         Left            =   600
         TabIndex        =   20
         Top             =   2160
         Width           =   2055
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
         Left            =   600
         TabIndex        =   19
         Top             =   120
         Width           =   3495
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   "車交易數量"
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
         Left            =   600
         TabIndex        =   18
         Top             =   1680
         Width           =   2055
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
         Left            =   600
         TabIndex        =   2
         Top             =   1080
         Width           =   2055
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   6960
      Top             =   9840
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
      TabIndex        =   24
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
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "frmOrderAddNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim addCount As Integer
Dim selectFields As String
Private Enum PArray
    P539Car = 0
    P5392K = 1
    P5393K = 2
    P5394K = 3
    P539Package = 4
    PHKNCar = 5
    PHKN2K = 6
    PHKN3K = 7
    PHKN4K = 8
    PHKNSpecial = 9
    PLottoCar = 10
    PLotto2K = 11
    PLotto3K = 12
    PLotto4K = 13
    PLottoSpecial = 14
End Enum

Private Sub cmdClose_Click()
    Call Form_Unload(0)
End Sub

Private Sub cmdUpdate_Click()
'On Error GoTo errout:

    Dim flag As Boolean, n As Integer
    Dim PID(100) As String
    Dim LastSwiftCode As String, LastGroup As Integer
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim BonusTarget As String
    
    flag = False
    
    
    'get LastSwiftCode
    SQL = "select * from [order] order by SwiftCode desc;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, order_rec)
    If order_rec.EOF Then
        LastSwiftCode = "0"
    Else
        LastSwiftCode = order_rec("SwiftCode")
    End If
    order_rec.Close
    
    
    SQL = "select * from product;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    Do Until product_rec.EOF
        Select Case product_rec("PName")
        Case "539_車"
            PID(0) = product_rec("PID")
        Case "539_2K"
            PID(1) = product_rec("PID")
        Case "539_3K"
            PID(2) = product_rec("PID")
        Case "539_4K"
            PID(3) = product_rec("PID")
        Case "539_3包"
            PID(4) = product_rec("PID")
        Case "港號_車"
            PID(5) = product_rec("PID")
        Case "港號_2K"
            PID(6) = product_rec("PID")
        Case "港號_3K"
            PID(7) = product_rec("PID")
        Case "港號_4K"
            PID(8) = product_rec("PID")
        Case "港號_特"
            PID(9) = product_rec("PID")
        Case "大樂透_車"
            PID(10) = product_rec("PID")
        Case "大樂透_2K"
            PID(11) = product_rec("PID")
        Case "大樂透_3K"
            PID(12) = product_rec("PID")
        Case "大樂透_4K"
            PID(13) = product_rec("PID")
        Case "大樂透_特"
            PID(14) = product_rec("PID")
        End Select
        product_rec.MoveNext
    Loop
    product_rec.Close
    
    
    If cmbPName.Text = "" Then
        flag = False
    Else
        Select Case cmbPName.Text
        Case "100 539_全"
            n = 0
        Case "110 港號_全"
            n = 5
        Case "120 大樂透_全"
            n = 10
        Case Else
            n = -99   'goto errout
        End Select
        
        
        'get LastGroup
        SQL = "select * from [order] where CID='" & basVariable.SelectCID & "' and CurrentDate='" & Format(txtCurrentDate.Text, "yyyy/MM/dd") & "' and CLng(PID)>=" & PID(n) & " and CLng(PID)<=" & PID(n + 4) & " order by Group desc;"
        Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, order_rec)
        If order_rec.EOF Then
            LastGroup = -1
        Else
            If order_rec("Group") = Null Then
                'this solution is use for handle old data
                LastGroup = -1
            Else
                LastGroup = Val(order_rec("Group"))
            End If
        End If
        order_rec.Close
        
    
        'update all data from UI
        If txtCurrentCount_Car.Text <> "" Then
            Adodc1.Recordset.Fields.Item("CID").Value = basVariable.SelectCID
            Adodc1.Recordset.Fields.Item("CurrentDate").Value = txtCurrentDate.Text
            Adodc1.Recordset.Fields.Item("Note").Value = txtNote.Text
            Adodc1.Recordset.Fields.Item("Group").Value = LastGroup + 1
        
            LastSwiftCode = Val(LastSwiftCode) + 1
            Adodc1.Recordset.Fields.Item("SwiftCode").Value = LastSwiftCode
            Adodc1.Recordset.Fields.Item("PID").Value = PID(n)
            Adodc1.Recordset.Fields.Item("CurrentCount").Value = txtCurrentCount_Car.Text
            If txtWinningCount_Car.Text = "" Then
                Adodc1.Recordset.Fields.Item("WinningCount").Value = "0"
            Else
                Adodc1.Recordset.Fields.Item("WinningCount").Value = txtWinningCount_Car.Text
            End If
            Adodc1.Recordset.Fields.Item("AddMoney").Value = txtAddMoney.Text
            Adodc1.Recordset.Fields.Item("BonusMoney").Value = txtBonusMoney.Text
        
            Call Adodc1.Recordset.Update
            Call Adodc1.Recordset.AddNew    'Call Form_Unload(0)
            
            flag = True
        End If
        If txtCurrentCount_2K.Text <> "" Then
            Adodc1.Recordset.Fields.Item("CID").Value = basVariable.SelectCID
            Adodc1.Recordset.Fields.Item("CurrentDate").Value = txtCurrentDate.Text
            Adodc1.Recordset.Fields.Item("Note").Value = txtNote.Text
            Adodc1.Recordset.Fields.Item("Group").Value = LastGroup + 1
            
            LastSwiftCode = Val(LastSwiftCode) + 1
            Adodc1.Recordset.Fields.Item("SwiftCode").Value = LastSwiftCode
            Adodc1.Recordset.Fields.Item("PID").Value = PID(n + 1)
            Adodc1.Recordset.Fields.Item("CurrentCount").Value = txtCurrentCount_2K.Text
            If txtWinningCount_2K.Text = "" Then
                Adodc1.Recordset.Fields.Item("WinningCount").Value = "0"
            Else
                Adodc1.Recordset.Fields.Item("WinningCount").Value = txtWinningCount_2K.Text
            End If
            Adodc1.Recordset.Fields.Item("AddMoney").Value = txtAddMoney.Text
            Adodc1.Recordset.Fields.Item("BonusMoney").Value = txtBonusMoney.Text
        
            Call Adodc1.Recordset.Update
            Call Adodc1.Recordset.AddNew    'Call Form_Unload(0)
            
            flag = True
        End If
        If txtCurrentCount_3K.Text <> "" Then
            Adodc1.Recordset.Fields.Item("CID").Value = basVariable.SelectCID
            Adodc1.Recordset.Fields.Item("CurrentDate").Value = txtCurrentDate.Text
            Adodc1.Recordset.Fields.Item("Note").Value = txtNote.Text
            Adodc1.Recordset.Fields.Item("Group").Value = LastGroup + 1
            
            LastSwiftCode = Val(LastSwiftCode) + 1
            Adodc1.Recordset.Fields.Item("SwiftCode").Value = LastSwiftCode
            Adodc1.Recordset.Fields.Item("PID").Value = PID(n + 2)
            Adodc1.Recordset.Fields.Item("CurrentCount").Value = txtCurrentCount_3K.Text
            If txtWinningCount_3K.Text = "" Then
                Adodc1.Recordset.Fields.Item("WinningCount").Value = "0"
            Else
                Adodc1.Recordset.Fields.Item("WinningCount").Value = txtWinningCount_3K.Text
            End If
            Adodc1.Recordset.Fields.Item("AddMoney").Value = txtAddMoney.Text
            Adodc1.Recordset.Fields.Item("BonusMoney").Value = txtBonusMoney.Text
        
            Call Adodc1.Recordset.Update
            Call Adodc1.Recordset.AddNew    'Call Form_Unload(0)
            
            flag = True
        End If
        If txtCurrentCount_4K.Text <> "" Then
            Adodc1.Recordset.Fields.Item("CID").Value = basVariable.SelectCID
            Adodc1.Recordset.Fields.Item("CurrentDate").Value = txtCurrentDate.Text
            Adodc1.Recordset.Fields.Item("Note").Value = txtNote.Text
            Adodc1.Recordset.Fields.Item("Group").Value = LastGroup + 1
            
            LastSwiftCode = Val(LastSwiftCode) + 1
            Adodc1.Recordset.Fields.Item("SwiftCode").Value = LastSwiftCode
            Adodc1.Recordset.Fields.Item("PID").Value = PID(n + 3)
            Adodc1.Recordset.Fields.Item("CurrentCount").Value = txtCurrentCount_4K.Text
            If txtWinningCount_4K.Text = "" Then
                Adodc1.Recordset.Fields.Item("WinningCount").Value = "0"
            Else
                Adodc1.Recordset.Fields.Item("WinningCount").Value = txtWinningCount_4K.Text
            End If
            Adodc1.Recordset.Fields.Item("AddMoney").Value = txtAddMoney.Text
            Adodc1.Recordset.Fields.Item("BonusMoney").Value = txtBonusMoney.Text
        
            Call Adodc1.Recordset.Update
            Call Adodc1.Recordset.AddNew    'Call Form_Unload(0)
            
            flag = True
        End If
        If txtCurrentCount_Special.Text <> "" Then
            Adodc1.Recordset.Fields.Item("CID").Value = basVariable.SelectCID
            Adodc1.Recordset.Fields.Item("CurrentDate").Value = txtCurrentDate.Text
            Adodc1.Recordset.Fields.Item("Note").Value = txtNote.Text
            Adodc1.Recordset.Fields.Item("Group").Value = LastGroup + 1
            
            LastSwiftCode = Val(LastSwiftCode) + 1
            Adodc1.Recordset.Fields.Item("SwiftCode").Value = LastSwiftCode
            Adodc1.Recordset.Fields.Item("PID").Value = PID(n + 4)
            Adodc1.Recordset.Fields.Item("CurrentCount").Value = txtCurrentCount_Special.Text
            If txtWinningCount_Special.Text = "" Then
                Adodc1.Recordset.Fields.Item("WinningCount").Value = "0"
            Else
                Adodc1.Recordset.Fields.Item("WinningCount").Value = txtWinningCount_Special.Text
            End If
            Adodc1.Recordset.Fields.Item("AddMoney").Value = txtAddMoney.Text
            Adodc1.Recordset.Fields.Item("BonusMoney").Value = txtBonusMoney.Text
        
            Call Adodc1.Recordset.Update
            Call Adodc1.Recordset.AddNew    'Call Form_Unload(0)
            
            flag = True
        End If
        
        
        'clear all old data
        If flag Then
            dtpCurrentDate.Value = Format(DateTime.Now, "yyyy/MM/dd")
            txtCurrentDate.Text = Format(DateTime.Now, "yyyy/MM/dd")
            txtNote.Text = ""
            
            'clear all old data from ui
            txtCurrentCount_Car.Text = ""
            txtWinningCount_Car.Text = ""
            txtCurrentCount_2K.Text = ""
            txtWinningCount_2K.Text = ""
            txtCurrentCount_3K.Text = ""
            txtWinningCount_3K.Text = ""
            txtCurrentCount_4K.Text = ""
            txtWinningCount_4K.Text = ""
            txtCurrentCount_Special.Text = ""
            txtWinningCount_Special.Text = ""
            txtAddMoney.Text = ""
            txtBonusMoney.Text = ""
            
            addCount = addCount + 1
            lblAddCount.Caption = "已新增" & addCount & "筆"
        End If
    End If
    
    
    Call txtCurrentDate.SetFocus
    
    
    If Not flag Then
errout:
        MsgBox "輸入的資料有問題，或產品名稱、交易日期、交易數量未填寫！"
    End If
End Sub

Private Sub dtpCurrentDate_CloseUp()
    txtCurrentDate.Text = Format(dtpCurrentDate.Value, "yyyy/MM/dd")
End Sub

'import database and export to datagrid when form load
Private Sub Form_Load()
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
    'Set txtCurrentCount.DataSource = Adodc1
    'Set txtWinningCount.DataSource = Adodc1
    'Set txtAddMoney.DataSource = Adodc1
    'Set txtBonusMoney.DataSource = Adodc1
    'Set txtNote.DataSource = Adodc1


    'add new
    
    Call Adodc1.Recordset.AddNew
    
    
    Call cmbPName.AddItem("100 539_全")
    Call cmbPName.AddItem("110 港號_全")
    Call cmbPName.AddItem("120 大樂透_全")

    
    dtpCurrentDate.Value = Format(DateTime.Now, "yyyy/MM/dd")
    txtCurrentDate.Text = Format(DateTime.Now, "yyyy/MM/dd")
    
    
    addCount = 0
    lblAddCount.Caption = "已新增" & addCount & "筆"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmOrder.Show
    Unload Me
End Sub


'KeyUp
Private Sub txtCurrentDate_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call cmbPName.SetFocus
    End If
End Sub
Private Sub cmbPName_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtCurrentCount_Car.SetFocus
    End If
End Sub
Private Sub txtCurrentCount_Car_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtWinningCount_Car.SetFocus
    End If
End Sub
Private Sub txtWinningCount_Car_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtCurrentCount_2K.SetFocus
    End If
End Sub
Private Sub txtCurrentCount_2K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtWinningCount_2K.SetFocus
    End If
End Sub
Private Sub txtWinningCount_2K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtCurrentCount_3K.SetFocus
    End If
End Sub
Private Sub txtCurrentCount_3K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtWinningCount_3K.SetFocus
    End If
End Sub
Private Sub txtWinningCount_3K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtCurrentCount_4K.SetFocus
    End If
End Sub
Private Sub txtCurrentCount_4K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtWinningCount_4K.SetFocus
    End If
End Sub
Private Sub txtWinningCount_4K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtCurrentCount_Special.SetFocus
    End If
End Sub
Private Sub txtCurrentCount_Special_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtWinningCount_Special.SetFocus
    End If
End Sub
Private Sub txtWinningCount_Special_KeyUp(KeyCode As Integer, Shift As Integer)
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
