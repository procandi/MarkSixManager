VERSION 5.00
Begin VB.Form frmProveNew 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  '單線固定
   Caption         =   "系統主頁面"
   ClientHeight    =   11355
   ClientLeft      =   5055
   ClientTop       =   525
   ClientWidth     =   6765
   Icon            =   "frmProveNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11355
   ScaleWidth      =   6765
   Begin VB.Frame Frame1 
      Caption         =   "一般報表"
      Height          =   5655
      Left            =   240
      TabIndex        =   13
      Top             =   5280
      Width           =   6135
      Begin VB.CommandButton cmdCustomDailyPriceDetail 
         Caption         =   "客戶每日交易價格表"
         Height          =   975
         Left            =   120
         TabIndex        =   22
         Top             =   4560
         Width           =   5895
      End
      Begin VB.CommandButton cmdAllMonthTransactionCounting 
         Caption         =   "全產品當月交易加總表"
         Height          =   975
         Left            =   120
         TabIndex        =   21
         Top             =   2400
         Width           =   2895
      End
      Begin VB.CommandButton cmdAllMonth4KTransactionCounting 
         Caption         =   "全產品4K當月交易加總表"
         Height          =   975
         Left            =   3120
         TabIndex        =   20
         Top             =   3480
         Width           =   2895
      End
      Begin VB.CommandButton cmdAllDaily4KTransactionCounting 
         Caption         =   "全產品4K每日交易加總表"
         Height          =   975
         Left            =   120
         TabIndex        =   19
         Top             =   3480
         Width           =   2895
      End
      Begin VB.CommandButton cmdMonthTransactionCounting 
         Caption         =   "當月交易加總表"
         Height          =   975
         Left            =   3120
         TabIndex        =   18
         Top             =   2400
         Width           =   2895
      End
      Begin VB.CommandButton cmdAllWeekTransactionCounting 
         Caption         =   "全產品一週交易加總表"
         Height          =   975
         Left            =   3120
         TabIndex        =   17
         Top             =   1320
         Width           =   2895
      End
      Begin VB.CommandButton cmdCustomDailyTransactionDetail 
         Caption         =   "客戶每日交易明細"
         Height          =   975
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   2895
      End
      Begin VB.CommandButton cmdDailyTransactionCounting 
         Caption         =   "每日交易加總表"
         Height          =   975
         Left            =   3120
         TabIndex        =   15
         Top             =   240
         Width           =   2895
      End
      Begin VB.CommandButton cmdAllDailyTransactionCounting 
         Caption         =   "全產品每日交易加總表"
         Height          =   975
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdProduct 
      Caption         =   "產品資料"
      Height          =   615
      Left            =   4080
      TabIndex        =   12
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdCustom 
      Caption         =   "客戶資料"
      Height          =   615
      Left            =   1320
      TabIndex        =   11
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox txtUpdateNote 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   10
      Top             =   3360
      Width           =   6465
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0C0&
      Height          =   2775
      Left            =   120
      ScaleHeight     =   2715
      ScaleWidth      =   6375
      TabIndex        =   5
      Top             =   120
      Width           =   6435
      Begin VB.Image Image1 
         Height          =   1035
         Left            =   240
         Picture         =   "frmProveNew.frx":8ACE
         Stretch         =   -1  'True
         Top             =   120
         Width           =   5820
      End
      Begin VB.Label lblPlatform 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BorderStyle     =   1  '單線固定
         Caption         =   "進銷存系統"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   660
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Tag             =   "1052"
         Top             =   1200
         Width           =   5745
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00400000&
         BorderStyle     =   1  '單線固定
         Caption         =   "米飛爾科技"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4560
         TabIndex        =   7
         Top             =   2400
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '透明
         Caption         =   $"frmProveNew.frx":3840F
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   360
         TabIndex        =   6
         Top             =   1920
         Width           =   5775
      End
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00808080&
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
      Height          =   406
      Left            =   7890
      TabIndex        =   0
      Top             =   4710
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.TextBox txtCode 
      BackColor       =   &H00808080&
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
      Height          =   406
      Left            =   7890
      TabIndex        =   2
      Top             =   4320
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.TextBox txtType 
      BackColor       =   &H00808080&
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
      Height          =   406
      Left            =   7890
      TabIndex        =   1
      Top             =   3960
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00808080&
      BorderStyle     =   1  '單線固定
      Caption         =   "Version 20151216"
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
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   240
      X2              =   6459
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label lblLabels 
      Appearance      =   0  '平面
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  '透明
      Caption         =   "系統功能"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Index           =   6
      Left            =   240
      TabIndex        =   4
      Top             =   4080
      Width           =   6120
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00808080&
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Index           =   5
      Left            =   120
      TabIndex        =   3
      Top             =   3960
      Width           =   6465
   End
End
Attribute VB_Name = "frmProveNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAllDaily4KTransactionCounting_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "AllDaily4KTransactionCounting"
    frmConfirmRuby.Show
End Sub

Private Sub cmdAllDailyTransactionCounting_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "AllDailyTransactionCounting"
    frmConfirmRuby.Show
End Sub

Private Sub cmdAllMonth4KTransactionCounting_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "AllMonth4KTransactionCounting"
    frmConfirmRuby.Show
End Sub

Private Sub cmdAllMonthTransactionCounting_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "AllMonthTransactionCounting"
    frmConfirmRuby.Show
End Sub

Private Sub cmdAllWeekTransactionCounting_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "AllWeekTransactionCounting"
    frmConfirmRuby.Show
End Sub

Private Sub cmdCustom_Click()
    basVariable.Action = "CustomDetail"
    frmCustom.Show
    Me.Hide
End Sub

Private Sub cmdCustomDailyPriceDetail_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "CustomDailyPriceDetail"
    frmConfirmRuby.Show
End Sub

Private Sub cmdCustomDailyTransactionDetail_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "CustomDailyTransactionDetail"
    frmConfirmRuby.Show
End Sub

Private Sub cmdDailyTransactionCounting_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "DailyTransactionCounting"
    frmConfirmRuby.Show
End Sub

Private Sub cmdMonthTransactionCounting_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "MonthTransactionCounting"
    frmConfirmRuby.Show
End Sub

Private Sub cmdProduct_Click()
    basVariable.Action = "ProductDetail"
    frmProduct.Show
    Me.Hide
End Sub

Private Sub Form_Load()
    Dim strWant As String, want() As Double, M() As Double, iM() As Double
    Dim temp() As String, inp As String
    Dim DeCode As String, EnCode As String

    Dim fpath As String
    fpath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
    
    Dim flag As Boolean


    'check serial number file
    If FSO.FileExists(fpath & "key") Then
        flag = True
    Else
        flag = False
    End If

    
    'read serial number
    If flag Then
        Open fpath & "key" For Input As #1
            Input #1, inp
        Close #1
    Else
        inp = InputBox("請輸入序號")
    End If


    'check serial number
    Call MakeMatrix(M, iM)
    
    DeCode = "MartSixManager_" & GetPhysicalAddress
    Call MatrixEncode(DeCode, M, want)
    
    EnCode = ""
    For i = 1 To UBound(want)
        EnCode = EnCode & want(i) & " "
    Next
    EnCode = Left(EnCode, Len(EnCode) - 1)
    
    'crack
    'inp = EnCode
    
    'If EnCode = inp Then
    If DateTime.DateDiff("d", DateTime.Now, "2016/07/01") > 0 Then
        'write serial number file when no file exist and check currect
        If Not flag Then
            Open fpath & "key" For Output As #1
                Write #1, EnCode
            Close #1
        End If
        
        'check ok
        flag = True
    Else
        'check fail
        flag = False
    End If
    
    
    'into system or exit system
    If flag Then
        'update version information
        lblVersion.Caption = "Version " & GetVersion()
    
    
        'connect to database
        basDataBase.Connection_String = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & fpath & "main.mdb;"
        'basDataBase.Connection_String = "Driver=SQLite3 ODBC Driver;Database=main.db;"
        
        Call basDataBase.Connect2DataBase(basDataBase.Connection_String, basDataBase.Connection)
    Else
        'fail message
        MsgBox "無註冊資訊或序號不正確！"
        End
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
