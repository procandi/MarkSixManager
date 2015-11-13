VERSION 5.00
Begin VB.Form frmProve 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  '單線固定
   Caption         =   "系統主頁面"
   ClientHeight    =   13110
   ClientLeft      =   5610
   ClientTop       =   825
   ClientWidth     =   6765
   Icon            =   "frmProve.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   13110
   ScaleWidth      =   6765
   Begin VB.Frame Frame4 
      Caption         =   "客別不分產品報表"
      Height          =   1695
      Left            =   240
      TabIndex        =   34
      Top             =   11160
      Width           =   6135
      Begin VB.CommandButton cmdCustromYearTransaction 
         Caption         =   "客別不分產品年交易金額表"
         Height          =   615
         Left            =   4680
         TabIndex        =   41
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdCustromMonthTransaction 
         Caption         =   "客別不分產品月交易金額表"
         Height          =   615
         Left            =   3240
         TabIndex        =   40
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdCustromWeekTransaction 
         Caption         =   "客別不分產品週交易金額表"
         Height          =   615
         Left            =   1680
         TabIndex        =   39
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdCustromYearReport 
         Caption         =   "客別不分產品年報表"
         Height          =   615
         Left            =   4680
         TabIndex        =   37
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdCustromMonthReport 
         Caption         =   "客別不分產品月報表"
         Height          =   615
         Left            =   3240
         TabIndex        =   36
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdCustromWeekReport 
         Caption         =   "客別不分產品週報表"
         Height          =   615
         Left            =   1680
         TabIndex        =   35
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "客別分產品報表"
      Height          =   1695
      Left            =   240
      TabIndex        =   31
      Top             =   9360
      Width           =   6135
      Begin VB.CommandButton cmdCustromProductWeekTransaction 
         Caption         =   "客別分產品週交易金額表"
         Height          =   615
         Left            =   1680
         TabIndex        =   38
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdCustromProductWeekReport 
         Caption         =   "客別分產品週報表"
         Height          =   615
         Left            =   1680
         TabIndex        =   33
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdCustromProductDayReport 
         Caption         =   "客別分產品日報表"
         Height          =   615
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "4K總表"
      Height          =   1695
      Left            =   240
      TabIndex        =   22
      Top             =   7440
      Width           =   6135
      Begin VB.CommandButton cmdFourKYearReport 
         Caption         =   "4K年報表"
         Height          =   615
         Left            =   4680
         TabIndex        =   30
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdFourKMonthReport 
         Caption         =   "4K月報表"
         Height          =   615
         Left            =   3240
         TabIndex        =   29
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdFourKWeekReport 
         Caption         =   "4K週報表"
         Height          =   615
         Left            =   1680
         TabIndex        =   28
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdFourKDayReport 
         Caption         =   "4K日報表"
         Height          =   615
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdFourKYearAccount 
         Caption         =   "4K年總帳"
         Height          =   615
         Left            =   4680
         TabIndex        =   26
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdFourKMonthAccount 
         Caption         =   "4K月總帳"
         Height          =   615
         Left            =   3240
         TabIndex        =   25
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdFourKWeekAccount 
         Caption         =   "4K週總帳"
         Height          =   615
         Left            =   1680
         TabIndex        =   24
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdFourKDayAccount 
         Caption         =   "4K日總帳"
         Height          =   615
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "一般總表"
      Height          =   1695
      Left            =   240
      TabIndex        =   13
      Top             =   5520
      Width           =   6135
      Begin VB.CommandButton cmdDayAccount 
         Caption         =   "日總帳"
         Height          =   615
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdYearAccount 
         Caption         =   "年總帳"
         Height          =   615
         Left            =   4680
         TabIndex        =   20
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdMonthAccount 
         Caption         =   "月總帳"
         Height          =   615
         Left            =   3240
         TabIndex        =   19
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdWeekAccount 
         Caption         =   "週總帳"
         Height          =   615
         Left            =   1680
         TabIndex        =   18
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdYearReport 
         Caption         =   "年報表"
         Height          =   615
         Left            =   4680
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdDayReport 
         Caption         =   "日報表"
         Height          =   615
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdWeekReport 
         Caption         =   "週報表"
         Height          =   615
         Left            =   1680
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdMonthReport 
         Caption         =   "月報表"
         Height          =   615
         Left            =   3240
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdProduct 
      Caption         =   "產品資料"
      Height          =   615
      Left            =   4080
      TabIndex        =   12
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdCustom 
      Caption         =   "客戶資料"
      Height          =   615
      Left            =   1320
      TabIndex        =   11
      Top             =   4800
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
      Top             =   3720
      Width           =   6465
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0C0&
      Height          =   3255
      Left            =   120
      ScaleHeight     =   3195
      ScaleWidth      =   6375
      TabIndex        =   5
      Top             =   120
      Width           =   6435
      Begin VB.Image Image1 
         Height          =   1035
         Left            =   240
         Picture         =   "frmProve.frx":8ACE
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
         Top             =   1320
         Width           =   5760
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
         Left            =   4440
         TabIndex        =   7
         Top             =   2880
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '透明
         Caption         =   $"frmProve.frx":3840F
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
         Top             =   2280
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
      Caption         =   "Version 20151111"
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
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   240
      X2              =   6459
      Y1              =   4680
      Y2              =   4680
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
      Top             =   4320
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
      Height          =   8715
      Index           =   5
      Left            =   120
      TabIndex        =   3
      Top             =   4320
      Width           =   6465
   End
End
Attribute VB_Name = "frmProve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCustom_Click()
    basVariable.Action = "CustomDetail"
    frmCustom.Show
    Me.Hide
End Sub

Private Sub cmdCustromMonthTransaction_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "CustromMonthTransaction"
    frmConfirm.Show
End Sub

Private Sub cmdCustromProductDayReport_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "CustromProductDayReport"
    frmConfirm.Show
End Sub

Private Sub cmdCustromProductWeekReport_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "CustromProductWeekReport"
    frmConfirm.Show
End Sub

Private Sub cmdCustromProductWeekTransaction_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "CustromProductWeekTransaction"
    frmConfirm.Show
End Sub

Private Sub cmdCustromWeekReport_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "CustromWeekReport"
    frmConfirm.Show
End Sub

Private Sub cmdCustromMonthReport_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "CustromMonthReport"
    frmConfirm.Show
End Sub

Private Sub cmdCustromWeekTransaction_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "CustromWeekTransaction"
    frmConfirm.Show
End Sub

Private Sub cmdCustromYearReport_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "CustromYearReport"
    frmConfirm.Show
End Sub

Private Sub cmdCustromYearTransaction_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "CustromYearTransaction"
    frmConfirm.Show
End Sub

Private Sub cmdDayAccount_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "DayAccount"
    frmConfirm.Show
End Sub

Private Sub cmdDayReport_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "DayReport"
    frmConfirm.Show
End Sub

Private Sub cmdFourKDayAccount_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "FourKDayAccount"
    frmConfirm.Show
End Sub

Private Sub cmdFourKDayReport_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "FourKDayReport"
    frmConfirm.Show
End Sub

Private Sub cmdFourKMonthAccount_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "FourKMonthAccount"
    frmConfirm.Show
End Sub

Private Sub cmdFourKMonthReport_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "FourKMonthReport"
    frmConfirm.Show
End Sub

Private Sub cmdFourKWeekAccount_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "FourKWeekAccount"
    frmConfirm.Show
End Sub

Private Sub cmdFourKWeekReport_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "FourKWeekReport"
    frmConfirm.Show
End Sub

Private Sub cmdFourKYearAccount_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "FourKYearAccount"
    frmConfirm.Show
End Sub

Private Sub cmdFourKYearReport_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "FourKYearReport"
    frmConfirm.Show
End Sub

Private Sub cmdMonthAccount_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "MonthAccount"
    frmConfirm.Show
End Sub

Private Sub cmdMonthReport_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "MonthReport"
    frmConfirm.Show
End Sub

Private Sub cmdProduct_Click()
    basVariable.Action = "ProductDetail"
    frmProduct.Show
    Me.Hide
End Sub

Private Sub cmdWeekAccount_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "WeekAccount"
    frmConfirm.Show
End Sub

Private Sub cmdWeekReport_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "WeekReport"
    frmConfirm.Show
End Sub

Private Sub cmdYearAccount_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "YearAccount"
    frmConfirm.Show
End Sub

Private Sub cmdYearReport_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "YearReport"
    frmConfirm.Show
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & GetVersion()

    basDataBase.Connection_String = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\") & "main.mdb;"
    'basDataBase.Connection_String = "Driver=SQLite3 ODBC Driver;Database=main.db;"
    
    Call basDataBase.Connect2DataBase(basDataBase.Connection_String, basDataBase.Connection)
    
    'For i = 1 To 127
    '    Debug.Print i & "," & Chr(i)
    'Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
