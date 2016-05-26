VERSION 5.00
Begin VB.Form frmNewProve 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  '單線固定
   Caption         =   "系統主頁面"
   ClientHeight    =   14250
   ClientLeft      =   5055
   ClientTop       =   525
   ClientWidth     =   6765
   Icon            =   "frmNewProve.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   14250
   ScaleWidth      =   6765
   Begin VB.Frame Frame6 
      Caption         =   "不分客別、分產品的報表"
      Height          =   975
      Left            =   240
      TabIndex        =   46
      Top             =   11760
      Width           =   6135
      Begin VB.CommandButton cmdProductWeekTransaction 
         Caption         =   "週交易金額表"
         Height          =   615
         Left            =   1680
         TabIndex        =   47
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "不分客別、不分產品的報表"
      Height          =   975
      Left            =   240
      TabIndex        =   39
      Top             =   12840
      Width           =   6135
      Begin VB.CommandButton cmdYearTransaction 
         Caption         =   "年交易金額表"
         Height          =   615
         Left            =   4680
         TabIndex        =   45
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdMonthTransaction 
         Caption         =   "月交易金額表"
         Height          =   615
         Left            =   3240
         TabIndex        =   44
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdWeekTransaction 
         Caption         =   "週交易金額表"
         Height          =   615
         Left            =   1680
         TabIndex        =   43
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "分客別、不分產品報表"
      Height          =   975
      Left            =   240
      TabIndex        =   34
      Top             =   10680
      Width           =   6135
      Begin VB.CommandButton cmdCustomYearTransaction 
         Caption         =   "年交易金額表"
         Height          =   615
         Left            =   4680
         TabIndex        =   42
         Top             =   -120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdCustomMonthTransaction 
         Caption         =   "月交易金額表"
         Height          =   615
         Left            =   3240
         TabIndex        =   41
         Top             =   -120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdCustomWeekTransaction 
         Caption         =   "週交易金額表"
         Height          =   615
         Left            =   1680
         TabIndex        =   40
         Top             =   -120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdCustomYearReport 
         Caption         =   "年報表"
         Height          =   615
         Left            =   4680
         TabIndex        =   37
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdCustomMonthReport 
         Caption         =   "月報表"
         Height          =   615
         Left            =   3240
         TabIndex        =   36
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdCustomWeekReport 
         Caption         =   "週報表"
         Height          =   615
         Left            =   1680
         TabIndex        =   35
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "分客別、分產品報表"
      Height          =   1695
      Left            =   240
      TabIndex        =   31
      Top             =   8880
      Width           =   6135
      Begin VB.CommandButton cmdCustomProductDayReportDetail 
         Caption         =   "日明細表"
         Height          =   615
         Left            =   240
         TabIndex        =   49
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdCustomProductWeekReportDetail 
         Caption         =   "週明細表"
         Height          =   615
         Left            =   1680
         TabIndex        =   48
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdCustomProductWeekTransaction 
         Caption         =   "週交易金額表"
         Height          =   615
         Left            =   1680
         TabIndex        =   38
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdCustomProductWeekReport 
         Caption         =   "週報表"
         Height          =   615
         Left            =   1680
         TabIndex        =   33
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdCustomProductDayReport 
         Caption         =   "日報表"
         Height          =   615
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "4K報表"
      Height          =   1695
      Left            =   240
      TabIndex        =   22
      Top             =   7080
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
      Caption         =   "一般報表"
      Height          =   1695
      Left            =   240
      TabIndex        =   13
      Top             =   5280
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
         Picture         =   "frmNewProve.frx":8ACE
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
         Caption         =   $"frmNewProve.frx":3840F
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
      Height          =   10095
      Index           =   5
      Left            =   120
      TabIndex        =   3
      Top             =   3960
      Width           =   6465
   End
End
Attribute VB_Name = "frmNewProve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCustom_Click()
    basVariable.Action = "CustomDetail"
    frmCustom.Show
    Me.Hide
End Sub

Private Sub cmdCustomMonthTransaction_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "CustomMonthTransaction"
    frmConfirmXLS.Show
End Sub

Private Sub cmdCustomProductDayReport_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "CustomProductDayReport"
    frmConfirmXLS.Show
End Sub

Private Sub cmdCustomProductDayReportDetail_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "CustomProductDayReportDetail"
    frmConfirmXLS.Show
End Sub

Private Sub cmdCustomProductWeekReport_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "CustomProductWeekReport"
    frmConfirmXLS.Show
End Sub

Private Sub cmdCustomProductWeekReportDetail_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "CustomProductWeekReportDetail"
    frmConfirmXLS.Show
End Sub

Private Sub cmdCustomProductWeekTransaction_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "CustomProductWeekTransaction"
    frmConfirmXLS.Show
End Sub

Private Sub cmdCustomWeekReport_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "CustomWeekReport"
    frmConfirmXLS.Show
End Sub

Private Sub cmdCustomMonthReport_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "CustomMonthReport"
    frmConfirmXLS.Show
End Sub

Private Sub cmdCustomWeekTransaction_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "CustomWeekTransaction"
    frmConfirmXLS.Show
End Sub

Private Sub cmdCustomYearReport_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "CustomYearReport"
    frmConfirmXLS.Show
End Sub

Private Sub cmdCustomYearTransaction_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "CustomYearTransaction"
    frmConfirmXLS.Show
End Sub

Private Sub cmdDayAccount_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "DayAccount"
    frmConfirmXLS.Show
End Sub

Private Sub cmdDayReport_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "DayReport"
    frmConfirmXLS.Show
End Sub

Private Sub cmdFourKDayAccount_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "FourKDayAccount"
    frmConfirmXLS.Show
End Sub

Private Sub cmdFourKDayReport_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "FourKDayReport"
    frmConfirmXLS.Show
End Sub

Private Sub cmdFourKMonthAccount_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "FourKMonthAccount"
    frmConfirmXLS.Show
End Sub

Private Sub cmdFourKMonthReport_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "FourKMonthReport"
    frmConfirmXLS.Show
End Sub

Private Sub cmdFourKWeekAccount_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "FourKWeekAccount"
    frmConfirmXLS.Show
End Sub

Private Sub cmdFourKWeekReport_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "FourKWeekReport"
    frmConfirmXLS.Show
End Sub

Private Sub cmdFourKYearAccount_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "FourKYearAccount"
    frmConfirmXLS.Show
End Sub

Private Sub cmdFourKYearReport_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "FourKYearReport"
    frmConfirmXLS.Show
End Sub

Private Sub cmdMonthAccount_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "MonthAccount"
    frmConfirmXLS.Show
End Sub

Private Sub cmdMonthReport_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "MonthReport"
    frmConfirmXLS.Show
End Sub

Private Sub cmdMonthTransaction_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "MonthTransaction"
    frmConfirmXLS.Show
End Sub

Private Sub cmdProduct_Click()
    basVariable.Action = "ProductDetail"
    frmProduct.Show
    Me.Hide
End Sub

Private Sub cmdProductWeekTransaction_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "ProductWeekTransaction"
    frmConfirmXLS.Show
End Sub

Private Sub cmdWeekAccount_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "WeekAccount"
    frmConfirmXLS.Show
End Sub

Private Sub cmdWeekReport_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "WeekReport"
    frmConfirmXLS.Show
End Sub

Private Sub cmdWeekTransaction_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "WeekTransaction"
    frmConfirmXLS.Show
End Sub

Private Sub cmdYearAccount_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "YearAccount"
    frmConfirmXLS.Show
End Sub

Private Sub cmdYearReport_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "YearReport"
    frmConfirmXLS.Show
End Sub

Private Sub cmdYearTransaction_Click()
    basVariable.Action = "PrintReport"
    basVariable.Parameter = "YearTransaction"
    frmConfirmXLS.Show
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
    
    If EnCode = inp And DateTime.DateDiff("d", DateTime.Now, "2016/07/01") > 0 Then
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
