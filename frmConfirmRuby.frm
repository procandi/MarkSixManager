VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmConfirmRuby 
   BorderStyle     =   1  '單線固定
   Caption         =   "Report Confirm"
   ClientHeight    =   3990
   ClientLeft      =   6315
   ClientTop       =   7785
   ClientWidth     =   4680
   Icon            =   "frmConfirmRuby.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   4680
   Begin VB.ComboBox cmbPName 
      Height          =   300
      Left            =   1800
      TabIndex        =   11
      Top             =   1800
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ComboBox cmbCName 
      Height          =   300
      Left            =   1800
      TabIndex        =   10
      Top             =   840
      Visible         =   0   'False
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
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1320
      Width           =   1650
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&N 取　消"
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
      Left            =   2400
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton cmdConfirm 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&O 確　定"
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
      Left            =   0
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   3360
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker dtpCurrentDate 
      Height          =   360
      Left            =   1800
      TabIndex        =   5
      Top             =   1320
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
      Format          =   94896131
      CurrentDate     =   37058
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
      Index           =   2
      Left            =   720
      TabIndex        =   9
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblEntry 
      Alignment       =   1  '靠右對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "客戶名稱"
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
      Left            =   720
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00808080&
      BackStyle       =   0  '透明
      Caption         =   "日報表(總帳)以選定的那天報表輸出。週、月、年報表(總帳)以選定的那天的當週、月、年報表輸出。"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   855
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   2280
      Width           =   4215
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
      Left            =   720
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00808080&
      BackStyle       =   0  '透明
      Caption         =   "報表列印"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00808080&
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmConfirmRuby"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Call Form_Unload(0)
End Sub

Private Sub cmdConfirm_Click()
    If txtCurrentDate.Text = "" Then
        MsgBox "請先選擇要列印的時間！"
    ElseIf (basVariable.Parameter = "CustomProductDayReport" Or basVariable.Parameter = "CustomProductWeekReport") And cmbCName.Text = "" And cmbPName.Text = "" Then
        MsgBox "尚未選擇客戶或產品！"
    ElseIf (basVariable.Parameter = "CustomWeekReport" Or basVariable.Parameter = "CustomMonthReport" Or basVariable.Parameter = "CustomYearReport") And cmbCName.Text = "" Then
        MsgBox "尚未選擇客戶！"
    Else
        Dim CData() As String
        Dim PData() As String
            
        Select Case basVariable.Parameter
        Case "CustomDailyTransactionDetail"
            'Label1(0).Caption = "客戶每日交易明細"
            
            CData = Split(cmbCName.Text, " ")
            PData = Split(cmbPName.Text, " ")
            Call Shell("C:\Ruby22-x64\bin\ruby.exe -Ku main.rb" & " DailyTransactionCounting " & PData(0) & " " & CData(0) & " " & txtCurrentDate.Text)
            
        Case "DailyTransactionCounting"
            'Label1(0).Caption = "每日交易加總表"
            
            PData = Split(cmbPName.Text, " ")
            Call Shell("C:\Ruby22-x64\bin\ruby.exe -Ku main.rb" & " DailyTransactionCounting " & PData(0) & " " & txtCurrentDate.Text)
            
        Case "AllDailyTransactionCounting"
            'Label1(0).Caption = "全產品每日交易加總表"
            
            Call Shell("C:\Ruby22-x64\bin\ruby.exe -Ku main.rb" & " AllDailyTransactionCounting " & txtCurrentDate.Text)
            
        Case "AllWeekTransactionCounting"
            'Label1(0).Caption = "全產品一週交易加總表"
            
            Call Shell("C:\Ruby22-x64\bin\ruby.exe -Ku main.rb" & " AllWeekTransactionCounting " & txtCurrentDate.Text)
            
        Case "AllMonthTransactionCounting"
            'Label1(0).Caption = "全產品當月交易加總表"
            
            Call Shell("C:\Ruby22-x64\bin\ruby.exe -Ku main.rb" & " AllMonthTransactionCounting " & txtCurrentDate.Text)
            
        Case "MonthTransactionCounting"
            'Label1(0).Caption = "當月交易加總表"
            
            PData = Split(cmbPName.Text, " ")
            Call Shell("C:\Ruby22-x64\bin\ruby.exe -Ku main.rb" & " MonthTransactionCounting " & PData(0) & " " & txtCurrentDate.Text)
            
        Case "AllDaily4KTransactionCounting"
            'Label1(0).Caption = "全產品4K每日交易加總表"
            
            Call Shell("C:\Ruby22-x64\bin\ruby.exe -Ku main.rb" & " AllDaily4KTransactionCounting " & txtCurrentDate.Text)
            
        Case "AllMonth4KTransactionCounting"
            'Label1(0).Caption = "全產品4K當月交易加總表"
            
            Call Shell("C:\Ruby22-x64\bin\ruby.exe -Ku main.rb" & " AllMonth4KTransactionCounting " & txtCurrentDate.Text)
            
        Case "CustomDailyPriceDetail"
            'Label1(0).Caption = "客戶每日交易價格表"
            
            Call Shell("C:\Ruby22-x64\bin\ruby.exe -Ku main.rb" & " CustomDailyPriceDetail " & txtCurrentDate.Text)
        End Select
        
        
        MsgBox "OK"
        'Call Shell("C:\Ruby22-x64\bin\ruby.exe -Ku test.rb test")
    End If
End Sub

Private Sub dtpCurrentDate_CloseUp()
    txtCurrentDate.Text = Format(dtpCurrentDate.Value, "yyyy/MM/dd")
End Sub

Private Sub Form_Load()
    Select Case basVariable.Parameter
    Case "CustomDailyTransactionDetail"
        Label1(0).Caption = "客戶每日交易明細"
        
        lblEntry(0).Visible = True
        lblEntry(2).Visible = True
        cmbCName.Visible = True
        cmbPName.Visible = True
        Call ComboBox_LoadFrom_DataBase_ByFile(cmbCName, "CID,CName", "custom", "", "", "")
        
        Call cmbPName.AddItem("100 539_全")
        Call cmbPName.AddItem("110 港號_全")
        Call cmbPName.AddItem("120 大樂透_全")
        
    Case "DailyTransactionCounting"
        Label1(0).Caption = "每日交易加總表"
        
        lblEntry(2).Visible = True
        cmbPName.Visible = True
        
        Call cmbPName.AddItem("100 539_全")
        Call cmbPName.AddItem("110 港號_全")
        Call cmbPName.AddItem("120 大樂透_全")
        
    Case "AllDailyTransactionCounting"
        Label1(0).Caption = "全產品每日交易加總表"
        
    Case "AllWeekTransactionCounting"
        Label1(0).Caption = "全產品一週交易加總表"
        
    Case "AllMonthTransactionCounting"
        Label1(0).Caption = "全產品當月交易加總表"
        
    Case "MonthTransactionCounting"
        Label1(0).Caption = "當月交易加總表"
        
        lblEntry(2).Visible = True
        cmbPName.Visible = True
        
        Call cmbPName.AddItem("100 539_全")
        Call cmbPName.AddItem("110 港號_全")
        Call cmbPName.AddItem("120 大樂透_全")
        
    Case "AllDaily4KTransactionCounting"
        Label1(0).Caption = "全產品4K每日交易加總表"
        
    Case "AllMonth4KTransactionCounting"
        Label1(0).Caption = "全產品4K當月交易加總表"
        
    Case "CustomDailyPriceDetail"
        Label1(0).Caption = "客戶每日交易價格表"

    End Select
    
    
    dtpCurrentDate.Value = Format(DateTime.Now, "yyyy/MM/dd")
    txtCurrentDate.Text = Format(DateTime.Now, "yyyy/MM/dd")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmProveNew.Show
    Unload Me
End Sub
