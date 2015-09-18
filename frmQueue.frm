VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmQueue 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  '單線固定
   Caption         =   "CRS 檢查/報告作業"
   ClientHeight    =   11040
   ClientLeft      =   2250
   ClientTop       =   1125
   ClientWidth     =   15270
   DrawWidth       =   3
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmQueue.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   11040
   ScaleWidth      =   15270
   StartUpPosition =   2  '螢幕中央
   Begin Threed.SSCommand SSCommand1 
      Height          =   495
      Left            =   12480
      TabIndex        =   113
      Top             =   1800
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "確認全部報告"
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
   End
   Begin VB.ListBox lstFilter 
      Height          =   1035
      Left            =   1200
      TabIndex        =   73
      Top             =   2400
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtStatus 
      Height          =   375
      Left            =   7320
      TabIndex        =   66
      Text            =   "Status"
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin Threed.SSPanel pnlBasic 
      Height          =   2535
      Left            =   120
      TabIndex        =   12
      Top             =   7320
      Width           =   7455
      _Version        =   65536
      _ExtentX        =   13150
      _ExtentY        =   4471
      _StockProps     =   15
      Caption         =   "  "
      ForeColor       =   0
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      Outline         =   -1  'True
      Font3D          =   3
      Alignment       =   0
      Begin VB.CommandButton cmdBasicEdit 
         BackColor       =   &H00FFC0C0&
         Caption         =   "更改基本資料"
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
         Left            =   4680
         Style           =   1  '圖片外觀
         TabIndex        =   80
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00404040&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   4
         Left            =   1320
         TabIndex        =   24
         Top             =   2040
         Width           =   6015
      End
      Begin VB.Label lblBasic 
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
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   23
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00404040&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   5
         Left            =   1320
         TabIndex        =   26
         Top             =   1680
         Width           =   6015
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "電話號碼"
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
         Index           =   5
         Left            =   120
         TabIndex        =   25
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00404040&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   6
         Left            =   1320
         TabIndex        =   31
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "身份證號"
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
         Index           =   6
         Left            =   120
         TabIndex        =   30
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00404040&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   1320
         TabIndex        =   20
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "出生日期"
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
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00404040&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   4680
         TabIndex        =   22
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "性別"
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
         Index           =   3
         Left            =   3480
         TabIndex        =   21
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00800000&
         BorderStyle     =   1  '單線固定
         Caption         =   "受檢病患基本資料"
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
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   120
         Width           =   3255
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00404040&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   4680
         TabIndex        =   18
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label lblBasic 
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
         Height          =   375
         Index           =   1
         Left            =   3480
         TabIndex        =   17
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00404040&
         BorderStyle     =   1  '單線固定
         DataField       =   "ChartNo"
         DataSource      =   "datBasic"
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
         Left            =   1320
         TabIndex        =   16
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "病歷號碼"
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
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1215
      End
   End
   Begin Threed.SSPanel pnlExamDetail 
      Height          =   2535
      Left            =   120
      TabIndex        =   45
      Top             =   7320
      Visible         =   0   'False
      Width           =   7095
      _Version        =   65536
      _ExtentX        =   12515
      _ExtentY        =   4471
      _StockProps     =   15
      Caption         =   " "
      ForeColor       =   12582912
      BackColor       =   4210752
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
      BevelInner      =   1
      Outline         =   -1  'True
      Font3D          =   3
      Alignment       =   1
      Begin VB.ListBox lstExamDetailEdit 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Left            =   240
         MultiSelect     =   1  '簡易多重選取
         TabIndex        =   47
         Top             =   600
         Width           =   7215
      End
      Begin VB.ListBox lstExamSpecific 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   7560
         MultiSelect     =   1  '簡易多重選取
         TabIndex        =   46
         Top             =   600
         Width           =   7215
      End
      Begin Threed.SSCommand cmdExamDetailOK 
         Height          =   435
         Left            =   11160
         TabIndex        =   48
         Top             =   1920
         Width           =   3615
         _Version        =   65536
         _ExtentX        =   6376
         _ExtentY        =   767
         _StockProps     =   78
         Caption         =   "完成"
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand cmdExamDetailCancel 
         Height          =   435
         Left            =   7560
         TabIndex        =   49
         Top             =   1920
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6165
         _ExtentY        =   767
         _StockProps     =   78
         Caption         =   "取消"
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin VB.Label lblExamType 
         BackColor       =   &H00800000&
         BackStyle       =   0  '透明
         Caption         =   "aaa"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   52
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label3 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "檢查細項選取"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2640
         TabIndex        =   51
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label4 
         BackColor       =   &H00800000&
         BackStyle       =   0  '透明
         Caption         =   "病理檢查細項選取"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   7560
         TabIndex        =   50
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404040&
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
         Height          =   2295
         Left            =   120
         TabIndex        =   53
         Top             =   120
         Width           =   14775
      End
   End
   Begin Threed.SSPanel pnlX 
      Height          =   2535
      Left            =   7680
      TabIndex        =   34
      Top             =   7320
      Width           =   7455
      _Version        =   65536
      _ExtentX        =   13150
      _ExtentY        =   4471
      _StockProps     =   15
      ForeColor       =   12582912
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
      Outline         =   -1  'True
      Alignment       =   6
      Begin VB.TextBox txtXDr_Order 
         BackColor       =   &H00404040&
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
         Left            =   1200
         TabIndex        =   97
         Tag             =   "1"
         Top             =   2040
         Width           =   2145
      End
      Begin VB.TextBox txtXDr_on 
         BackColor       =   &H00404040&
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
         Left            =   4560
         TabIndex        =   95
         Tag             =   "1"
         Top             =   960
         Width           =   2745
      End
      Begin VB.TextBox txtXExamDetail 
         BackColor       =   &H00404040&
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
         Height          =   720
         Left            =   4560
         MaxLength       =   255
         TabIndex        =   93
         Tag             =   "1"
         Top             =   1680
         Width           =   2745
      End
      Begin VB.TextBox txtXDr_from 
         BackColor       =   &H00404040&
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
         Left            =   1200
         TabIndex        =   92
         Tag             =   "1"
         Top             =   1680
         Width           =   2145
      End
      Begin VB.TextBox txtXRoom 
         BackColor       =   &H00404040&
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
         Left            =   4560
         TabIndex        =   91
         Tag             =   "1"
         Top             =   600
         Width           =   2745
      End
      Begin VB.TextBox txtXDr_report 
         BackColor       =   &H00404040&
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
         Left            =   4560
         TabIndex        =   90
         Tag             =   "1"
         Top             =   1320
         Width           =   2745
      End
      Begin VB.TextBox txtXType 
         BackColor       =   &H00404040&
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
         Left            =   1200
         TabIndex        =   89
         Tag             =   "1"
         Top             =   600
         Width           =   2145
      End
      Begin VB.TextBox txtUni_key 
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
         Left            =   240
         TabIndex        =   86
         Top             =   0
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFC0C0&
         Caption         =   "更改排檢明細"
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
         Left            =   4560
         Style           =   1  '圖片外觀
         TabIndex        =   81
         Top             =   120
         Width           =   2775
      End
      Begin VB.TextBox txtXOrderDate 
         BackColor       =   &H00404040&
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
         Left            =   1200
         TabIndex        =   54
         Tag             =   "1"
         Top             =   960
         Width           =   2145
      End
      Begin VB.TextBox txtXExamDate 
         BackColor       =   &H00404040&
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
         Left            =   1200
         TabIndex        =   35
         Tag             =   "1"
         Top             =   1320
         Width           =   2145
      End
      Begin VB.TextBox txtClinicalInfo 
         BackColor       =   &H00404040&
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
         Height          =   1800
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   3  '兩者皆有
         TabIndex        =   98
         Tag             =   "1"
         Top             =   600
         Visible         =   0   'False
         Width           =   7185
      End
      Begin VB.Label lblSS7TemplateName 
         BorderStyle     =   1  '單線固定
         Caption         =   "SS7 Template"
         Height          =   375
         Left            =   3840
         TabIndex        =   103
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblEntry 
         BorderStyle     =   1  '單線固定
         Caption         =   "開單醫師"
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
         Index           =   12
         Left            =   120
         TabIndex        =   96
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         BorderStyle     =   1  '單線固定
         Caption         =   "技        師"
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
         Index           =   10
         Left            =   3480
         TabIndex        =   94
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblSS7Template 
         BorderStyle     =   1  '單線固定
         Caption         =   "SS7 Template"
         Height          =   375
         Left            =   2640
         TabIndex        =   74
         Top             =   120
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblEntry 
         Alignment       =   2  '置中對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "檢查細項 "
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   11
         Left            =   3480
         TabIndex        =   43
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         BorderStyle     =   1  '單線固定
         Caption         =   "開單日期"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   13
         Left            =   120
         TabIndex        =   55
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   "檢查日期"
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
         Height          =   345
         Index           =   5
         Left            =   120
         TabIndex        =   39
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   "檢查類別"
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
         Height          =   360
         Index           =   3
         Left            =   120
         TabIndex        =   41
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         BorderStyle     =   1  '單線固定
         Caption         =   "檢查科室"
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
         Index           =   6
         Left            =   3480
         TabIndex        =   38
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         BorderStyle     =   1  '單線固定
         Caption         =   "來源類別"
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
         Index           =   8
         Left            =   120
         TabIndex        =   37
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00800000&
         BorderStyle     =   1  '單線固定
         Caption         =   "排檢內容明細"
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
         Left            =   120
         TabIndex        =   36
         Top             =   120
         Width           =   3255
      End
      Begin VB.Label lblEntry 
         BorderStyle     =   1  '單線固定
         Caption         =   "報告醫師"
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
         Left            =   3480
         TabIndex        =   40
         Top             =   1320
         Width           =   1095
      End
   End
   Begin Threed.SSPanel pnlButtom 
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   9960
      Width           =   15015
      _Version        =   65536
      _ExtentX        =   26485
      _ExtentY        =   1720
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      Outline         =   -1  'True
      Begin Threed.SSCommand cmdAbout 
         Height          =   375
         Left            =   11040
         TabIndex        =   7
         Top             =   525
         Width           =   3855
         _Version        =   65536
         _ExtentX        =   6800
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "ImageSVR Examine Service 9.0"
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         Font3D          =   3
         Outline         =   0   'False
      End
      Begin VB.Label lblRecordCount 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00404040&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   14040
         TabIndex        =   69
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lblEntry 
         BorderStyle     =   1  '單線固定
         Caption         =   "記錄數"
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
         Index           =   17
         Left            =   13200
         TabIndex        =   70
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblUserType 
         BackColor       =   &H00404040&
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   12120
         TabIndex        =   58
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblEntry 
         BorderStyle     =   1  '單線固定
         Caption         =   "用戶類別"
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
         Index           =   16
         Left            =   11040
         TabIndex        =   59
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblUser 
         BackColor       =   &H00404040&
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   8760
         TabIndex        =   57
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label lblEntry 
         BorderStyle     =   1  '單線固定
         Caption         =   "用戶姓名"
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
         Index           =   15
         Left            =   7680
         TabIndex        =   56
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "Registry Records Total"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11040
         TabIndex        =   2
         Top             =   600
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H80000009&
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
         Height          =   375
         Index           =   3
         Left            =   1320
         TabIndex        =   6
         Top             =   525
         Width           =   9615
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H80000009&
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
         Height          =   375
         Index           =   2
         Left            =   1320
         TabIndex        =   5
         Top             =   120
         Width           =   6135
      End
      Begin VB.Label lblStatus 
         BorderStyle     =   1  '單線固定
         Caption         =   "系統訊息"
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
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   525
         Width           =   1215
      End
      Begin VB.Label lblStatus 
         BorderStyle     =   1  '單線固定
         Caption         =   "作業狀態"
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
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   1215
      End
   End
   Begin Threed.SSCommand cmdReview 
      Height          =   375
      Left            =   12480
      TabIndex        =   60
      Top             =   1800
      Visible         =   0   'False
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&R 調閱檢查記錄"
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
   End
   Begin VSFlex7Ctl.VSFlexGrid dbgOnline 
      Height          =   4815
      Left            =   240
      TabIndex        =   67
      Top             =   2280
      Width           =   14775
      _cx             =   26061
      _cy             =   8493
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   16744576
      ForeColorFixed  =   16777215
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmQueue.frx":0442
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   0   'False
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin Threed.SSPanel pnlFilter 
      Height          =   1575
      Left            =   120
      TabIndex        =   8
      Top             =   120
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
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "資料統計"
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
         Left            =   9120
         Style           =   1  '圖片外觀
         TabIndex        =   112
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "開啟排班"
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
         TabIndex        =   111
         Top             =   120
         Width           =   1695
      End
      Begin VB.TextBox txtDate1 
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
         TabIndex        =   107
         Top             =   600
         Width           =   1650
      End
      Begin VB.CommandButton cmdAll 
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
         TabIndex        =   75
         Top             =   1080
         Width           =   1335
      End
      Begin VB.ComboBox cmbPhysician 
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
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtReqNo 
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
         TabIndex        =   102
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFC0C0&
         Caption         =   "刪除檢查報告"
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
         TabIndex        =   88
         Top             =   120
         Width           =   2295
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
         TabIndex        =   87
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
         TabIndex        =   85
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtName 
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
         TabIndex        =   83
         Top             =   1080
         Width           =   1815
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
         TabIndex        =   79
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmdReport 
         BackColor       =   &H00FFC0C0&
         Caption         =   "檢查統計"
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
         Left            =   11160
         Style           =   1  '圖片外觀
         TabIndex        =   78
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdAppend 
         BackColor       =   &H00FFC0C0&
         Caption         =   "新增檢查報告"
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
         TabIndex        =   77
         Top             =   120
         Width           =   2295
      End
      Begin VB.CommandButton cmdOpen 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&O 開啟檢查報告"
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
         TabIndex        =   76
         Top             =   120
         Width           =   2295
      End
      Begin VB.ComboBox cmbDivision 
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
         ItemData        =   "frmQueue.frx":0521
         Left            =   7320
         List            =   "frmQueue.frx":0523
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtChartNo 
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
         TabIndex        =   68
         Top             =   1080
         Width           =   1935
      End
      Begin VB.ComboBox cmbOrder_field 
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
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.ComboBox cmbType1 
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
         ItemData        =   "frmQueue.frx":0525
         Left            =   8040
         List            =   "frmQueue.frx":0527
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox cmbType 
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
         ItemData        =   "frmQueue.frx":0529
         Left            =   10320
         List            =   "frmQueue.frx":052B
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   600
         Width           =   1815
      End
      Begin VB.ComboBox cmbDoctor 
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
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtDate 
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
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   0
         Top             =   600
         Width           =   1530
      End
      Begin VB.ComboBox cmbDr_from 
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
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker dtpDateSort 
         Height          =   360
         Left            =   1200
         TabIndex        =   110
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
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
         DateIsNull      =   -1  'True
         Format          =   70778883
         CurrentDate     =   37058
      End
      Begin MSComCtl2.DTPicker dtpDateSort1 
         Height          =   360
         Left            =   4200
         TabIndex        =   109
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
         Format          =   70778883
         CurrentDate     =   37058
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&O 開啟檢查報告"
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
         TabIndex        =   114
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "起訖日期"
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
         TabIndex        =   106
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "技        師"
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
         TabIndex        =   105
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "申請單號"
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
         TabIndex        =   101
         Top             =   1080
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
         Left            =   6240
         TabIndex        =   84
         Top             =   1080
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
         TabIndex        =   82
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "檢查科室"
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
         TabIndex        =   71
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "排序項目"
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
         Left            =   120
         TabIndex        =   33
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "病歷號"
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
         TabIndex        =   28
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "檢查類別"
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
         TabIndex        =   11
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "起始日期"
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
         TabIndex        =   10
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "報告醫師"
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
         Index           =   4
         Left            =   9240
         TabIndex        =   9
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "來源別"
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
         Index           =   14
         Left            =   9360
         TabIndex        =   100
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.Label lblGrid 
      Alignment       =   2  '置中對齊
      Caption         =   "未報告"
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
      Index           =   6
      Left            =   1440
      TabIndex        =   44
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lblGrid 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00FF8080&
      Caption         =   "已檢查"
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
      Height          =   495
      Index           =   5
      Left            =   240
      TabIndex        =   108
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblGrid 
      Alignment       =   2  '置中對齊
      Caption         =   "待檢查"
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
      Index           =   1
      Left            =   4440
      TabIndex        =   64
      Top             =   1920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblGrid 
      Alignment       =   2  '置中對齊
      Caption         =   "未報到"
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
      Index           =   0
      Left            =   5760
      TabIndex        =   65
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblGrid 
      Alignment       =   2  '置中對齊
      Caption         =   "已檢查"
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
      Index           =   2
      Left            =   1440
      TabIndex        =   63
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblGrid 
      Alignment       =   2  '置中對齊
      Caption         =   "已報告"
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
      Index           =   3
      Left            =   2400
      TabIndex        =   62
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lblGrid 
      Alignment       =   2  '置中對齊
      Caption         =   "全部"
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
      Index           =   4
      Left            =   3360
      TabIndex        =   61
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lblRecordList 
      Appearance      =   0  '平面
      BackColor       =   &H00800000&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   120
      TabIndex        =   42
      Top             =   1800
      Width           =   15015
   End
End
Attribute VB_Name = "frmQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "user32" _
(ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
ByVal x As Long, ByVal Y As Long, ByVal cx As Long, _
ByVal cy As Long, ByVal wFlags As Long) As Long

Const SWP_NOMOVE = &H2 '不更動目前視窗位置
Const SWP_NOSIZE = &H1 '不更動目前視窗大小
Const HWND_TOPMOST = -1 '設定為最上層
Const HWND_NOTOPMOST = -2 '取消最上層設定
Const flags = SWP_NOMOVE Or SWP_NOSIZE

Dim boxDate As Variant
Dim datSource_SQL$
Dim queue_SQL$
Dim queueCaption$

Dim msSortCol As String
Dim mbCtrlKey As Integer
Dim sBookMk As Variant

Dim xChartNo$, xUni_key$, xTable$
Dim nowChartNo$, nowUni_key$, nowType$, nowDate$

Dim aryRecord()
Dim aryResult()

'用於refresh畫面時，暫停一筆一筆更新記錄
Dim tempRef As Boolean

Public is_SRreport As Boolean

Sub cmbDivision_Click()
    Dim filter$, i%
    
    cmbType.Clear
    filter$ = " Divisions='" & cmbDivision.Text & "' "
    'filter$ = " Divisions='心臟內科' "
    Call cmb_Table_Initial(db_Name$, "Cris_ExamType", "Type", filter$, cmbType)
    
    Call cmb_DR_Initial(db_Name$, "CRIS_User", "Name", "System='" & cmbDivision & "'", cmbDoctor)
    cmbDoctor.AddItem "全部"
    cmbPhysician.Clear
    For i% = 0 To cmbDoctor.ListCount - 1: cmbPhysician.AddItem cmbDoctor.List(i%): Next
    
End Sub

Private Sub cmbOrder_field_Click()
    DoEvents
    
    Call dat_Refresh
    
End Sub

Private Sub cmdAll_Click()
    dbgOnline.Enabled = False
    cmdOpen.Visible = False
    Call dat_Refresh
    dbgOnline.Enabled = True
    cmdOpen.Visible = True
End Sub

Private Sub cmdAppend_Click()
    Dim frmDetail As Object
    
    'global_currOption = "新增模式"
    
    If Not (lblUserType = "技術師" Or lblUserType = "醫師" Or lblUserType = "管理員") Then
        MsgBox "權限不符"
        Exit Sub
    End If
    
    Load frmAddNew
    'frmCpInfo.lblMode.Caption = "新增模式"
    DoEvents
    
    frmAddNew.Show
    Call frmAddNew.setMode("Insert")
    Call frmAddNew.cmdManully_Click
    
End Sub

Private Sub cmdBasicEdit_Click()
    Dim frmDetail As Object
    
'    If Not (lblUserType = "技術師" Or lblUserType = "醫師") Then
'        MsgBox "權限不符"
'        Exit Sub
'    End If
    
    
    If curr_Record.chartno = "" Then Exit Sub
    
'     With curr_Record
 '       .Uni_key = NoNull(adoOnline.Recordset("Uni_key"))
 '       .ChartNo = NoNull(adoOnline.Recordset("ChartNo"))
 '       .Date = NoNull(adoOnline.Recordset("Date"))
 '       .Type = NoNull(adoOnline.Recordset("Type"))
 '       .Room = NoNull(adoOnline.Recordset("Room"))
 '       .Age = NoNull(adoOnline.Recordset("Age"))
 '       .Item1 = NoNull(adoOnline.Recordset("Item1"))
 '       .Item2 = NoNull(adoOnline.Recordset("Item2"))
 '       .Item3 = NoNull(adoOnline.Recordset("Item3"))
 '       .Item4 = NoNull(adoOnline.Recordset("Item4"))
 '       .Item5 = NoNull(adoOnline.Recordset("Item5"))
 '       .Item6 = NoNull(adoOnline.Recordset("Item6"))
 '       .Others = NoNull(adoOnline.Recordset("Others"))
 '       .Dr_from = NoNull(adoOnline.Recordset("Dr_from"))
 '       .Dr_on = NoNull(adoOnline.Recordset("Dr_on"))
 '       .Status = NoNull(adoOnline.Recordset("Status"))
 '       .Class = NoNull(adoOnline.Recordset("Class"))
 '       .ImgPicked = NoNull(adoOnline.Recordset("ImgPicked"))
 '   End With
    
    Load frmUpdate
    frmUpdate.Show

End Sub

Private Sub cmdClear_Click()
    
    cmbType = ""
    cmbDoctor = ""
    cmbDivision = ""
    cmbDr_from = ""
    cmbPhysician = ""
    
    txtDate = ""
    txtDate1 = ""
    txtAccessionNo = ""
    txtChartNo = ""
    txtName = ""
    txtStatus = ""
    txtReqNo = ""
    
End Sub

Sub cmdClose_Click()
    Dim tmpFileName$
    
    '清除 tmp 目錄下的html檔案
    tmpFileName$ = Dir(path_System & "Tmp\*.html", vbNormal)
    Do While Len(tmpFileName$) > 0
          Kill path_System & "Tmp\" & tmpFileName$
          tmpFileName$ = Dir(path_System & "Tmp\*.html", vbNormal)
    Loop
    
    Unload Me
    
End Sub


Private Sub cmdDelete_Click()
    Dim tmpStud_No As String, SQL$
    Dim dbS As New adoDB.Connection
    Dim dbT As New adoDB.Recordset
    
    
    If Not (lblUserType = "技術師" Or lblUserType = "醫師" Or lblUserType = "管理員") Then
        MsgBox "權限不符"
        Exit Sub
    End If
    
    
    'If dbgMain.SelBookmarks.Count <> 0 Then
'    If adoOnline.Recordset.EOF Then
'        MsgBox "請先選擇一筆記錄."
'        Exit Sub
'    End If
    
    
'    If adoOnline.Recordset.EOF Then Exit Sub
'     With curr_Record
'        .Uni_key = NoNull(adoOnline.Recordset("Uni_key"))
'        .ChartNo = NoNull(adoOnline.Recordset("ChartNo"))
'        .Date = NoNull(adoOnline.Recordset("Date"))
'        .Type = NoNull(adoOnline.Recordset("Type"))
'        .Room = NoNull(adoOnline.Recordset("Room"))
'        .Age = NoNull(adoOnline.Recordset("Age"))
'        .Item1 = NoNull(adoOnline.Recordset("Item1"))
'        .Item2 = NoNull(adoOnline.Recordset("Item2"))
'        .Item3 = NoNull(adoOnline.Recordset("Item3"))
'        .Item4 = NoNull(adoOnline.Recordset("Item4"))
'        .Item5 = NoNull(adoOnline.Recordset("Item5"))
'        .Item6 = NoNull(adoOnline.Recordset("Item6"))
'        .Others = NoNull(adoOnline.Recordset("Others"))
'        .Dr_from = NoNull(adoOnline.Recordset("Dr_from"))
'        .Dr_on = NoNull(adoOnline.Recordset("Dr_on"))
'        .Status = NoNull(adoOnline.Recordset("Status"))
'        .Class = NoNull(adoOnline.Recordset("Class"))
'        .ImgPicked = NoNull(adoOnline.Recordset("ImgPicked"))
'        .Birthday = NoNull(adoOnline.Recordset("Birthday"))
'        .Time = NoNull(adoOnline.Recordset!Time)
'    End With
    
    Msg = "確定刪除病歷號：" & curr_Record.chartno & "，檢查別：" & curr_Record.Type & "，日期：" & curr_Record.Date & "，排檢編號：" & curr_Record.uni_key & " 的排檢記錄？"
    style = vbYesNo + vbCritical + vbDefaultButton2
    Title = "<注意>"
    Response = MsgBox(Msg, style, Title)
    If Response = vbYes Then
    
       dbS.Open dbConnection$

       SQL$ = "UPDATE CRIS_Exam_Online SET Status='已刪除' WHERE ChartNo='" & curr_Record.chartno & "' AND Type='" & curr_Record.Type & "' AND ExamDate='" & curr_Record.Date & "' AND Uni_key='" & curr_Record.uni_key & "'" ' AND ExamTime='" & NoNull(curr_Record.Time) & "'"
       Call DBRecordLog("delete", SQL$, "刪除記錄，更新cris_exam_online")
       dbS.Execute SQL$
       dbS.Close
       
       Set dbS = Nothing
'       Call Raw_refresh("")

    End If
    
    
    Exit Sub
    
RefErr:
    MsgBox "Error:" & err & " " & err.Description
    Resume Next

End Sub

Private Sub cmdOpen_Click()
    If Left(UCase(Trim(curr_Record.Dr_from)), 6) = "HIS_IN" And curr_Record.Division_on <> "先天性心臟病科" Then
        Msg = "此筆記錄尚未取得HIS資料!" & vbCrLf & "可能尚未報到，請先報到後再編輯，已避免報告記錄被覆蓋!" & vbCrLf & "是否先停止，等報到後再開啟報告編輯？"
        style = vbYesNo + vbCritical + vbDefaultButton2
        Title = "<注意>"
        Response = MsgBox(Msg, style, Title)
        If Response = vbYes Then
            Exit Sub
        End If
    End If
    If cmdOpen.Enabled = False Then
        Exit Sub
    End If
    '/*記錄一些要用到的使用者資料，並把醫生的號碼及姓名做分割，再去找該醫師是否有手機號碼，並記錄下來，發簡訊時會用到*/
    Login_LastOpenReportDrFrom = txtXDr_from.Text
    Login_LastOpenReportUnikey = txtUni_key.Text
    Login_LastOpenReportModifyDate = txtXExamDate.Text
'    Login_LastOpenReportChartNO = dbgOnline.TextMatrix(dbgOnline.RowSel, 1)
    Login_LastOpenReportChartNO = dbgOnline.TextMatrix(dbgOnline.RowSel, 2)
'    Login_LastOpenReportChartName = dbgOnline.TextMatrix(dbgOnline.RowSel, 2)
    Login_LastOpenReportChartName = dbgOnline.TextMatrix(dbgOnline.RowSel, 3)
    Login_LastOpenReportType = txtXType.Text
    Login_LastOpenReportStatus = curr_Record.Status
    
    
    Login_LastOpenReportModifyDate = curr_Record.OrderDate
    Login_LastOpenReportModifyTime = curr_Record.OrderTime
    Login_LastOpenReportReturnDate = curr_Record.ReportDate
    Login_LastOpenReportReturnTime = curr_Record.ReportTime
    
    
    Login_LastOpenReportCreateDRName = txtXDr_Order.Text
    
    Dim SQL_String As String
    Dim num As String, str As String
    Call Str_Classify(Login_LastOpenReportCreateDRName, num, str, Asc("0"), Asc("9"))

    'SQL_String = "select system,phone from cris_user where userid='" & num & "' "
    SQL_String = "select system,phone from cris_user where userid='" & num & "' "
    
    Call OpenRecordset(SQL_String, Connection, Recordset)
    If Not Recordset.EOF Then
        If IsNull(Recordset("phone")) Then
            Login_LastOpenReportCreateDRPhone = ""
        Else
            Login_LastOpenReportCreateDRPhone = Recordset("phone")
        End If
        
        If IsNull(Recordset("system")) Then
            Login_LastOpenReportSystem = ""
        Else
            Login_LastOpenReportSystem = Recordset("system")
        End If
    End If
    '/**/



    Dim frmDetail As Object
    Dim xUni_key$, xReportTemplateFileName$
    Dim tmpTemplate$, nameTemplate$, i%, ret As Long
    Dim tmpAPP$, apiCall&, temp$
    Dim SQL$
    Dim dbS As New adoDB.Connection
    
    'for 高階健檢專用----------------------------------------
    If Trim(curr_Record.Type) = "PET" Then
        
        tmpAPP$ = "HCR.EXE " & curr_Record.uni_key
        apiCall& = Shell(tmpAPP$, vbNormalFocus)
       
       Exit Sub
    End If
    '--------------------------------------------------------
    
    If curr_Record.Status = "未報到" Then
       MsgBox "本筆檢查尚未報到，必須先變更排檢後再編輯報告內容！"
       Exit Sub
    End If
    
    '如果是HP檢查時
    If UCase(Trim(txtXType.Text)) = "HP檢查" Then
        If Trim(curr_Record.Dr_order) = "" Then
            MsgBox "無開單醫師資料，請先填寫後才可繼續作業!"
        Else
            Response = MsgBox("請選擇陽性(是) 或 陰性(否) 或 Cancel(取消)", vbYesNoCancel, "HP檢查")
            If Response <> vbCancel Then
                If Response = vbYes Then
                    temp$ = "陽性"
                Else
                    temp$ = "陰性"
                End If
                dbS.Open dbConnection$
                SQL$ = "update cris_exam_online set item6 = '" & temp$ & "', "
                SQL$ = SQL$ & "AccessionNumber='" & curr_Record.uni_key & "', "
                SQL$ = SQL$ & "LastUpdateDate='" & Format(Date, "yyyy/MM/dd") & "', "
                SQL$ = SQL$ & "LastUpdateTime='" & Format(time, "hh:NN:ss") & "', "
                SQL$ = SQL$ & "LastUpdateUser='" & frmQueue.lblUser & "', "
                SQL$ = SQL$ & "HISup='50', "
                SQL$ = SQL$ & "Class='UPLOAD', "
                If curr_Record.Status <> "已報告" Then
                    SQL$ = SQL$ & "SigninSerial='" & FindSigninSerial & "', "
                    SQL$ = SQL$ & "ReportDate='" & Format(Date, "yyyy/MM/dd") & "', "
                    SQL$ = SQL$ & "ReportTime='" & Format(time, "hh:NN:ss") & "', "
                    SQL$ = SQL$ & "Dr_on='" & curr_Record.Dr_order & "', "
                    SQL$ = SQL$ & "Dr_report='" & curr_Record.Dr_order & "', "
                    SQL$ = SQL$ & "follow_dr='" & curr_Record.Dr_order & "', "
                    SQL$ = SQL$ & "Status='已報告', "
                End If
                SQL$ = SQL$ & "UPLOADCODE='30' "
                SQL$ = SQL$ & "WHERE status<>'已刪除' and ChartNo='" & curr_Record.chartno & "' AND Type='" & curr_Record.Type & "' AND Uni_key='" & curr_Record.uni_key & "'"
                Call DBRecordLog("update", SQL$, "上傳cris_exam_online")
                dbS.Execute SQL$
                dbS.Close
                
                Set dbS = Nothing
                Call dat_Refresh
            End If
        End If
        Exit Sub
    End If
    
    dbgOnline.Enabled = False
'    DoEvents
    
    lblSS7Template = ""
    tmpTemplate$ = ""
    If Not curr_Record.chartno = "" Then
'        If curr_Record.Type = ReportName Then
'                 'Text Report 肺功能檢查　--------------------------------------
'                 Call array_DictionByType_Initial(curr_Record.Type)
'                 lblSS7Template = "CustomerReport"
'                 Load frmCustomerReport
'                 DoEvents
'                 frmCustomerReport.Show
'        Else
             For i% = 0 To UBound(xReportTemplate)
                  '因同一科室會有多筆檢查故應為每種檢查預設表單
                  'If xReportTemplate(i%).DivisionName = curr_Record.Division_on And xReportTemplate(i%).DefaultUse = "Y" Then
                  If xReportTemplate(i%).DivisionName = curr_Record.Division_on And _
                     xReportTemplate(i%).ExamName = curr_Record.Type And _
                     xReportTemplate(i%).DefaultUse = "Y" Then
                     
                     '/**/
                     Spread_ID = xReportTemplate(i%).ExamID
                     Spread_Name = xReportTemplate(i%).ExamName
                     '/**/
                     
                     tmpTemplate$ = xReportTemplate(i%).TemplateFileSource & xReportTemplate(i%).TemplateFileName
                     nameTemplate$ = xReportTemplate(i%).ExamDescription
                     Exit For
                  End If
             Next
             
             '若沒找到科室內檢查類別相符的報表時，尋找科室報表
             If tmpTemplate$ = "" Then
                 For i% = 0 To UBound(xReportTemplate)
                      If xReportTemplate(i%).DivisionName = curr_Record.Division_on Then
                         tmpTemplate$ = xReportTemplate(i%).TemplateFileSource & xReportTemplate(i%).TemplateFileName
                         nameTemplate$ = xReportTemplate(i%).ExamDescription
                         
                         Spread_ID = xReportTemplate(i%).ExamID
                         Spread_Name = xReportTemplate(i%).ExamName
                         Exit For
                      End If
                 Next
             End If
             
             '找不到科室內預設報表時，尋找檢查類別報表
             If tmpTemplate$ = "" Then
                 For i% = 0 To UBound(xReportTemplate)
                      If xReportTemplate(i%).ExamName = curr_Record.Type Then
                         tmpTemplate$ = xReportTemplate(i%).TemplateFileSource & xReportTemplate(i%).TemplateFileName
                         nameTemplate$ = xReportTemplate(i%).ExamDescription
                         
                         Spread_ID = xReportTemplate(i%).ExamID
                         Spread_Name = xReportTemplate(i%).ExamName
                         Exit For
                      End If
                 Next
             End If
             
             If Not tmpTemplate$ = "" Then
                 'Spread 表單樣式報告模組　---------------------------------------
                 lblSS7Template = Trim(tmpTemplate$)
                 lblSS7TemplateName = Trim(nameTemplate$)
                 If xSpread2$ = "" Then
                    xSpread2$ = "肝膽腸胃內科"
                 End If
'                 If curr_Record.Division_on = xSpread2$ Or curr_Record.Division_on = "大腸直腸外科" Then
                If curr_Record.Division_on = "肝膽腸胃內科" Or curr_Record.Division_on = "大腸直腸外科" Then
                    Load frmSpread2
                    DoEvents
                    frmSpread2.Show
                Else
                    Load frmSpread
                    DoEvents
                    frmSpread.Show
                End If
                lblStatus(2) = ""
                lblStatus(3) = ""
                
                
                Me.Visible = False
                Me.Enabled = False
             Else
                lblStatus(2) = "開啟報告錯誤，請再確認檢查類別與科室別是否正確"
'                 'Text Report 樣式報告模組　--------------------------------------
'                 Call array_DictionByType_Initial(curr_Record.Type)
'                 lblSS7Template = "None"
'                 Load frmReport
'                 DoEvents
'                 frmReport.Show
             
             End If
'        End If
'        lblStatus(2) = ""
'        lblStatus(3) = ""
'
'
'        Me.Visible = False
'        Me.Enabled = False
    Else
       lblStatus(2) = "排檢資料有誤, 請通知系統管理人員!"
    End If
    
    dbgOnline.Enabled = True
    
End Sub

Private Sub cmdReport_Click()
'    Dim tmp$, i&
'
'    dbgOnline.PrintGrid , True, 2, 720, 720
    Shell Report_Name$, vbNormalFocus
End Sub

Private Sub cmdReview_Click()
    Dim xChartNo$, i&, tmp$
        
    On Error GoTo Review_error
    
    If Not (lblUserType = "技術師" Or lblUserType = "醫師" Or lblUserType = "保業室") Then
        MsgBox "權限不符"
        Exit Sub
    End If
    
    xChartNo$ = InputBox("請輸入病歷號碼", "調閱")
    If Len(xChartNo$) <= 6 Then
       xChartNo$ = String(7 - Len(Trim(xChartNo$)), "0") & Trim(xChartNo$)
    End If
    
    If xChartNo$ > "" Then
        tmp$ = App.Path & "\CRISViewer.exe " & xChartNo$
        i& = Shell(tmp$, vbNormalFocus)
        DoEvents
    End If
    On Error GoTo 0
    Exit Sub
    
Review_error:
    If err = 53 Then
       MsgBox App.Path & "\CRISViewer.exe 並不存在，無法實行調閱"
    Else
       MsgBox Error(err)
    End If
    
End Sub

Private Sub cmdSync_Click()
    
    Call dat_Refresh

End Sub

Private Sub cmdUpdate_Click()
    Dim dbS As New adoDB.Connection
    Dim dbT As New Recordset
    Dim conn$, SQL$, currRow&, i%
    
    currRow& = dbgOnline.row
'    If (Not lblUser = adoOnline.Recordset!Dr_on) And NoNull(adoOnline.Recordset!Dr_on) > "" Then 'curr_Record.Dr_on Then 'cmbXDoctor Then
'        If Not lblUserType = "技術師" Then
'            MsgBox "權限不符" '，必須由報到櫃檯變更檢查醫師"
'            Exit Sub
'        End If
'    End If
    
'    If (curr_Record.Status = "已報告" And (lblUser <> curr_Record.Dr_report Or curr_Record.Dr_report = "")) Then
''    If (curr_Record.Status = "已報告" And lblUser <> curr_Record.Dr_report) Then
'        MsgBox "已報告需報告醫師更改"
'        Exit Sub
'    End If
    
    Load frmAddNew
    'frmCpInfo.lblMode.Caption = "新增模式"
    DoEvents
    
    frmAddNew.Show
    
    frmAddNew.txtBasic(0) = curr_Record.chartno
    Call frmAddNew.cmdGet_Click
    frmAddNew.txtAccessionNo = curr_Record.uni_key
    frmAddNew.dtpDate = IIf(IsDate(curr_Record.Date), curr_Record.Date, Date)
    frmAddNew.dtpTime = IIf(Len(curr_Record.time) > 4, curr_Record.time, Format(time, "hh:NN:ss"))
    frmAddNew.dtpTime = IIf(Len(curr_Record.time) = 8, curr_Record.time, Format(time, "hh:NN:ss"))
    frmAddNew.dtpOrderDate = IIf(IsDate(curr_Record.OrderDate), curr_Record.OrderDate, Date)
    frmAddNew.dtpOrderTime = IIf(Len(curr_Record.OrderTime) = 8, curr_Record.OrderTime, Format(time, "hh:NN:ss"))
    
    frmAddNew.txtOrderBy = curr_Record.Dr_order
    frmAddNew.cmbDr_from = IIf(Len(curr_Record.Dr_from) > 1, curr_Record.Dr_from, "門診")
    frmAddNew.cmbRoom = IIf(Len(curr_Record.Room) > 1, curr_Record.Room, RoomName$)
    
    frmAddNew.cmbSystem = IIf(Len(curr_Record.Division_on) > 1, curr_Record.Division_on, UserDivision$)
    Call frmAddNew.cmbSystem_Click
    
    If curr_Record.Sex = "男" Then
        frmAddNew.optSex(0).Value = True
    ElseIf curr_Record.Sex = "女" Then
        frmAddNew.optSex(1).Value = True
    End If
    
    frmAddNew.cmbDr_on = IIf(Len(curr_Record.Dr_on) > 1, curr_Record.Dr_on, UserName$)
    frmAddNew.cmbDr_Report = IIf(Len(curr_Record.Dr_report) > 1, curr_Record.Dr_report, UserName$)
    frmAddNew.cmbOrderBy = IIf(Len(curr_Record.Dr_order) > 1, curr_Record.Dr_order, UserName$)
    
    frmAddNew.cmbChargeBy = IIf(Len(curr_Record.ChargeBy) > 1, curr_Record.ChargeBy, "健保")
    
    frmAddNew.cmbType = curr_Record.Type
    Call frmAddNew.cmbType_Click
    
    frmAddNew.cmbStatus = IIf(curr_Record.Status = "未報到", "待檢查", curr_Record.Status)
    
    frmAddNew.txtExamDetail = curr_Record.ExamDetail
'    For i% = 0 To frmAddNew.lstExamDetail.ListCount - 1
'        If InStr(curr_Record.ExamDetail, Trim(left(frmAddNew.lstExamDetail.List(i%), 10))) > 0 Then
'           frmAddNew.lstExamDetail.Selected(i%) = True
'        End If
'    Next
    Call chkExamDetail(frmAddNew.lstExamDetail, frmAddNew.txtExamDetail)
    
    Call frmAddNew.setMode("Edit")

End Sub


Sub UpdateItem_Mode(xEnabled%)
    Dim i%
    
    With Me
        For i% = 0 To .Controls.Count - 1
          If (.Controls(i%).Tag = "1") Then
             If (TypeOf .Controls(i%) Is TextBox) Then .Controls(i%).Enabled = xEnabled%
             If (TypeOf .Controls(i%) Is ComboBox) Then .Controls(i%).Enabled = xEnabled%
             
             If (TypeOf .Controls(i%) Is ListBox) Then .Controls(i%).Enabled = xEnabled%
             If (TypeOf .Controls(i%) Is SSCommand) Then .Controls(i%).Enabled = xEnabled%
             If (TypeOf .Controls(i%) Is DTPicker) Then .Controls(i%).Enabled = xEnabled%
          End If
        Next
    End With
    
End Sub


Private Sub Command1_Click()
    Shell Report_Name1$, vbNormalFocus
End Sub

Private Sub Command2_Click()
'    Shell Report_Name2$, vbNormalFocus
is_SRreport = True
frmQueue.Enabled = False
frmSRQuery.Show
'SetWindowPos frmSRQuery.hWnd, HWND_TOPMOST, 0, 0, 0, 0, flags

End Sub

Private Sub dbgOnline_DblClick()
    
    'If lblRecordCount < 1 Or dbgOnline.row > Val(lblRecordCount) Then Exit Sub
    
    'nowChartNo$ = dbgOnline.Cell(flexcpText, dbgOnline.row, 1, dbgOnline.row, 1) ' NoNull(aryResult(1, xRow&))
    'nowUni_key$ = dbgOnline.Cell(flexcpText, dbgOnline.row, 0, dbgOnline.row, 0) ' NoNull(aryResult(0, xRow&))
    'nowType$ = dbgOnline.Cell(flexcpText, dbgOnline.row, 5, dbgOnline.row, 5)  'NoNull(aryResult(5, xRow&))
    'nowDate$ = dbgOnline.Cell(flexcpText, dbgOnline.row, 6, dbgOnline.row, 6)  'NoNull(aryResult(6, xRow&))
    
    Call cmdOpen_Click

End Sub

Sub dbgOnline_SelChange()
    
    If Val(lblRecordCount) < 1 Or dbgOnline.row > Val(lblRecordCount) Then Exit Sub
    Me.Enabled = False
'    nowChartNo$ = dbgOnline.Cell(flexcpText, dbgOnline.row, 1, dbgOnline.row, 1) ' NoNull(aryResult(1, xRow&))
    nowChartNo$ = dbgOnline.Cell(flexcpText, dbgOnline.row, 2, dbgOnline.row, 2)
'    nowUni_key$ = dbgOnline.Cell(flexcpText, dbgOnline.row, 0, dbgOnline.row, 0) ' NoNull(aryResult(0, xRow&))
    nowUni_key$ = dbgOnline.Cell(flexcpText, dbgOnline.row, 1, dbgOnline.row, 1)
'    nowType$ = dbgOnline.Cell(flexcpText, dbgOnline.row, 5, dbgOnline.row, 5)  'NoNull(aryResult(5, xRow&))
    nowType$ = dbgOnline.Cell(flexcpText, dbgOnline.row, 6, dbgOnline.row, 6)
'    nowDate$ = dbgOnline.Cell(flexcpText, dbgOnline.row, 6, dbgOnline.row, 6)  'NoNull(aryResult(6, xRow&))
    nowDate$ = dbgOnline.Cell(flexcpText, dbgOnline.row, 7, dbgOnline.row, 7)
    If Not tempRef Then
        Call currRecord_Refresh(nowChartNo$, nowDate$, nowType$, nowUni_key$)
    End If
    lblStatus(2) = ""
    Me.Enabled = True
End Sub

Sub currRecord_Refresh(aChartNo$, aDate$, aType$, aUni_key$)
    Dim dbS As New adoDB.Connection
    Dim dbTargetRS As New adoDB.Recordset
    Dim xrow&, datWhere$, datSourc_SQL$
    Dim tmp$, tmp1$
    Dim i As Integer
    
    dbS.Open dbConnection$
'    datWhere$ = " WHERE CRIS_Exam_Online.status<>'已刪除' and CRIS_Exam_Online.ChartNo='" & aChartNo$ & "' AND Type='" & aType$ & "' AND Uni_key='" & aUni_key$ & "' "
    datWhere$ = " WHERE CRIS_Exam_Online.status<>'已刪除' and CRIS_Exam_Online.ChartNo='" & aChartNo$ & "' AND Uni_key='" & aUni_key$ & "' "
    'adoOnline.ConnectionString = dbConnection$
    'adoOnline.Recordset.Close
    'adoOnline.RecordSource = datSource_SQL$ & datWhere$
    
    
    dbTargetRS.Open datSource_SQL$ & datWhere$, dbS, adOpenForwardOnly, adLockReadOnly
    If dbTargetRS.BOF Or dbTargetRS.EOF Then
       GoSub Record_Absent
       GoSub Record_Empty
    Else
       GoSub Record_Exist
       GoSub Record_Assign
    End If
    dbTargetRS.Close
    Set dbTargetRS = Nothing
    
    dbS.Close
    Set dbS = Nothing
    
'    DoEvents
    
    
    xrow& = dbgOnline.row - 1
    
    Dim xEnabled%
    
    xEnabled% = False
    'Call UpdateItem_Mode(xEnabled%)
    
'    cmdUpdateEnable.Visible = True
'    cmdUpdateEnable.Enabled = True
'    cmdUpdateEnable.ZOrder 0
    Call View_Only
    Exit Sub
    
Record_Absent:
'    cmdUpdate.Enabled = False
    cmdBasicEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdOpen.Enabled = False
         
    txtUni_key = ""

    txtXType = ""
    txtXOrderDate = ""
    txtXExamDate = ""
    txtXDr_from = ""
    txtXDr_Order = ""
    txtXDr_report = ""
    
    txtXRoom = ""
    txtXDr_on = ""
    txtXExamDetail = ""
    txtClinicalInfo = ""
    
    lblInfo(1) = ""
    lblInfo(2) = ""
    lblInfo(3) = ""
    lblInfo(4) = ""
    lblInfo(5) = ""
    lblInfo(6) = ""
    
    Return


Record_Exist:
'    cmdUpdate.Enabled = True
    cmdBasicEdit.Enabled = True
    cmdDelete.Enabled = True
    cmdOpen.Enabled = True
         
    txtUni_key = NoNull(dbTargetRS!uni_key)

    txtXType = NoNull(dbTargetRS!Type)
    txtXOrderDate = NoNull(dbTargetRS!OrderDate)
    txtXExamDate = NoNull(dbTargetRS!examdate)
    txtXDr_from = NoNull(dbTargetRS!Dr_from)
    txtXDr_Order = NoNull(dbTargetRS!Dr_order)
    
    txtXRoom = NoNull(dbTargetRS!Division_on)
    txtXDr_on = NoNull(dbTargetRS!Dr_on)
    txtXDr_report = NoNull(dbTargetRS!Dr_report)
    txtXExamDetail = NoNull(dbTargetRS!ExamDetail)
    txtClinicalInfo = NoNull(dbTargetRS!ClinicalImp)
     
    lblInfo(0) = NoNull(dbTargetRS!chartno)
    lblInfo(1) = NoNull(dbTargetRS!Name)
    lblInfo(2) = NoNull(dbTargetRS!BirthDay)
    lblInfo(3) = NoNull(dbTargetRS!Sex)
    lblInfo(4) = NoNull(dbTargetRS!Address)
    lblInfo(5) = NoNull(dbTargetRS!Phone)
    lblInfo(6) = NoNull(dbTargetRS!CitizenID)
    
    Return
    
Record_Assign:
       With curr_Record
            .System = NoNull(dbTargetRS!System)
            
            .uni_key = NoNull(dbTargetRS!uni_key)
            .chartno = NoNull(dbTargetRS!chartno)
            
            .Name = NoNull(dbTargetRS!Name)
            .Sex = NoNull(dbTargetRS!Sex)
            .BirthDay = NoNull(dbTargetRS!BirthDay)
            .Phone = NoNull(dbTargetRS!Phone)
            .Address = NoNull(dbTargetRS!Address)
            .CitizenID = NoNull(dbTargetRS!CitizenID)
            
            .Date = NoNull(dbTargetRS!examdate)
            .Type = NoNull(dbTargetRS!Type)
            .Room = NoNull(dbTargetRS!Room)
            .Age = NoNull(dbTargetRS!Age)
            .Item1 = NoNull(dbTargetRS!Item1)
            .Item2 = NoNull(dbTargetRS!Item2)
            .Item3 = NoNull(dbTargetRS!Item3)
            .Item4 = NoNull(dbTargetRS!Item4)
            .Item5 = NoNull(dbTargetRS!Item5)
            .Item6 = NoNull(dbTargetRS!Item6)
            .Others = NoNull(dbTargetRS!Others)
            .Dr_from = NoNull(dbTargetRS!Dr_from)
            .Dr_on = NoNull(dbTargetRS!Dr_on)
            .Dr_order = NoNull(dbTargetRS!Dr_order)
            .Dr_report = NoNull(dbTargetRS!Dr_report)
            .Status = NoNull(dbTargetRS!Status)
            .UploadCode = NoNull(dbTargetRS!UploadCode)
            
            .Class = NoNull(dbTargetRS!Class)
            .ImgPicked = NoNull(dbTargetRS!ImgPicked)
            .time = NoNull(dbTargetRS!examtime)
            .Modality = NoNull(dbTargetRS!Modality)
            .Reg_Date = NoNull(dbTargetRS!Reg_Date)
            .ExamDetail = NoNull(dbTargetRS!ExamDetail)
            
            .OrderDate = NoNull(dbTargetRS!OrderDate)
            .OrderTime = NoNull(dbTargetRS!OrderTime)
            .ReportDate = NoNull(dbTargetRS!ReportDate)
            .ReportTime = NoNull(dbTargetRS!ReportTime)
            .LastUpdateDate = NoNull(dbTargetRS!LastUpdateDate)
            .LastUpdateTime = NoNull(dbTargetRS!LastUpdateTime)
            
            .Division_from = NoNull(dbTargetRS!Division_from)
            .Division_on = NoNull(dbTargetRS!Division_on)
            
            .Division_Seq = NoNull(dbTargetRS!Division_Seq)
            .ClinicalImp = NoNull(dbTargetRS!ClinicalImp)
            .TemplateName = NoNull(dbTargetRS!TemplateName)
            .TemplateFile = NoNull(dbTargetRS!TemplateFile)
            
            .ChargeBy = NoNull(dbTargetRS!ChargeBy)
            
            .HIS_ReqNo = NoNull(dbTargetRS!HIS_ReqNo)
            For i = 0 To dbTargetRS.Fields.Count - 1
                If UCase(Trim(dbTargetRS.Fields(i).Name)) = "FOLLOW_DR" Then
                    .Dr_follow = NoNull(dbTargetRS!Follow_Dr)
                End If
            Next
       End With
       
       With save_Record
            .uni_key = NoNull(dbTargetRS!uni_key)
            .chartno = NoNull(dbTargetRS!chartno)
            
            .Name = NoNull(dbTargetRS!Name)
            .Sex = NoNull(dbTargetRS!Sex)
            .BirthDay = NoNull(dbTargetRS!BirthDay)
            .Phone = NoNull(dbTargetRS!Phone)
            .Address = NoNull(dbTargetRS!Address)
            .CitizenID = NoNull(dbTargetRS!CitizenID)
            
            .Date = NoNull(dbTargetRS!examdate)
            .Type = NoNull(dbTargetRS!Type)
            .Room = NoNull(dbTargetRS!Room)
            .Age = NoNull(dbTargetRS!Age)
            .Item1 = NoNull(dbTargetRS!Item1)
            .Item2 = NoNull(dbTargetRS!Item2)
            .Item3 = NoNull(dbTargetRS!Item3)
            .Item4 = NoNull(dbTargetRS!Item4)
            .Item5 = NoNull(dbTargetRS!Item5)
            .Item6 = NoNull(dbTargetRS!Item6)
            .Others = NoNull(dbTargetRS!Others)
            .Dr_from = NoNull(dbTargetRS!Dr_from)
            .Dr_on = NoNull(dbTargetRS!Dr_on)
            .Dr_order = NoNull(dbTargetRS!Dr_order)
            .Dr_report = NoNull(dbTargetRS!Dr_report)
            .Status = NoNull(dbTargetRS!Status)
            .Class = NoNull(dbTargetRS!Class)
            .UploadCode = NoNull(dbTargetRS!UploadCode)
            
            .ImgPicked = NoNull(dbTargetRS!ImgPicked)
            .time = NoNull(dbTargetRS!examtime)
            .Modality = NoNull(dbTargetRS!Modality)
            .Reg_Date = NoNull(dbTargetRS!Reg_Date)
            .ExamDetail = NoNull(dbTargetRS!ExamDetail)
            
            .OrderDate = NoNull(dbTargetRS!OrderDate)
            .OrderTime = NoNull(dbTargetRS!OrderTime)
            .ReportDate = NoNull(dbTargetRS!ReportDate)
            .ReportTime = NoNull(dbTargetRS!ReportTime)
            .LastUpdateDate = NoNull(dbTargetRS!LastUpdateDate)
            .LastUpdateTime = NoNull(dbTargetRS!LastUpdateTime)
            
            .Division_from = NoNull(dbTargetRS!Division_from)
            .Division_on = NoNull(dbTargetRS!Division_on)
            
            .Division_Seq = NoNull(dbTargetRS!Division_Seq)
            .ClinicalImp = NoNull(dbTargetRS!ClinicalImp)
            .TemplateName = NoNull(dbTargetRS!TemplateName)
            .TemplateFile = NoNull(dbTargetRS!TemplateFile)
       
            .ChargeBy = NoNull(dbTargetRS!ChargeBy)
            
            .HIS_ReqNo = NoNull(dbTargetRS!HIS_ReqNo)
            For i = 0 To dbTargetRS.Fields.Count - 1
                If UCase(Trim(dbTargetRS.Fields(i).Name)) = "FOLLOW_DR" Then
                    .Dr_follow = NoNull(dbTargetRS!Follow_Dr)
                End If
            Next
       
       End With
    
    Return
    
Record_Empty:
       With curr_Record
            .uni_key = ""
            .chartno = ""
            
            .Name = ""
            .Sex = ""
            .BirthDay = ""
            .Phone = ""
            .Address = ""
            .CitizenID = ""
            .Date = ""
            .Type = ""
            .Room = ""
            .Age = ""
            .Item1 = ""
            .Item2 = ""
            .Item3 = ""
            .Item4 = ""
            .Item5 = ""
            .Item6 = ""
            .Others = ""
            .Dr_from = ""
            .Dr_on = ""
            .Dr_order = ""
            .Dr_report = ""
            .Status = ""
            .Class = ""
            .UploadCode = ""
            
            .ImgPicked = ""
            .time = ""
            .Modality = ""
            .Reg_Date = ""
            .ExamDetail = ""
            
            .OrderDate = ""
            .OrderTime = ""
            .ReportDate = ""
            .ReportTime = ""
            .LastUpdateDate = ""
            .LastUpdateTime = ""
            
            .Division_from = ""
            .Division_on = ""
            
            .Division_Seq = ""
            .ClinicalImp = ""
            .TemplateName = ""
            .TemplateFile = ""
            .ChargeBy = ""
            
            .HIS_ReqNo = ""
            .Dr_follow = ""
       End With
       
       With save_Record
            .uni_key = ""
            .chartno = ""
            
            .Name = ""
            .Sex = ""
            .BirthDay = ""
            .Phone = ""
            .Address = ""
            .CitizenID = ""
            .Date = ""
            .Type = ""
            .Room = ""
            .Age = ""
            .Item1 = ""
            .Item2 = ""
            .Item3 = ""
            .Item4 = ""
            .Item5 = ""
            .Item6 = ""
            .Others = ""
            .Dr_from = ""
            .Dr_on = ""
            .Dr_order = ""
            .Dr_report = ""
            .Status = ""
            .Class = ""
            .UploadCode = ""
            
            .ImgPicked = ""
            .time = ""
            .Modality = ""
            .Reg_Date = ""
            .ExamDetail = ""
            
            .OrderDate = ""
            .OrderTime = ""
            .ReportDate = ""
            .ReportTime = ""
            .LastUpdateDate = ""
            .LastUpdateTime = ""
            
            .Division_from = ""
            .Division_on = ""
            
            .Division_Seq = ""
            .ClinicalImp = ""
            .TemplateName = ""
            .TemplateFile = ""
            
            .ChargeBy = ""
            
            .HIS_ReqNo = ""
            .Dr_follow = ""
       End With
    
    Return
End Sub

Private Sub dtpDateSort_CloseUp()
    txtDate = Format(dtpDateSort, "yyyy/MM/dd")
End Sub

Private Sub dtpDateSort1_CloseUp()
    txtDate1 = Format(dtpDateSort1, "yyyy/MM/dd")
End Sub

Private Sub Form_Activate()
    Dim x$
    Dim i%
    
    For i% = 0 To lblGrid.Count - 1
        lblGrid(i%).Enabled = False
    Next
    dbgOnline.Enabled = False
    cmdOpen.Visible = False
    Call dat_Refresh
    DoEvents
    Call dbgOnline_SelChange
    txtChartNo.SetFocus
    is_SRreport = False
    dbgOnline.Enabled = True
    cmdOpen.Visible = True
    For i% = 0 To lblGrid.Count - 1
        lblGrid(i%).Enabled = True
    Next
'    lblStatus(3).Caption = Show_Seed(Decoding_Seed_Number)
'    lblStatus(3).Caption = xInputINI("ImgSVR Host", "SiteEng", App.Path & "\xExamSVR.ini")
'    lblStatus(3).Caption = Winsock1.LocalHostName & " / " & Winsock1.LocalIP
'    lblStatus(3).Caption = GetPhysicalAddress & " / " & GetDiskSerialNumber("C:\") & " / " & Get_MB_SNo & " / " & GetIPAddress
    Call View_Only
    
End Sub

Private Sub View_Only()
    If Enable_Report$ = "NO" Then
        frmQueue.Caption = "CRS 檢查/報告作業 - View Only"
        cmdAppend.Enabled = False
        cmdDelete.Enabled = False
        cmdUpdate.Enabled = False
        cmdBasicEdit.Enabled = False
        SSCommand1.Enabled = False
    End If
End Sub

Private Sub Object_Initialize()
    Dim tmp$, filter$, i%

'    filter$ = "Tb_name='Exam_online'"
'    Call cmb_Table_Initial(db_Name$, "CRIS_Tb_def_detail", "Tb_fields", filter$, lstTb_fields)
'    Call cmb_Table_Initial(db_Name$, "CRIS_Tb_def_detail", "Tb_fields_caption", filter$, lstTb_fields_caption)
'    Call cmb_Table_Initial(db_Name$, "CRIS_Tb_def_detail", "Tb_fields_length", filter$, lstTb_fields_length)
'    Call cmb_Table_Initial(db_Name$, "CRIS_Tb_def_detail", "Tb_fields_caption", filter$, cmbOrder_field)
    
    Call cmb_Table_Initial(db_Name$, "CRIS_ExamType", "Label", "", cmbType)
    'For i% = 0 To cmbType.ListCount - 1: cmbXType.AddItem cmbType.List(i%): Next
    Call cmb_Table_Initial(db_Name$, "CRIS_ExamType", "Type", "", cmbType1)

'    Call cmb_Table_Initial(db_Name$, "CRIS_Reference", "DISTINCT Type", "Class='Doctor'", cmbDivision)
    Call cmb_Table_Initial(db_Name$, "CRIS_User", "DISTINCT System", "", cmbDivision)
    
    'Call cmb_DR_Initial(db_Name$, "CRIS_User", "Name", "Type='醫師'", cmbDoctor)
    

    'filter$ = " Class='Room' AND Type='共用' "
    'Call cmb_Table_Initial(db_Name$, "CRIS_Reference", "Remark", filter$, cmbXRoom)
    
    'filter$ = " Class='Dr_from' AND Type='共用' "
    'Call cmb_Table_Initial(db_Name$, "CRIS_Reference", "Remark", filter$, cmbXDr_from)
    cmbDr_from.AddItem "健檢以外"
    'cmbDr_from.AddItem "門診"
    'cmbDr_from.AddItem "急診"
    'cmbDr_from.AddItem "住院"
    cmbDr_from.AddItem "健診"
    cmbDr_from.AddItem "健管"
    cmbDr_from.AddItem "全部"
    cmbDr_from = Dr_from$
    
    cmbType.AddItem "全部"
    cmbType1.AddItem "-"
    cmbDoctor.AddItem "全部"
    
    txtDate1 = Format(Date, "yyyy/mm/dd")
    dtpDateSort1 = Date
    txtDate = Format(Date, "yyyy/mm/dd")
    dtpDateSort = Date
    cmbOrder_field = "狀態"
End Sub

Sub Form_Load()
    Dim bParmQry As Integer
    Dim tmp$, filter$, i%, k%
    Dim AllConfig$
    Dim SortOrder() As String
    Dim ttt(3) As String, yyy(3) As String
    Dim tSortOrder$
    
    tempRef = False
    
    ttt(0) = "CHARTNO": yyy(0) = "CHARTNO"
    ttt(1) = "UNI_KEY": yyy(1) = "UNI_KEY"
    ttt(2) = "EXAMDATE": yyy(2) = "EXAMDATE DESC"
    ttt(3) = "EXAMTIME": yyy(3) = "EXAMTIME DESC"
    
    AllConfig$ = InputINI("ImgSVR Host", "AllConfig", App.Path & "\ExamSVR.ini")
    
    If xDisplay_UnikeyName$ <> "" Then
        lblEntry(21).Caption = xDisplay_UnikeyName$
    End If
    
    If xDivision_On$ = "肝膽腸胃內科" Then
        lblEntry(22).Caption = "檢查醫師"
        lblEntry(10).Caption = "檢查醫師"
    End If
    
    If Len(LableStatus) > 0 Then
        lblGrid(5).Caption = LableStatus
    End If
    'Q_SortOrder$的檢查，以防在ini設定時填寫的格式錯誤
    If Len(Q_SortOrder$) > 0 Then
        SortOrder = Split(UCase(Q_SortOrder$), ",")
        tSortOrder$ = ""
        For i% = 0 To UBound(SortOrder)
            For k% = 0 To 3
                If Trim(SortOrder(i%)) = ttt(k%) Then
                    If Len(tSortOrder$) > 0 Then
                        tSortOrder$ = tSortOrder$ & ", "
                    End If
                    tSortOrder$ = tSortOrder$ & yyy(k%)
                End If
            Next
        Next
    End If
    If Len(tSortOrder$) < 1 Then
        Q_SortOrder$ = " ORDER BY ExamDate DESC, ExamTime DESC"
    Else
        Q_SortOrder$ = " ORDER BY " & tSortOrder$
    End If
    
    If AllConfig$ = "YES" Then
        SSCommand1.Enabled = True
    Else
        SSCommand1.Enabled = False
    End If
    
    is_SRreport = False
    
    If Len(Report_Name$) > 5 Then
        cmdReport.Enabled = True
    Else
        cmdReport.Enabled = False
    End If
    
    If Len(Report_Name1$) > 5 Then
        Command1.Enabled = True
    Else
        Command1.Enabled = False
    End If
    
    If Len(Report_Name2$) > 5 Then
        Command2.Enabled = True
    Else
        Command2.Enabled = False
    End If
    'tmp$ = Command
    'path_System$ = tmp$ & "\"
    'path_Define$ = tmp$ & "\Defines\"
    'path_Images$ = tmp$ & "\Images\"
    'db_Name$ = tmp$ & "\database\DBSVR.mdb"
    
    'k& = Shell("NET USE " & "\\10.15.5.20\CRIS_Images sameway /user:MPACS\cris", vbHide)
    'If k& <= 0 Then
    '   MsgBox "影像儲存設備異常，請聯繫系統管理員！"
    '   Exit Sub
    'End If
    
    'Call CodeSet_Define         '10 to 35 進位轉換對照表
    Call xReportTemplate_Get    '取入各科室報告對應表
    'Call setPrnForm            '取入報告列印底表
    DoEvents
    
    'adoOnline.ConnectionString = dbConnection$
    '/**/
    'datSource_SQL$ = "SELECT Uni_key, CRIS_Exam_online.ChartNo, Name, Age, Type, ExamDate, ExamTime, Room, " & _
                     "Dr_from, Dr_on, Division_from, Division_on, CRIS_Exam_online.Status, Class, UploadCode, " & _
                     "Item1, Item2, Item3, Item4, Item5, Item6, Others, " & _
                     "ImgPicked, Modality, Reg_Date, ExamDetail, " & _
                     "OrderDate, OrderTime, ReportDate, ReportTime, LastUpdateDate, LastUpdateTime, " & _
                     "Sex, Birthday, CitizenID, Phone, Address, System, " & _
                     "Dr_Order, Dr_Report, Division_Seq, ClinicalImp, TemplateName, TemplateFile, ChargeBy, HIS_ReqNo " & _
                     "FROM CRIS_Exam_online INNER JOIN CRIS_Patient_Info ON CRIS_Exam_online.ChartNo=CRIS_Patient_Info.ChartNo"
    '/**/
    'datSource_SQL$ = "SELECT Uni_key, CRIS_Exam_online.ChartNo, Name, Age, Type, ExamDate, ExamTime, Room, " & _
                     "Dr_from, Dr_on, Division_from, Division_on, CRIS_Exam_online.Status, Class, UploadCode, " & _
                     "Item1, Item2, Item3, Item4, Item5, Item6, Others, " & _
                     "ImgPicked, Modality, Reg_Date, ExamDetail, " & _
                     "OrderDate, OrderTime, ReportDate, ReportTime, LastUpdateDate, LastUpdateTime, " & _
                     "Sex, Birthday, CitizenID, Phone, Address, System, " & _
                     "Dr_Order, Dr_Report, Division_Seq, ClinicalImp, TemplateName, TemplateFile, ChargeBy, HIS_ReqNo " & _
                     "FROM CRIS_Exam_online with(nolock) INNER JOIN CRIS_Patient_Info with(nolock) ON CRIS_Exam_online.ChartNo=CRIS_Patient_Info.ChartNo"
    datSource_SQL$ = "SELECT Uni_key, CRIS_Exam_online.ChartNo, Name, Age, Type, ExamDate, ExamTime, Room, " & _
                     "Dr_from, Dr_on, Division_from, Division_on, CRIS_Exam_online.Status, Class, UploadCode, " & _
                     "Item1, Item2, Item3, Item4, Item5, Item6, Others, " & _
                     "ImgPicked, Modality, Reg_Date, ExamDetail, Zone, " & _
                     "OrderDate, OrderTime, ReportDate, ReportTime, LastUpdateDate, LastUpdateTime, " & _
                     "Sex, Birthday, CitizenID, Phone, Address, System, follow_dr, " & _
                     "Dr_Order, Dr_Report, Division_Seq, ClinicalImp, TemplateName, TemplateFile, ChargeBy, HIS_ReqNo " & _
                     "FROM CRIS_Exam_online INNER JOIN CRIS_Patient_Info ON CRIS_Exam_online.ChartNo=CRIS_Patient_Info.ChartNo"
    '/**/
    
    'datSource_SQL$ = "SELECT Uni_key, CRIS_Exam_online.ChartNo, Name, Age, Type, ExamDate, ExamTime, Room, " & _
                     "Dr_from, Dr_on, Division_from, Division_on, CRIS_Exam_online.Status, Class, UploadCode, " & _
                     "Item1, Item2, Item3, Item4, Item5, Item6, Others, " & _
                     "ImgPicked, Modality, Reg_Date, ExamDetail, " & _
                     "OrderDate, OrderTime, ReportDate, ReportTime, LastUpdateDate, LastUpdateTime, " & _
                     "Sex, Birthday, CitizenID, Phone, Address, System, " & _
                     "Dr_Order, Dr_Report, Division_Seq, ClinicalImp, TemplateName, TemplateFile, ChargeBy, HIS_ReqNo " & _
                     "FROM CRIS_Exam_online JOIN CRIS_Patient_Info ON CRIS_Exam_online.ChartNo=CRIS_Patient_Info.ChartNo"
    
    '2006/11/16 要求開單日不顯示 改為顯示申請單號
    '/**/
    'queue_SQL$ = "SELECT Uni_key, CRIS_Exam_online.ChartNo, Name, Sex, Age, Type, ExamDetail, ExamDate, ExamTime, Dr_from, Dr_on, Dr_Order, HIS_ReqNo " & _
                 "FROM CRIS_Exam_online INNER JOIN CRIS_Patient_Info ON CRIS_Exam_online.ChartNo=CRIS_Patient_Info.ChartNo"
    '/**/
    'queue_SQL$ = "SELECT Uni_key, CRIS_Exam_online.ChartNo, Name, Sex, Age, Type, ExamDetail, ExamDate, ExamTime, Dr_from, Dr_on, Dr_Order, HIS_ReqNo " & _
                 "FROM CRIS_Exam_online with(nolock) INNER JOIN CRIS_Patient_Info with(nolock) ON CRIS_Exam_online.ChartNo=CRIS_Patient_Info.ChartNo"
    
    '如果是肝膽腸胃內科的話，第一個欄位改為院區別標記；不然則仍為impression標記
    If xDivision_On$ <> "肝膽腸胃內科" Then
        queue_SQL$ = "SELECT (case when LENGTH(TRIM(cris_smart_save.impression)) > 0 then 'Y' else '' end) I, "
    Else
        queue_SQL$ = "SELECT '', "
    End If
    queue_SQL$ = queue_SQL$ & "CRIS_Exam_online.Uni_key, CRIS_Exam_online.ChartNo, Name, Sex, Age, Type, ExamDetail, ExamDate, ExamTime, Dr_from, Dr_on, Dr_Report, HIS_ReqNo, CRIS_Exam_online.hisup, CRIS_Exam_online.Zone "
    queue_SQL$ = queue_SQL$ & "FROM CRIS_Exam_online INNER JOIN CRIS_Patient_Info ON CRIS_Exam_online.ChartNo=CRIS_Patient_Info.ChartNo"
    If xDivision_On$ <> "肝膽腸胃內科" Then
        queue_SQL$ = queue_SQL$ & " left join cris_smart_save on CRIS_Exam_online.uni_key = cris_smart_save.uni_key"
    Else
'        queue_SQL$ = queue_SQL$ & " left join cris_smart_save on CRIS_Exam_online.uni_key = cris_smart_save.uni_key"
    End If
    
    '/**/
    queueCaption$ = "I  |檢查編號|病歷號碼　|姓名　　|性別|年齡|檢查類別　　|檢查細項  |檢查日期　|時 間|來源 |" & lblEntry(22).Caption & "       |報告醫師　　|申請單號　　|狀態|院區"
    
    Call Object_Initialize      '設定各個 Combo box 的初始值
'    Call lblGrid_Click(1)       '指定列出狀態為 待檢查 的報告
    
    Call array_Diction_Initial  '取入所有檢查別的片語辭組
    Call array_Phrase_Initial
    
    dbgOnline.FormatString = queueCaption$
    
    Call View_Only
    
    Me.ZOrder 0
    
    Exit Sub

loadErr:
    MsgBox "Error:" & err & " " & err.Description
    Unload Me

End Sub
Sub dat_Refresh()
    Dim i%, j%, datOrder$, datWhere$, SQL$, tblName1$, tblName2$, xFrom$
    Dim dbgControl As Variant, adoControl As Variant
    Dim RecordsNo&
    
    Dim adoDB As New adoDB.Connection
    Dim adoOnline1 As New adoDB.Recordset
    Dim adoRecord1 As New adoDB.Recordset
    Dim conn$, tmp$
    Dim tmpDr_from$
    
    Screen.MousePointer = vbHourglass
'    picWait.Visible = False
    DoEvents
    
    'On Error GoTo Refresh_Error
    
    lblRecordCount = "0"

    adoDB.Open dbConnection$ '"Persist Security Info=True;User ID=alantso;pwd=5682;dsn=EndoSVR;LoginTimeOut=3;"
'    ReDim aryRecord(34, 1000)
    ReDim aryRecord(15, 1000)
       
'    dbgOnline.BindToArray aryRecord
'    DoEvents
       
    '取 Online 記錄
    tblName1$ = "CRIS_Exam_Online": tblName2$ = "CRIS_Patient_Info": xFrom$ = "Online"
    
    Set adoControl = adoOnline1
    'dbgOnline.BindToArray aryRecord
    
    GoSub dbg_Refresh
    
    Set adoControl = Nothing
'    DoEvents
    
    If RecordsNo& > 1000 Then RecordsNo& = 1000
        lblRecordCount = str(RecordsNo&)
'        ReDim aryResult(34, RecordsNo&)
        ReDim aryResult(15, RecordsNo&)
        For j% = 0 To RecordsNo& - 1
            For i% = 0 To 15
                aryResult(i%, j%) = aryRecord(i%, j%)
            Next
        Next
        dbgOnline.BindToArray aryResult
'        DoEvents
        dbgOnline.ColHidden(14) = True
        dbgOnline.SelectionMode = flexSelectionByRow
        dbgOnline.FillStyle = flexFillRepeat
        
        tempRef = True
        For i = 1 To RecordsNo&
            dbgOnline.Select i, 0, i, 15
            '若HISUP=300時，黃色顯示該筆紀錄
            If dbgOnline.TextMatrix(dbgOnline.RowSel, 14) = "300" Then
                dbgOnline.CellBackColor = &H80FFFF
            Else
                dbgOnline.CellBackColor = &H80000005
            End If
        Next
        
        '清除最後一筆空白紀錄的顏色
        dbgOnline.Select i, 0, i, 15
        dbgOnline.CellBackColor = &H80000005
        tempRef = False
        
        If RecordsNo& >= 1 Then
           dbgOnline.Select 1, 1
        Else
            With Me
                For i% = 0 To .Controls.Count - 1
                    If (.Controls(i%).Tag = "xFields") Then
                       .Controls(i%).Caption = ""
                    End If
                Next
            End With
        End If
    
'    dbgOnline.BindToArray aryResult
    
    adoOnline1.Close
    Set adoOnline1 = Nothing
    
    adoDB.Close
    Set adoDB = Nothing
    
    DoEvents
    
'    picWait.Visible = False
    Screen.MousePointer = vbDefault
        
    On Error GoTo 0
    
    Exit Sub
    
Refresh_Error:
    MsgBox Error(err)
    Resume Next
    Return
    
    
dbg_Refresh:
    Dim sMode$, sqlSource$, tmpChartNo$
    Dim tempdate$
    
    lstFilter.Clear
    sqlSource$ = queue_SQL$
    
    '加入 WHERE Statement
    
    For i% = 0 To lblGrid.Count - 1
        If lblGrid(i%).ForeColor = &HFFFFFF Then
            If i = 5 Then
                txtStatus.Text = lblGrid(1).Caption
            ElseIf i = 6 Then
                txtStatus.Text = lblGrid(2).Caption
            Else
                txtStatus.Text = lblGrid(i%).Caption
            End If
        End If
    Next
    
    
    datWhere$ = ""
    If Len(txtName) > 0 Then
       tmpChartNo$ = Field_get("", "CRIS_Patient_Info", "ChartNo", "Name='" & txtName & "'")
       txtChartNo = tmpChartNo$
    End If
    
    If Len(txtAccessionNo) > 0 Then lstFilter.AddItem " CRIS_Exam_Online.Uni_key='" & txtAccessionNo & "' "
    'If Len(txtReqNo) > 0 Then lstFilter.AddItem " HIS_ReqNo='" & txtReqNo & "' "
    If Len(txtReqNo) > 0 Then lstFilter.AddItem " CRIS_Exam_Online.Uni_key='" & txtReqNo & "' "
    If Len(txtChartNo) > 0 Then lstFilter.AddItem " CRIS_Exam_Online.ChartNo='" & txtChartNo & "' "
        
    'tempdate$
    If txtStatus.Text = "已報告" Then
        tempdate$ = "report"
    Else
        tempdate$ = "exam"
    End If
    
    If Not lstFilter.ListCount > 0 Then
        If IsDate(txtDate) And IsDate(txtDate1) Then
            lstFilter.AddItem " (" & tempdate$ & "Date between '" & txtDate & "' and '" & txtDate1 & "') "
        Else
            If IsDate(txtDate) Then
                lstFilter.AddItem " " & tempdate$ & "Date='" & txtDate & "' "
            Else
                If IsDate(txtDate1) Then
                    lstFilter.AddItem " " & tempdate$ & "Date='" & txtDate1 & "' "
                End If
            End If
        End If
            
        If Len(cmbType) > 0 And cmbType <> "全部" Then lstFilter.AddItem " Type='" & cmbType & "' "
        If Len(cmbDoctor) > 0 And cmbDoctor <> "全部" Then lstFilter.AddItem " Dr_report='" & cmbDoctor & "' "
        If Len(cmbPhysician) > 0 And cmbPhysician <> "全部" Then lstFilter.AddItem " Dr_on='" & cmbPhysician & "' "
        'If Len(cmbDivision) > 0 And cmbDivision <> "全部" Then lstFilter.AddItem " Division_on='" & cmbDivision & "' "
        
        'If Len(cmbDr_from) > 0 And cmbDr_from <> "全部" Then lstFilter.AddItem " Dr_from='" & cmbDr_from & "' "
        tmpDr_from$ = ""
        Select Case cmbDr_from
               Case "門診": tmpDr_from$ = " Dr_from='門診' "
               Case "急診": tmpDr_from$ = " Dr_from='急診' "
               Case "住院": tmpDr_from$ = " Dr_from NOT IN ('門診','急診','健診','健管') "
               Case "健診": tmpDr_from$ = " Dr_from='健診' "
               Case "健管": tmpDr_from$ = " Dr_from='健管' "
               Case "健檢以外": tmpDr_from$ = " Dr_from NOT IN ('健診','健管') "
               Case "全部":: tmpDr_from$ = ""
        End Select
        If Len(tmpDr_from$) > 1 Then lstFilter.AddItem tmpDr_from$
        
        If Len(txtStatus) > 0 And txtStatus <> "全部" Then lstFilter.AddItem " CRIS_Exam_Online.Status='" & txtStatus & "' "
        
    End If
'        If UserType$ = "醫師" And Len(UserName$) > 0 Then
'           lstFilter.AddItem " Dr_on='" & UserName & "' "
'           cmbDoctor = UserName$
'        End If
'    End If
     
     lstFilter.AddItem " CRIS_Exam_Online.Status<>'已刪除' "
    
    tmp$ = ""
    For i% = 0 To lstFilter.ListCount - 1
        tmp$ = tmp$ & lstFilter.List(i%) & " AND "
    Next
    
    If Len(tmp$) > 5 Then tmp$ = " WHERE " & Left(tmp$, Len(tmp$) - 5)
    datWhere$ = tmp$
    
'    '當Q_SortOrder$="UNI_KEY"時變更為依照檢查單號做排序
'    If Q_SortOrder$ = "UNI_KEY" Then
'        datWhere$ = datWhere$ & " ORDER BY Uni_Key"
'    ElseIf Q_SortOrder$ = "CHARTNO" Then
'        datWhere$ = datWhere$ & " ORDER BY Chartno"
'    Else
'       datWhere$ = datWhere$ & " ORDER BY ExamDate DESC, ExamTime DESC"
'    End If
    datWhere$ = datWhere$ & Q_SortOrder$
    
    sqlSource$ = sqlSource$ & datWhere$
    adoControl.Open sqlSource$, adoDB, adOpenForwardOnly, adLockReadOnly
    Do While Not adoControl.EOF
       
       If RecordsNo& > 1000 Then
          'MsgBox "符合此條件之記錄超過 1000 筆，請加上精確條件後再重新搜尋"
          Exit Do
       End If
       
       For i% = 1 To adoControl.Fields.Count
           aryRecord(i% - 1, RecordsNo&) = NoNull(adoControl.Fields(i% - 1))
       Next
       RecordsNo& = RecordsNo& + 1
       adoControl.MoveNext
    Loop
    
'    DoEvents
    If InStr(datWhere$, "Status") < 1 Then
        For i% = 0 To lblGrid.Count - 1
            lblGrid(i%).BackColor = &H8000000F
            lblGrid(i%).ForeColor = &H80000012
        Next
    End If
    
    Return

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim rtn As Long
    
    rtn = WriteIni("Default Font", "Size", Font_Size%, path_Define$ & "ExamSVR.ini")
    
    'Load frmSplash
    'frmSplash.Show
    End
    
End Sub


Sub lblGrid_Click(Index As Integer)
    Dim i%
    
    For i% = 0 To lblGrid.Count - 1
        lblGrid(i%).BackColor = &H8000000F
        lblGrid(i%).ForeColor = &H80000012
        lblGrid(i%).Enabled = False
    Next
    cmdOpen.Visible = False
    dbgOnline.Enabled = False
    lblGrid(Index).BackColor = &HFF8080
    lblGrid(Index).ForeColor = &HFFFFFF
    If Index = 5 Then
        txtStatus.Text = lblGrid(1).Caption
    ElseIf Index = 6 Then
        txtStatus.Text = lblGrid(2).Caption
    Else
        txtStatus.Text = lblGrid(Index).Caption
    End If
    Call dat_Refresh
    For i% = 0 To lblGrid.Count - 1
        lblGrid(i%).Enabled = True
    Next
    dbgOnline.Enabled = True
    cmdOpen.Visible = True
End Sub


Private Sub SSCommand1_Click()
    Dim rcount As Integer
    Dim tUni_key$, tChartno$, tSql$
    Dim adoDB As adoDB.Connection
    
    Set adoDB = New adoDB.Connection
    adoDB.Open dbConnection$
    
    For rcount = 1 To dbgOnline.Rows - 1
        dbgOnline.Select rcount, 0, rcount, 13
        '若為已報告時，黃色顯示該筆紀錄
        If dbgOnline.TextMatrix(dbgOnline.RowSel, 14) = "已檢查" Then
            tUni_key$ = dbgOnline.TextMatrix(dbgOnline.RowSel, 1)
            tChartno$ = dbgOnline.TextMatrix(dbgOnline.RowSel, 2)
            tSql$ = "update cris_exam_online set "
            tSql$ = tSql$ & " class = 'UPREP', "
            tSql$ = tSql$ & " status = '已報告', "
            tSql$ = tSql$ & " hisup = '51' "
            tSql$ = tSql$ & " where status<>'已刪除' and uni_key = '" & tUni_key$ & "' "
            tSql$ = tSql$ & " and chartno = '" & tChartno$ & "' "
            Call DBRecordLog("update", tSql$, "查詢畫面確認全部報告，更新cris_exam_online")
            adoDB.Execute tSql$
        End If
    Next
    adoDB.Close
    Set adoDB = Nothing
    Call dat_Refresh
End Sub

Private Sub txtChartNo_GotFocus()
    
    txtChartNo.SelStart = 0
    txtChartNo.SelLength = Len(txtChartNo)

End Sub

Private Sub txtChartNo_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim tmp$, i%
    
    If KeyCode = 13 And Len(txtChartNo) > 0 Then
       
       'If Len(txtChartNo) < 10 Then
       '   txtChartNo = String(10 - Len(Trim(txtChartNo)), "0") & Trim(txtChartNo)
       'End If
       'txtChartNo = Format(txtChartNo, "0000000000")

       GoSub lblGrid_Clear
'       GoSub filter_Clear
       Call dat_Refresh
    Else
        If Len(txtChartNo) = 10 Then
           'If Check_ChartNo(txtChartNo) Then '如果有病例號合法性檢查公式的話
              GoSub lblGrid_Clear
'              GoSub filter_Clear
              Call dat_Refresh
           'End If
       End If
    End If
    Exit Sub
    
filter_Clear:
    
    cmbDoctor = ""
    cmbType = ""
    txtDate = ""
    txtDate1 = ""
    txtStatus = ""
    Return
    
lblGrid_Clear:
    For i% = 0 To lblGrid.Count - 1
        lblGrid(i%).BackColor = &H8000000F
        lblGrid(i%).ForeColor = &H80000012
    Next
    Return
    
End Sub



Private Sub txtReqNo_GotFocus()
    
    txtReqNo.SelStart = 0
    txtReqNo.SelLength = Len(txtChartNo)


End Sub

Private Sub txtReqNo_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim tmp$, i%
    Dim adoDB As New adoDB.Connection
    Dim adoRec1 As New adoDB.Recordset
    Dim sqlSource$
    
    If KeyCode = 13 And Len(txtReqNo) > 0 Then
        '20130304根據中山醫心內運心的USER要求，改為刷bar code後先檢查cris_exam_online內有無此筆單號
        '若無時則執行撈單，不管有無撈到單，之後皆按照原查詢程序作業
        
        '撈單
        If IS_Hync$ <> "NO" Then
            frmQueue.Enabled = False
        
            If (Len(Trim(txtReqNo.Text)) = 8) Or ((Len(Trim(txtReqNo.Text)) = 10) And (Left(Trim(txtReqNo.Text), 2) = "14")) Then
                Set currForm = Me
                Call GetHisSync(Trim(txtReqNo.Text))
            End If
        
            frmQueue.Enabled = True
        End If
        GoSub lblGrid_Clear
        GoSub filter_Clear
        Call dat_Refresh
    End If
    Exit Sub
    
filter_Clear:
    
    cmbDoctor = ""
    cmbType = ""
    txtDate = ""
    txtDate1 = ""
    txtStatus = ""
    txtChartNo = ""
    Return
    
lblGrid_Clear:
    For i% = 0 To lblGrid.Count - 1
        lblGrid(i%).BackColor = &H8000000F
        lblGrid(i%).ForeColor = &H80000012
    Next
    Return
    

End Sub
Private Sub clearDirectory()

    'tmpFileName$ = Dir(App.Path & "\*.html", vbNormal)
    'Do While Len(tmpFileName$) > 0
              
    '   Kill App.Path & "\" & tmpFileName$
    '   tmpFileName$ = Dir(App.Path & "\*.html", vbNormal)
    'Loop

End Sub

'傳回此筆記錄是本月份同類檢查的第幾份已報告記錄，含已刪除的報告
'只要有report date就算，因有的可能已報告又被蓋掉回已檢查
Function FindSigninSerial() As String
    Dim SQL_String As String
    Dim yy$, mm$
    Dim xCount As Integer
    
    yy$ = Format(Date, "YYYY")
    mm$ = Format(Date, "MM")

    SQL_String = "select * from cris_exam_online where type = '" & curr_Record.Type & "' "
    SQL_String = SQL_String & " and Division_On = '" & curr_Record.Division_on & "' "
    SQL_String = SQL_String & " and ReportDate >= '" & yy$ & "/" & mm$ & "/01' "
    SQL_String = SQL_String & " and ReportDate <= '" & yy$ & "/" & mm$ & "/31' "
    xCount = 1
    Call OpenRecordset(SQL_String, Connection, Recordset)
    While Not Recordset.EOF
        xCount = xCount + 1
        Recordset.MoveNext
    Wend
    FindSigninSerial = yy$ & mm$ & "-" & Format(xCount, "0000")
End Function
