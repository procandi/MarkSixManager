VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmOrderAddNewAll 
   BackColor       =   &H00808080&
   BorderStyle     =   1  '單線固定
   Caption         =   "產品價格變更"
   ClientHeight    =   10215
   ClientLeft      =   615
   ClientTop       =   840
   ClientWidth     =   9795
   Icon            =   "frmOrderAddNewAll.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10215
   ScaleWidth      =   9795
   Begin VB.TextBox txtBonusMoney_HKN2K 
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
      Left            =   3360
      TabIndex        =   30
      Top             =   6360
      Width           =   1455
   End
   Begin VB.TextBox txtBonusMoney_Lotto4K 
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
      Left            =   6480
      TabIndex        =   58
      Top             =   8760
      Width           =   1455
   End
   Begin VB.TextBox txtBonusMoney_LottoSpecial 
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
      Left            =   8040
      TabIndex        =   62
      Top             =   8760
      Width           =   1455
   End
   Begin VB.TextBox txtAddMoney_LottoSpecial 
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
      Left            =   8040
      TabIndex        =   61
      Top             =   8280
      Width           =   1455
   End
   Begin VB.TextBox txtWinningCount_LottoSpecial 
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
      Left            =   8040
      TabIndex        =   60
      Top             =   7800
      Width           =   1455
   End
   Begin VB.TextBox txtCurrentCount_LottoSpecial 
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
      Left            =   8040
      TabIndex        =   59
      Top             =   7320
      Width           =   1455
   End
   Begin VB.TextBox txtAddMoney_Lotto4K 
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
      Left            =   6480
      TabIndex        =   57
      Top             =   8280
      Width           =   1455
   End
   Begin VB.TextBox txtWinningCount_Lotto4K 
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
      Left            =   6480
      TabIndex        =   56
      Top             =   7800
      Width           =   1455
   End
   Begin VB.TextBox txtCurrentCount_Lotto4K 
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
      Left            =   6480
      TabIndex        =   55
      Top             =   7320
      Width           =   1455
   End
   Begin VB.TextBox txtBonusMoney_Lotto3K 
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
      TabIndex        =   54
      Top             =   8760
      Width           =   1455
   End
   Begin VB.TextBox txtAddMoney_Lotto3K 
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
      TabIndex        =   53
      Top             =   8280
      Width           =   1455
   End
   Begin VB.TextBox txtWinningCount_Lotto3K 
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
      TabIndex        =   52
      Top             =   7800
      Width           =   1455
   End
   Begin VB.TextBox txtCurrentCount_Lotto3K 
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
      TabIndex        =   51
      Top             =   7320
      Width           =   1455
   End
   Begin VB.TextBox txtBonusMoney_Lotto2K 
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
      Left            =   3360
      TabIndex        =   50
      Top             =   8760
      Width           =   1455
   End
   Begin VB.TextBox txtAddMoney_Lotto2K 
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
      Left            =   3360
      TabIndex        =   49
      Top             =   8280
      Width           =   1455
   End
   Begin VB.TextBox txtWinningCount_Lotto2K 
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
      Left            =   3360
      TabIndex        =   48
      Top             =   7800
      Width           =   1455
   End
   Begin VB.TextBox txtCurrentCount_Lotto2K 
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
      Left            =   3360
      TabIndex        =   47
      Top             =   7320
      Width           =   1455
   End
   Begin VB.TextBox txtBonusMoney_LottoCar 
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
      Left            =   1800
      TabIndex        =   46
      Top             =   8760
      Width           =   1455
   End
   Begin VB.TextBox txtAddMoney_LottoCar 
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
      Left            =   1800
      TabIndex        =   45
      Top             =   8280
      Width           =   1455
   End
   Begin VB.TextBox txtWinningCount_LottoCar 
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
      Left            =   1800
      TabIndex        =   44
      Top             =   7800
      Width           =   1455
   End
   Begin VB.TextBox txtCurrentCount_LottoCar 
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
      Left            =   1800
      TabIndex        =   43
      Top             =   7320
      Width           =   1455
   End
   Begin VB.TextBox txtBonusMoney_HKNSpecial 
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
      Left            =   8040
      TabIndex        =   42
      Top             =   6360
      Width           =   1455
   End
   Begin VB.TextBox txtAddMoney_HKNSpecial 
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
      Left            =   8040
      TabIndex        =   41
      Top             =   5880
      Width           =   1455
   End
   Begin VB.TextBox txtWinningCount_HKNSpecial 
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
      Left            =   8040
      TabIndex        =   40
      Top             =   5400
      Width           =   1455
   End
   Begin VB.TextBox txtCurrentCount_HKNSpecial 
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
      Left            =   8040
      TabIndex        =   39
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox txtBonusMoney_HKN4K 
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
      Left            =   6480
      TabIndex        =   38
      Top             =   6360
      Width           =   1455
   End
   Begin VB.TextBox txtAddMoney_HKN4K 
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
      Left            =   6480
      TabIndex        =   37
      Top             =   5880
      Width           =   1455
   End
   Begin VB.TextBox txtWinningCount_HKN4K 
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
      Left            =   6480
      TabIndex        =   36
      Top             =   5400
      Width           =   1455
   End
   Begin VB.TextBox txtCurrentCount_HKN4K 
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
      Left            =   6480
      TabIndex        =   35
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox txtBonusMoney_HKN3K 
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
      TabIndex        =   34
      Top             =   6360
      Width           =   1455
   End
   Begin VB.TextBox txtAddMoney_HKN3K 
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
      TabIndex        =   33
      Top             =   5880
      Width           =   1455
   End
   Begin VB.TextBox txtWinningCount_HKN3K 
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
      TabIndex        =   32
      Top             =   5400
      Width           =   1455
   End
   Begin Threed.SSPanel pnlBasic 
      Height          =   9375
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   9495
      _Version        =   65536
      _ExtentX        =   16748
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
      Begin VB.TextBox txtCurrentCount_HKN3K 
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
         TabIndex        =   31
         Top             =   4200
         Width           =   1455
      End
      Begin VB.TextBox txtAddMoney_HKN2K 
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
         Left            =   3120
         TabIndex        =   29
         Top             =   5160
         Width           =   1455
      End
      Begin VB.TextBox txtWinningCount_HKN2K 
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
         Left            =   3120
         TabIndex        =   28
         Top             =   4680
         Width           =   1455
      End
      Begin VB.TextBox txtCurrentCount_HKN2K 
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
         Left            =   3120
         TabIndex        =   27
         Top             =   4200
         Width           =   1455
      End
      Begin VB.TextBox txtBonusMoney_HKNCar 
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
         TabIndex        =   26
         Top             =   5640
         Width           =   1455
      End
      Begin VB.TextBox txtAddMoney_HKNCar 
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
         TabIndex        =   25
         Top             =   5160
         Width           =   1455
      End
      Begin VB.TextBox txtWinningCount_HKNCar 
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
         TabIndex        =   24
         Top             =   4680
         Width           =   1455
      End
      Begin VB.TextBox txtCurrentCount_HKNCar 
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
         TabIndex        =   23
         Top             =   4200
         Width           =   1455
      End
      Begin VB.TextBox txtBonusMoney_539Package 
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
         Left            =   7800
         TabIndex        =   22
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox txtAddMoney_539Package 
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
         Left            =   7800
         TabIndex        =   21
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox txtWinningCount_539Package 
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
         Left            =   7800
         TabIndex        =   20
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtCurrentCount_539Package 
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
         Left            =   7800
         TabIndex        =   19
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtBonusMoney_5394K 
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
         Left            =   6240
         TabIndex        =   18
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox txtAddMoney_5394K 
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
         Left            =   6240
         TabIndex        =   17
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox txtWinningCount_5394K 
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
         Left            =   6240
         TabIndex        =   16
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtCurrentCount_5394K 
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
         Left            =   6240
         TabIndex        =   15
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtBonusMoney_5393K 
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
         TabIndex        =   14
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox txtAddMoney_5393K 
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
         TabIndex        =   13
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox txtWinningCount_5393K 
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
         TabIndex        =   12
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtCurrentCount_5393K 
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
         TabIndex        =   11
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtBonusMoney_5392K 
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
         Left            =   3120
         TabIndex        =   10
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox txtAddMoney_5392K 
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
         Left            =   3120
         TabIndex        =   9
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox txtWinningCount_5392K 
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
         Left            =   3120
         TabIndex        =   8
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtCurrentCount_5392K 
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
         Left            =   3120
         TabIndex        =   7
         Top             =   1680
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
         Height          =   375
         Left            =   4320
         TabIndex        =   63
         Top             =   600
         Width           =   4935
      End
      Begin VB.TextBox txtBonusMoney_539Car 
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
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox txtAddMoney_539Car 
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
         Top             =   2640
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
         Left            =   1560
         MaxLength       =   256
         TabIndex        =   70
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtWinningCount_539Car 
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
         Left            =   7200
         Style           =   1  '圖片外觀
         TabIndex        =   65
         Top             =   8640
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
         Left            =   1440
         Style           =   1  '圖片外觀
         TabIndex        =   64
         Tag             =   "Edit"
         Top             =   8640
         Width           =   1335
      End
      Begin VB.TextBox txtCurrentCount_539Car 
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
         TabIndex        =   3
         Top             =   1680
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpCurrentDate 
         Height          =   375
         Left            =   1560
         TabIndex        =   71
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
         Format          =   103481347
         CurrentDate     =   42267
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
         Index           =   28
         Left            =   360
         TabIndex        =   97
         Top             =   6600
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
         Index           =   27
         Left            =   360
         TabIndex        =   96
         Top             =   7080
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
         Index           =   26
         Left            =   360
         TabIndex        =   95
         Top             =   7560
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
         Index           =   25
         Left            =   360
         TabIndex        =   94
         Top             =   8040
         Width           =   1215
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
         Index           =   24
         Left            =   360
         TabIndex        =   93
         Top             =   4200
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
         Index           =   23
         Left            =   360
         TabIndex        =   92
         Top             =   4680
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
         Index           =   22
         Left            =   360
         TabIndex        =   91
         Top             =   5160
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
         Index           =   21
         Left            =   360
         TabIndex        =   90
         Top             =   5640
         Width           =   1215
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   "大樂透_車"
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
         Index           =   20
         Left            =   1560
         TabIndex        =   89
         Top             =   6120
         Width           =   1455
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   "大樂透_2K"
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
         Index           =   19
         Left            =   3120
         TabIndex        =   88
         Top             =   6120
         Width           =   1455
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   "大樂透_3K"
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
         Left            =   4680
         TabIndex        =   87
         Top             =   6120
         Width           =   1455
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   "大樂透_4K"
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
         Left            =   6240
         TabIndex        =   86
         Top             =   6120
         Width           =   1455
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   "大樂透_特"
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
         Left            =   7800
         TabIndex        =   85
         Top             =   6120
         Width           =   1455
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   "港號_車"
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
         Left            =   1560
         TabIndex        =   84
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   "港號_2K"
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
         Left            =   3120
         TabIndex        =   83
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   "港號_3K"
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
         Left            =   4680
         TabIndex        =   82
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   "港號_4K"
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
         Left            =   6240
         TabIndex        =   81
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   "港號_特"
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
         Left            =   7800
         TabIndex        =   80
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   "539_3包"
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
         Index           =   9
         Left            =   7800
         TabIndex        =   79
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   "539_4K"
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
         Index           =   8
         Left            =   6240
         TabIndex        =   78
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   "539_3K"
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
         Index           =   7
         Left            =   4680
         TabIndex        =   77
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   "539_2K"
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
         Index           =   4
         Left            =   3120
         TabIndex        =   76
         Top             =   1200
         Width           =   1455
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
         Left            =   3120
         TabIndex        =   75
         Top             =   600
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
         TabIndex        =   74
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
         TabIndex        =   73
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
         TabIndex        =   69
         Top             =   600
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
         TabIndex        =   68
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
         TabIndex        =   67
         Top             =   120
         Width           =   8895
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
         TabIndex        =   66
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   "539_車"
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
         Left            =   1560
         TabIndex        =   2
         Top             =   1200
         Width           =   1455
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
      TabIndex        =   72
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
      TabIndex        =   1
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "frmOrderAddNewAll"
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

    Dim flag As Boolean
    Dim PID(100) As String
    Dim LastSwiftCode As String
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim BonusTarget As String
    
    flag = False
    
    SQL = "select * from [order] order by SwiftCode desc;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, order_rec)
    If order_rec.EOF Then
        LastSwiftCode = "0"
    Else
        LastSwiftCode = order_rec("SwiftCode")
    End If
    order_rec.Close
    
    SQL = "select * from product';"
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
    
    

    
    
    'update all data from UI
    
    '539
    If txtCurrentCount_539Car.Text <> "" Then
        Adodc1.Recordset.Fields.Item("CID").Value = basVariable.SelectCID
        Adodc1.Recordset.Fields.Item("CurrentDate").Value = txtCurrentDate.Text
        Adodc1.Recordset.Fields.Item("Note").Value = txtNote.Text
    
        LastSwiftCode = Val(LastSwiftCode) + 1
        Adodc1.Recordset.Fields.Item("SwiftCode").Value = LastSwiftCode
        Adodc1.Recordset.Fields.Item("PID").Value = PID(0)
        Adodc1.Recordset.Fields.Item("CurrentCount").Value = txtCurrentCount_539Car.Text
        If txtWinningCount_539Car.Text = "" Then
            Adodc1.Recordset.Fields.Item("WinningCount").Value = "0"
        Else
            Adodc1.Recordset.Fields.Item("WinningCount").Value = txtWinningCount_539Car.Text
        End If
        Adodc1.Recordset.Fields.Item("AddMoney").Value = txtAddMoney_539Car.Text
        Adodc1.Recordset.Fields.Item("BonusMoney").Value = txtBonusMoney_539Car.Text
    
        Call Adodc1.Recordset.Update
        Call Adodc1.Recordset.AddNew    'Call Form_Unload(0)
        
        flag = True
    End If
    If txtCurrentCount_5392K.Text <> "" Then
        Adodc1.Recordset.Fields.Item("CID").Value = basVariable.SelectCID
        Adodc1.Recordset.Fields.Item("CurrentDate").Value = txtCurrentDate.Text
        Adodc1.Recordset.Fields.Item("Note").Value = txtNote.Text
        
        LastSwiftCode = Val(LastSwiftCode) + 1
        Adodc1.Recordset.Fields.Item("SwiftCode").Value = LastSwiftCode
        Adodc1.Recordset.Fields.Item("PID").Value = PID(1)
        Adodc1.Recordset.Fields.Item("CurrentCount").Value = txtCurrentCount_5392K.Text
        If txtWinningCount_5392K.Text = "" Then
            Adodc1.Recordset.Fields.Item("WinningCount").Value = "0"
        Else
            Adodc1.Recordset.Fields.Item("WinningCount").Value = txtWinningCount_5392K.Text
        End If
        Adodc1.Recordset.Fields.Item("AddMoney").Value = txtAddMoney_5392K.Text
        Adodc1.Recordset.Fields.Item("BonusMoney").Value = txtBonusMoney_5392K.Text
    
        Call Adodc1.Recordset.Update
        Call Adodc1.Recordset.AddNew    'Call Form_Unload(0)
        
        flag = True
    End If
    If txtCurrentCount_5393K.Text <> "" Then
        Adodc1.Recordset.Fields.Item("CID").Value = basVariable.SelectCID
        Adodc1.Recordset.Fields.Item("CurrentDate").Value = txtCurrentDate.Text
        Adodc1.Recordset.Fields.Item("Note").Value = txtNote.Text
        
        LastSwiftCode = Val(LastSwiftCode) + 1
        Adodc1.Recordset.Fields.Item("SwiftCode").Value = LastSwiftCode
        Adodc1.Recordset.Fields.Item("PID").Value = PID(2)
        Adodc1.Recordset.Fields.Item("CurrentCount").Value = txtCurrentCount_5393K.Text
        If txtWinningCount_5393K.Text = "" Then
            Adodc1.Recordset.Fields.Item("WinningCount").Value = "0"
        Else
            Adodc1.Recordset.Fields.Item("WinningCount").Value = txtWinningCount_5393K.Text
        End If
        Adodc1.Recordset.Fields.Item("AddMoney").Value = txtAddMoney_5393K.Text
        Adodc1.Recordset.Fields.Item("BonusMoney").Value = txtBonusMoney_5393K.Text
    
        Call Adodc1.Recordset.Update
        Call Adodc1.Recordset.AddNew    'Call Form_Unload(0)
        
        flag = True
    End If
    If txtCurrentCount_5394K.Text <> "" Then
        Adodc1.Recordset.Fields.Item("CID").Value = basVariable.SelectCID
        Adodc1.Recordset.Fields.Item("CurrentDate").Value = txtCurrentDate.Text
        Adodc1.Recordset.Fields.Item("Note").Value = txtNote.Text
        
        LastSwiftCode = Val(LastSwiftCode) + 1
        Adodc1.Recordset.Fields.Item("SwiftCode").Value = LastSwiftCode
        Adodc1.Recordset.Fields.Item("PID").Value = PID(3)
        Adodc1.Recordset.Fields.Item("CurrentCount").Value = txtCurrentCount_5394K.Text
        If txtWinningCount_5394K.Text = "" Then
            Adodc1.Recordset.Fields.Item("WinningCount").Value = "0"
        Else
            Adodc1.Recordset.Fields.Item("WinningCount").Value = txtWinningCount_5394K.Text
        End If
        Adodc1.Recordset.Fields.Item("AddMoney").Value = txtAddMoney_5394K.Text
        Adodc1.Recordset.Fields.Item("BonusMoney").Value = txtBonusMoney_5394K.Text
    
        Call Adodc1.Recordset.Update
        Call Adodc1.Recordset.AddNew    'Call Form_Unload(0)
        
        flag = True
    End If
    If txtCurrentCount_539Package.Text <> "" Then
        Adodc1.Recordset.Fields.Item("CID").Value = basVariable.SelectCID
        Adodc1.Recordset.Fields.Item("CurrentDate").Value = txtCurrentDate.Text
        Adodc1.Recordset.Fields.Item("Note").Value = txtNote.Text
        
        LastSwiftCode = Val(LastSwiftCode) + 1
        Adodc1.Recordset.Fields.Item("SwiftCode").Value = LastSwiftCode
        Adodc1.Recordset.Fields.Item("PID").Value = PID(4)
        Adodc1.Recordset.Fields.Item("CurrentCount").Value = txtCurrentCount_539Package.Text
        If txtWinningCount_539Package.Text = "" Then
            Adodc1.Recordset.Fields.Item("WinningCount").Value = "0"
        Else
            Adodc1.Recordset.Fields.Item("WinningCount").Value = txtWinningCount_539Package.Text
        End If
        Adodc1.Recordset.Fields.Item("AddMoney").Value = txtAddMoney_539Package.Text
        Adodc1.Recordset.Fields.Item("BonusMoney").Value = txtBonusMoney_539Package.Text
    
        Call Adodc1.Recordset.Update
        Call Adodc1.Recordset.AddNew    'Call Form_Unload(0)
        
        flag = True
    End If
    
    'HKN
    If txtCurrentCount_HKNCar.Text <> "" Then
        Adodc1.Recordset.Fields.Item("CID").Value = basVariable.SelectCID
        Adodc1.Recordset.Fields.Item("CurrentDate").Value = txtCurrentDate.Text
        Adodc1.Recordset.Fields.Item("Note").Value = txtNote.Text
    
        LastSwiftCode = Val(LastSwiftCode) + 1
        Adodc1.Recordset.Fields.Item("SwiftCode").Value = LastSwiftCode
        Adodc1.Recordset.Fields.Item("PID").Value = PID(5)
        Adodc1.Recordset.Fields.Item("CurrentCount").Value = txtCurrentCount_HKNCar.Text
        If txtWinningCount_HKNCar.Text = "" Then
            Adodc1.Recordset.Fields.Item("WinningCount").Value = "0"
        Else
            Adodc1.Recordset.Fields.Item("WinningCount").Value = txtWinningCount_HKNCar.Text
        End If
        Adodc1.Recordset.Fields.Item("AddMoney").Value = txtAddMoney_HKNCar.Text
        Adodc1.Recordset.Fields.Item("BonusMoney").Value = txtBonusMoney_HKNCar.Text
    
        Call Adodc1.Recordset.Update
        Call Adodc1.Recordset.AddNew    'Call Form_Unload(0)
        
        flag = True
    End If
    If txtCurrentCount_HKN2K.Text <> "" Then
        Adodc1.Recordset.Fields.Item("CID").Value = basVariable.SelectCID
        Adodc1.Recordset.Fields.Item("CurrentDate").Value = txtCurrentDate.Text
        Adodc1.Recordset.Fields.Item("Note").Value = txtNote.Text
        
        LastSwiftCode = Val(LastSwiftCode) + 1
        Adodc1.Recordset.Fields.Item("SwiftCode").Value = LastSwiftCode
        Adodc1.Recordset.Fields.Item("PID").Value = PID(6)
        Adodc1.Recordset.Fields.Item("CurrentCount").Value = txtCurrentCount_HKN2K.Text
        If txtWinningCount_HKN2K.Text = "" Then
            Adodc1.Recordset.Fields.Item("WinningCount").Value = "0"
        Else
            Adodc1.Recordset.Fields.Item("WinningCount").Value = txtWinningCount_HKN2K.Text
        End If
        Adodc1.Recordset.Fields.Item("AddMoney").Value = txtAddMoney_HKN2K.Text
        Adodc1.Recordset.Fields.Item("BonusMoney").Value = txtBonusMoney_HKN2K.Text
    
        Call Adodc1.Recordset.Update
        Call Adodc1.Recordset.AddNew    'Call Form_Unload(0)
        
        flag = True
    End If
    If txtCurrentCount_HKN3K.Text <> "" Then
        Adodc1.Recordset.Fields.Item("CID").Value = basVariable.SelectCID
        Adodc1.Recordset.Fields.Item("CurrentDate").Value = txtCurrentDate.Text
        Adodc1.Recordset.Fields.Item("Note").Value = txtNote.Text
        
        LastSwiftCode = Val(LastSwiftCode) + 1
        Adodc1.Recordset.Fields.Item("SwiftCode").Value = LastSwiftCode
        Adodc1.Recordset.Fields.Item("PID").Value = PID(7)
        Adodc1.Recordset.Fields.Item("CurrentCount").Value = txtCurrentCount_HKN3K.Text
        If txtWinningCount_HKN3K.Text = "" Then
            Adodc1.Recordset.Fields.Item("WinningCount").Value = "0"
        Else
            Adodc1.Recordset.Fields.Item("WinningCount").Value = txtWinningCount_HKN3K.Text
        End If
        Adodc1.Recordset.Fields.Item("AddMoney").Value = txtAddMoney_HKN3K.Text
        Adodc1.Recordset.Fields.Item("BonusMoney").Value = txtBonusMoney_HKN3K.Text
    
        Call Adodc1.Recordset.Update
        Call Adodc1.Recordset.AddNew    'Call Form_Unload(0)
        
        flag = True
    End If
    If txtCurrentCount_HKN4K.Text <> "" Then
        Adodc1.Recordset.Fields.Item("CID").Value = basVariable.SelectCID
        Adodc1.Recordset.Fields.Item("CurrentDate").Value = txtCurrentDate.Text
        Adodc1.Recordset.Fields.Item("Note").Value = txtNote.Text
        
        LastSwiftCode = Val(LastSwiftCode) + 1
        Adodc1.Recordset.Fields.Item("SwiftCode").Value = LastSwiftCode
        Adodc1.Recordset.Fields.Item("PID").Value = PID(8)
        Adodc1.Recordset.Fields.Item("CurrentCount").Value = txtCurrentCount_HKN4K.Text
        If txtWinningCount_HKN4K.Text = "" Then
            Adodc1.Recordset.Fields.Item("WinningCount").Value = "0"
        Else
            Adodc1.Recordset.Fields.Item("WinningCount").Value = txtWinningCount_HKN4K.Text
        End If
        Adodc1.Recordset.Fields.Item("AddMoney").Value = txtAddMoney_HKN4K.Text
        Adodc1.Recordset.Fields.Item("BonusMoney").Value = txtBonusMoney_HKN4K.Text
    
        Call Adodc1.Recordset.Update
        Call Adodc1.Recordset.AddNew    'Call Form_Unload(0)
        
        flag = True
    End If
    If txtCurrentCount_HKNSpecial.Text <> "" Then
        Adodc1.Recordset.Fields.Item("CID").Value = basVariable.SelectCID
        Adodc1.Recordset.Fields.Item("CurrentDate").Value = txtCurrentDate.Text
        Adodc1.Recordset.Fields.Item("Note").Value = txtNote.Text
        
        LastSwiftCode = Val(LastSwiftCode) + 1
        Adodc1.Recordset.Fields.Item("SwiftCode").Value = LastSwiftCode
        Adodc1.Recordset.Fields.Item("PID").Value = PID(9)
        Adodc1.Recordset.Fields.Item("CurrentCount").Value = txtCurrentCount_HKNSpecial.Text
        If txtWinningCount_HKNSpecial.Text = "" Then
            Adodc1.Recordset.Fields.Item("WinningCount").Value = "0"
        Else
            Adodc1.Recordset.Fields.Item("WinningCount").Value = txtWinningCount_HKNSpecial.Text
        End If
        Adodc1.Recordset.Fields.Item("AddMoney").Value = txtAddMoney_HKNSpecial.Text
        Adodc1.Recordset.Fields.Item("BonusMoney").Value = txtBonusMoney_HKNSpecial.Text
    
        Call Adodc1.Recordset.Update
        Call Adodc1.Recordset.AddNew    'Call Form_Unload(0)
        
        flag = True
    End If
    
    'Lotto
    If txtCurrentCount_LottoCar.Text <> "" Then
        Adodc1.Recordset.Fields.Item("CID").Value = basVariable.SelectCID
        Adodc1.Recordset.Fields.Item("CurrentDate").Value = txtCurrentDate.Text
        Adodc1.Recordset.Fields.Item("Note").Value = txtNote.Text
    
        LastSwiftCode = Val(LastSwiftCode) + 1
        Adodc1.Recordset.Fields.Item("SwiftCode").Value = LastSwiftCode
        Adodc1.Recordset.Fields.Item("PID").Value = PID(10)
        Adodc1.Recordset.Fields.Item("CurrentCount").Value = txtCurrentCount_LottoCar.Text
        If txtWinningCount_LottoCar.Text = "" Then
            Adodc1.Recordset.Fields.Item("WinningCount").Value = "0"
        Else
            Adodc1.Recordset.Fields.Item("WinningCount").Value = txtWinningCount_LottoCar.Text
        End If
        Adodc1.Recordset.Fields.Item("AddMoney").Value = txtAddMoney_LottoCar.Text
        Adodc1.Recordset.Fields.Item("BonusMoney").Value = txtBonusMoney_LottoCar.Text
    
        Call Adodc1.Recordset.Update
        Call Adodc1.Recordset.AddNew    'Call Form_Unload(0)
        
        flag = True
    End If
    If txtCurrentCount_Lotto2K.Text <> "" Then
        Adodc1.Recordset.Fields.Item("CID").Value = basVariable.SelectCID
        Adodc1.Recordset.Fields.Item("CurrentDate").Value = txtCurrentDate.Text
        Adodc1.Recordset.Fields.Item("Note").Value = txtNote.Text
        
        LastSwiftCode = Val(LastSwiftCode) + 1
        Adodc1.Recordset.Fields.Item("SwiftCode").Value = LastSwiftCode
        Adodc1.Recordset.Fields.Item("PID").Value = PID(11)
        Adodc1.Recordset.Fields.Item("CurrentCount").Value = txtCurrentCount_Lotto2K.Text
        If txtWinningCount_Lotto2K.Text = "" Then
            Adodc1.Recordset.Fields.Item("WinningCount").Value = "0"
        Else
            Adodc1.Recordset.Fields.Item("WinningCount").Value = txtWinningCount_Lotto2K.Text
        End If
        Adodc1.Recordset.Fields.Item("AddMoney").Value = txtAddMoney_Lotto2K.Text
        Adodc1.Recordset.Fields.Item("BonusMoney").Value = txtBonusMoney_Lotto2K.Text
    
        Call Adodc1.Recordset.Update
        Call Adodc1.Recordset.AddNew    'Call Form_Unload(0)
        
        flag = True
    End If
    If txtCurrentCount_Lotto3K.Text <> "" Then
        Adodc1.Recordset.Fields.Item("CID").Value = basVariable.SelectCID
        Adodc1.Recordset.Fields.Item("CurrentDate").Value = txtCurrentDate.Text
        Adodc1.Recordset.Fields.Item("Note").Value = txtNote.Text
        
        LastSwiftCode = Val(LastSwiftCode) + 1
        Adodc1.Recordset.Fields.Item("SwiftCode").Value = LastSwiftCode
        Adodc1.Recordset.Fields.Item("PID").Value = PID(12)
        Adodc1.Recordset.Fields.Item("CurrentCount").Value = txtCurrentCount_Lotto3K.Text
        If txtWinningCount_Lotto3K.Text = "" Then
            Adodc1.Recordset.Fields.Item("WinningCount").Value = "0"
        Else
            Adodc1.Recordset.Fields.Item("WinningCount").Value = txtWinningCount_Lotto3K.Text
        End If
        Adodc1.Recordset.Fields.Item("AddMoney").Value = txtAddMoney_Lotto3K.Text
        Adodc1.Recordset.Fields.Item("BonusMoney").Value = txtBonusMoney_Lotto3K.Text
    
        Call Adodc1.Recordset.Update
        Call Adodc1.Recordset.AddNew    'Call Form_Unload(0)
        
        flag = True
    End If
    If txtCurrentCount_Lotto4K.Text <> "" Then
        Adodc1.Recordset.Fields.Item("CID").Value = basVariable.SelectCID
        Adodc1.Recordset.Fields.Item("CurrentDate").Value = txtCurrentDate.Text
        Adodc1.Recordset.Fields.Item("Note").Value = txtNote.Text
        
        LastSwiftCode = Val(LastSwiftCode) + 1
        Adodc1.Recordset.Fields.Item("SwiftCode").Value = LastSwiftCode
        Adodc1.Recordset.Fields.Item("PID").Value = PID(13)
        Adodc1.Recordset.Fields.Item("CurrentCount").Value = txtCurrentCount_Lotto4K.Text
        If txtWinningCount_Lotto4K.Text = "" Then
            Adodc1.Recordset.Fields.Item("WinningCount").Value = "0"
        Else
            Adodc1.Recordset.Fields.Item("WinningCount").Value = txtWinningCount_Lotto4K.Text
        End If
        Adodc1.Recordset.Fields.Item("AddMoney").Value = txtAddMoney_Lotto4K.Text
        Adodc1.Recordset.Fields.Item("BonusMoney").Value = txtBonusMoney_Lotto4K.Text
    
        Call Adodc1.Recordset.Update
        Call Adodc1.Recordset.AddNew    'Call Form_Unload(0)
        
        flag = True
    End If
    If txtCurrentCount_LottoSpecial.Text <> "" Then
        Adodc1.Recordset.Fields.Item("CID").Value = basVariable.SelectCID
        Adodc1.Recordset.Fields.Item("CurrentDate").Value = txtCurrentDate.Text
        Adodc1.Recordset.Fields.Item("Note").Value = txtNote.Text
        
        LastSwiftCode = Val(LastSwiftCode) + 1
        Adodc1.Recordset.Fields.Item("SwiftCode").Value = LastSwiftCode
        Adodc1.Recordset.Fields.Item("PID").Value = PID(14)
        Adodc1.Recordset.Fields.Item("CurrentCount").Value = txtCurrentCount_LottoSpecial.Text
        If txtWinningCount_LottoSpecial.Text = "" Then
            Adodc1.Recordset.Fields.Item("WinningCount").Value = "0"
        Else
            Adodc1.Recordset.Fields.Item("WinningCount").Value = txtWinningCount_LottoSpecial.Text
        End If
        Adodc1.Recordset.Fields.Item("AddMoney").Value = txtAddMoney_LottoSpecial.Text
        Adodc1.Recordset.Fields.Item("BonusMoney").Value = txtBonusMoney_LottoSpecial.Text
    
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
        txtCurrentCount_539Car.Text = ""
        txtWinningCount_539Car.Text = ""
        txtAddMoney_539Car.Text = ""
        txtBonusMoney_539Car.Text = ""
        txtCurrentCount_5392K.Text = ""
        txtWinningCount_5392K.Text = ""
        txtAddMoney_5392K.Text = ""
        txtBonusMoney_5392K.Text = ""
        txtCurrentCount_5393K.Text = ""
        txtWinningCount_5393K.Text = ""
        txtAddMoney_5393K.Text = ""
        txtBonusMoney_5393K.Text = ""
        txtCurrentCount_5394K.Text = ""
        txtWinningCount_5394K.Text = ""
        txtAddMoney_5394K.Text = ""
        txtBonusMoney_5394K.Text = ""
        txtCurrentCount_539Package.Text = ""
        txtWinningCount_539Package.Text = ""
        txtAddMoney_539Package.Text = ""
        txtBonusMoney_539Package.Text = ""
        
        txtCurrentCount_HKNCar.Text = ""
        txtWinningCount_HKNCar.Text = ""
        txtAddMoney_HKNCar.Text = ""
        txtBonusMoney_HKNCar.Text = ""
        txtCurrentCount_HKN2K.Text = ""
        txtWinningCount_HKN2K.Text = ""
        txtAddMoney_HKN2K.Text = ""
        txtBonusMoney_HKN2K.Text = ""
        txtCurrentCount_HKN3K.Text = ""
        txtWinningCount_HKN3K.Text = ""
        txtAddMoney_HKN3K.Text = ""
        txtBonusMoney_HKN3K.Text = ""
        txtCurrentCount_HKN4K.Text = ""
        txtWinningCount_HKN4K.Text = ""
        txtAddMoney_HKN4K.Text = ""
        txtBonusMoney_HKN4K.Text = ""
        txtCurrentCount_HKNSpecial.Text = ""
        txtWinningCount_HKNSpecial.Text = ""
        txtAddMoney_HKNSpecial.Text = ""
        txtBonusMoney_HKNSpecial.Text = ""
        
        txtCurrentCount_LottoCar.Text = ""
        txtWinningCount_LottoCar.Text = ""
        txtAddMoney_LottoCar.Text = ""
        txtBonusMoney_LottoCar.Text = ""
        txtCurrentCount_Lotto2K.Text = ""
        txtWinningCount_Lotto2K.Text = ""
        txtAddMoney_Lotto2K.Text = ""
        txtBonusMoney_Lotto2K.Text = ""
        txtCurrentCount_Lotto3K.Text = ""
        txtWinningCount_Lotto3K.Text = ""
        txtAddMoney_Lotto3K.Text = ""
        txtBonusMoney_Lotto3K.Text = ""
        txtCurrentCount_Lotto4K.Text = ""
        txtWinningCount_Lotto4K.Text = ""
        txtAddMoney_Lotto4K.Text = ""
        txtBonusMoney_Lotto4K.Text = ""
        txtCurrentCount_LottoSpecial.Text = ""
        txtWinningCount_LottoSpecial.Text = ""
        txtAddMoney_LottoSpecial.Text = ""
        txtBonusMoney_LottoSpecial.Text = ""
        
        addCount = addCount + 1
        lblAddCount.Caption = "已新增" & addCount & "筆"
    End If
    
    Call txtCurrentDate.SetFocus
    
    
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

    
    dtpCurrentDate.Value = Format(DateTime.Now, "yyyy/MM/dd")
    txtCurrentDate.Text = Format(DateTime.Now, "yyyy/MM/dd")
    
    
    addCount = 0
    lblAddCount.Caption = "已新增" & addCount & "筆"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmOrder.Show
    Unload Me
End Sub

Private Sub txtCurrentDate_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtNote.SetFocus
    End If
End Sub

Private Sub txtNote_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtCurrentCount_539Car.SetFocus
    End If
End Sub


'539 KeyUp
Private Sub txtCurrentCount_539Car_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtWinningCount_539Car.SetFocus
    End If
End Sub
Private Sub txtWinningCount_539Car_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtAddMoney_539Car.SetFocus
    End If
End Sub
Private Sub txtAddMoney_539Car_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtBonusMoney_539Car.SetFocus
    End If
End Sub
Private Sub txtBonusMoney_539Car_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtCurrentCount_5392K.SetFocus
    End If
End Sub
Private Sub txtCurrentCount_5392K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtWinningCount_5392K.SetFocus
    End If
End Sub
Private Sub txtWinningCount_5392K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtAddMoney_5392K.SetFocus
    End If
End Sub
Private Sub txtAddMoney_5392K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtBonusMoney_5392K.SetFocus
    End If
End Sub
Private Sub txtBonusMoney_5392K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtCurrentCount_5393K.SetFocus
    End If
End Sub
Private Sub txtCurrentCount_5393K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtWinningCount_5393K.SetFocus
    End If
End Sub
Private Sub txtWinningCount_5393K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtAddMoney_5393K.SetFocus
    End If
End Sub
Private Sub txtAddMoney_5393K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtBonusMoney_5393K.SetFocus
    End If
End Sub
Private Sub txtBonusMoney_5393K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtCurrentCount_5394K.SetFocus
    End If
End Sub
Private Sub txtCurrentCount_5394K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtWinningCount_5394K.SetFocus
    End If
End Sub
Private Sub txtWinningCount_5394K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtAddMoney_5394K.SetFocus
    End If
End Sub
Private Sub txtAddMoney_5394K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtBonusMoney_5394K.SetFocus
    End If
End Sub
Private Sub txtBonusMoney_5394K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtCurrentCount_539Package.SetFocus
    End If
End Sub
Private Sub txtCurrentCount_539Package_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtWinningCount_539Package.SetFocus
    End If
End Sub
Private Sub txtWinningCount_539Package_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtAddMoney_539Package.SetFocus
    End If
End Sub
Private Sub txtAddMoney_539Package_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtBonusMoney_539Package.SetFocus
    End If
End Sub
Private Sub txtBonusMoney_539Package_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtCurrentCount_HKNCar.SetFocus
    End If
End Sub

'HKN KeyUp
Private Sub txtCurrentCount_HKNCar_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtWinningCount_HKNCar.SetFocus
    End If
End Sub
Private Sub txtWinningCount_HKNCar_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtAddMoney_HKNCar.SetFocus
    End If
End Sub
Private Sub txtAddMoney_HKNCar_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtBonusMoney_HKNCar.SetFocus
    End If
End Sub
Private Sub txtBonusMoney_HKNCar_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtCurrentCount_HKN2K.SetFocus
    End If
End Sub
Private Sub txtCurrentCount_HKN2K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtWinningCount_HKN2K.SetFocus
    End If
End Sub
Private Sub txtWinningCount_HKN2K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtAddMoney_HKN2K.SetFocus
    End If
End Sub
Private Sub txtAddMoney_HKN2K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtBonusMoney_HKN2K.SetFocus
    End If
End Sub
Private Sub txtBonusMoney_HKN2K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtCurrentCount_HKN3K.SetFocus
    End If
End Sub
Private Sub txtCurrentCount_HKN3K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtWinningCount_HKN3K.SetFocus
    End If
End Sub
Private Sub txtWinningCount_HKN3K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtAddMoney_HKN3K.SetFocus
    End If
End Sub
Private Sub txtAddMoney_HKN3K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtBonusMoney_HKN3K.SetFocus
    End If
End Sub
Private Sub txtBonusMoney_HKN3K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtCurrentCount_HKN4K.SetFocus
    End If
End Sub
Private Sub txtCurrentCount_HKN4K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtWinningCount_HKN4K.SetFocus
    End If
End Sub
Private Sub txtWinningCount_HKN4K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtAddMoney_HKN4K.SetFocus
    End If
End Sub
Private Sub txtAddMoney_HKN4K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtBonusMoney_HKN4K.SetFocus
    End If
End Sub
Private Sub txtBonusMoney_HKN4K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtCurrentCount_HKNSpecial.SetFocus
    End If
End Sub
Private Sub txtCurrentCount_HKNSpecial_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtWinningCount_HKNSpecial.SetFocus
    End If
End Sub
Private Sub txtWinningCount_HKNSpecial_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtAddMoney_HKNSpecial.SetFocus
    End If
End Sub
Private Sub txtAddMoney_HKNSpecial_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtBonusMoney_HKNSpecial.SetFocus
    End If
End Sub
Private Sub txtBonusMoney_HKNSpecial_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtCurrentCount_LottoCar.SetFocus
    End If
End Sub

'Lotto KeyUp
Private Sub txtCurrentCount_LottoCar_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtWinningCount_LottoCar.SetFocus
    End If
End Sub
Private Sub txtWinningCount_LottoCar_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtAddMoney_LottoCar.SetFocus
    End If
End Sub
Private Sub txtAddMoney_LottoCar_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtBonusMoney_LottoCar.SetFocus
    End If
End Sub
Private Sub txtBonusMoney_LottoCar_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtCurrentCount_Lotto2K.SetFocus
    End If
End Sub
Private Sub txtCurrentCount_Lotto2K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtWinningCount_Lotto2K.SetFocus
    End If
End Sub
Private Sub txtWinningCount_Lotto2K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtAddMoney_Lotto2K.SetFocus
    End If
End Sub
Private Sub txtAddMoney_Lotto2K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtBonusMoney_Lotto2K.SetFocus
    End If
End Sub
Private Sub txtBonusMoney_Lotto2K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtCurrentCount_Lotto3K.SetFocus
    End If
End Sub
Private Sub txtCurrentCount_Lotto3K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtWinningCount_Lotto3K.SetFocus
    End If
End Sub
Private Sub txtWinningCount_Lotto3K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtAddMoney_Lotto3K.SetFocus
    End If
End Sub
Private Sub txtAddMoney_Lotto3K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtBonusMoney_Lotto3K.SetFocus
    End If
End Sub
Private Sub txtBonusMoney_Lotto3K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtCurrentCount_Lotto4K.SetFocus
    End If
End Sub
Private Sub txtCurrentCount_Lotto4K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtWinningCount_Lotto4K.SetFocus
    End If
End Sub
Private Sub txtWinningCount_Lotto4K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtAddMoney_Lotto4K.SetFocus
    End If
End Sub
Private Sub txtAddMoney_Lotto4K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtBonusMoney_Lotto4K.SetFocus
    End If
End Sub
Private Sub txtBonusMoney_Lotto4K_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtCurrentCount_LottoSpecial.SetFocus
    End If
End Sub
Private Sub txtCurrentCount_LottoSpecial_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtWinningCount_LottoSpecial.SetFocus
    End If
End Sub
Private Sub txtWinningCount_LottoSpecial_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtAddMoney_LottoSpecial.SetFocus
    End If
End Sub
Private Sub txtAddMoney_LottoSpecial_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtBonusMoney_LottoSpecial.SetFocus
    End If
End Sub
Private Sub txtBonusMoney_LottoSpecial_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call cmdUpdate_Click
    End If
End Sub
