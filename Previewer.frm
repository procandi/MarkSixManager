VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRepPreview 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  '��u�T�w
   Caption         =   "CRS ���i�w���Ҧ�"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15225
   Icon            =   "Previewer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   734
   ScaleMode       =   3  '����
   ScaleWidth      =   1015
   StartUpPosition =   2  '�ù�����
   Begin VB.CommandButton cmdMain 
      Caption         =   "��v�T�{���i"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      TabIndex        =   1
      ToolTipText     =   "�N�w���C�L���e�ର�v���äW��"
      Top             =   10080
      Width           =   2175
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "���i�x�s"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   6600
      TabIndex        =   37
      ToolTipText     =   "�ର�w���i���x�s"
      Top             =   10080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "�ȤW�Ǽƾ�"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   5760
      TabIndex        =   21
      ToolTipText     =   "�ȤW�ǼƾڡA�����ܧ���i���A"
      Top             =   10080
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "�ȦC�L"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2520
      TabIndex        =   4
      ToolTipText     =   "�ȦC�L���W��"
      Top             =   10080
      Width           =   1455
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "�C�L�νT�{���i"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4200
      TabIndex        =   36
      ToolTipText     =   "�C�L�B�T�{���i"
      Top             =   10080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   9855
      Left            =   13080
      ScaleHeight     =   9795
      ScaleWidth      =   1995
      TabIndex        =   22
      Top             =   120
      Width           =   2055
      Begin VB.Image imgOptionFalse 
         Height          =   1335
         Index           =   5
         Left            =   2760
         Picture         =   "Previewer.frx":0442
         Stretch         =   -1  'True
         Top             =   8280
         Width           =   1095
      End
      Begin VB.Image imgOptionFalse 
         Height          =   1335
         Index           =   4
         Left            =   2760
         Picture         =   "Previewer.frx":3AB4
         Stretch         =   -1  'True
         Top             =   6840
         Width           =   1095
      End
      Begin VB.Image imgOptionTrue 
         Height          =   1335
         Index           =   5
         Left            =   3960
         Picture         =   "Previewer.frx":8116
         Stretch         =   -1  'True
         Top             =   8280
         Width           =   1095
      End
      Begin VB.Image imgOptionTrue 
         Height          =   1335
         Index           =   4
         Left            =   3960
         Picture         =   "Previewer.frx":B2D8
         Stretch         =   -1  'True
         Top             =   6840
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         X1              =   120
         X2              =   1920
         Y1              =   6600
         Y2              =   6600
      End
      Begin VB.Label Label5 
         Alignment       =   2  '�m�����
         Appearance      =   0  '����
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '�z��
         Caption         =   "Only Text"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1335
         Index           =   5
         Left            =   120
         TabIndex        =   33
         Top             =   8280
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   2  '�m�����
         Appearance      =   0  '����
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '�z��
         Caption         =   "Only Image "
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1335
         Index           =   4
         Left            =   120
         TabIndex        =   32
         Top             =   6840
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   2  '�m�����
         Appearance      =   0  '����
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '�z��
         Caption         =   "1X1"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1335
         Index           =   3
         Left            =   120
         TabIndex        =   31
         Top             =   4920
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   2  '�m�����
         Appearance      =   0  '����
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '�z��
         Caption         =   "4x2"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1335
         Index           =   2
         Left            =   120
         TabIndex        =   30
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   2  '�m�����
         Appearance      =   0  '����
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '�z��
         Caption         =   "3x2"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1335
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   2  '�m�����
         Appearance      =   0  '����
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '�z��
         Caption         =   "2x2"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1335
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   735
      End
      Begin VB.Image imgOption 
         BorderStyle     =   1  '��u�T�w
         Height          =   1335
         Index           =   5
         Left            =   840
         Picture         =   "Previewer.frx":F242
         Stretch         =   -1  'True
         Top             =   8280
         Width           =   1095
      End
      Begin VB.Image imgOption 
         BorderStyle     =   1  '��u�T�w
         Height          =   1335
         Index           =   4
         Left            =   840
         Picture         =   "Previewer.frx":12404
         Stretch         =   -1  'True
         Top             =   6840
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   2  '�m�����
         Appearance      =   0  '����
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '�z��
         Caption         =   "���i�C�L�˦�"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   1695
      End
      Begin VB.Image imgOptionTrue 
         Height          =   1335
         Index           =   0
         Left            =   3960
         Picture         =   "Previewer.frx":1636E
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1095
      End
      Begin VB.Image imgOptionTrue 
         Height          =   1335
         Index           =   1
         Left            =   3960
         Picture         =   "Previewer.frx":304BC
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Image imgOptionTrue 
         Height          =   1335
         Index           =   2
         Left            =   3960
         Picture         =   "Previewer.frx":4A2A6
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Image imgOptionTrue 
         Height          =   1335
         Index           =   3
         Left            =   3960
         Picture         =   "Previewer.frx":64090
         Stretch         =   -1  'True
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Image imgOptionFalse 
         Height          =   1335
         Index           =   3
         Left            =   2760
         Picture         =   "Previewer.frx":7E3CA
         Stretch         =   -1  'True
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Image imgOptionFalse 
         Height          =   1335
         Index           =   2
         Left            =   2760
         Picture         =   "Previewer.frx":981B4
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Image imgOptionFalse 
         Height          =   1335
         Index           =   1
         Left            =   2760
         Picture         =   "Previewer.frx":B1A56
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Image imgOptionFalse 
         Height          =   1335
         Index           =   0
         Left            =   2760
         Picture         =   "Previewer.frx":CADB8
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1095
      End
      Begin VB.Image imgOption 
         BorderStyle     =   1  '��u�T�w
         Height          =   1335
         Index           =   3
         Left            =   840
         Picture         =   "Previewer.frx":E49BA
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   1095
      End
      Begin VB.Image imgOption 
         BorderStyle     =   1  '��u�T�w
         Height          =   1335
         Index           =   2
         Left            =   840
         Picture         =   "Previewer.frx":FECF4
         Stretch         =   -1  'True
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Image imgOption 
         BorderStyle     =   1  '��u�T�w
         Height          =   1335
         Index           =   1
         Left            =   840
         Picture         =   "Previewer.frx":118ADE
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Image imgOption 
         BorderStyle     =   1  '��u�T�w
         Height          =   1335
         Index           =   0
         Left            =   840
         Picture         =   "Previewer.frx":1328C8
         Stretch         =   -1  'True
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.ListBox lstPrinter 
      Height          =   780
      Left            =   3600
      TabIndex        =   20
      Top             =   9000
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.ListBox lstPrnFile 
      Height          =   2400
      Left            =   9960
      TabIndex        =   18
      Top             =   5400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdPage 
      Caption         =   "&D ���@��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   10080
      TabIndex        =   11
      Top             =   10560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdPage 
      Caption         =   "&U �e�@��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   8760
      TabIndex        =   10
      Top             =   10440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog ComDialog 
      Left            =   120
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frmOption 
      Caption         =   "Option"
      Height          =   3855
      Left            =   2160
      TabIndex        =   6
      Top             =   4080
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CommandButton cmdOption 
         Caption         =   "&C ����"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3240
         TabIndex        =   8
         Top             =   3360
         Width           =   2175
      End
      Begin VB.CommandButton cmdOption 
         Caption         =   "&S �s��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   7
         Top             =   3360
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "���i��v��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   7560
      TabIndex        =   3
      Top             =   10080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "&X ���}"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   13080
      TabIndex        =   2
      Top             =   10080
      Width           =   2055
   End
   Begin SHDocVwCtl.WebBrowser wbReport 
      Height          =   9855
      Left            =   120
      TabIndex        =   34
      Top             =   120
      Width           =   12855
      ExtentX         =   22675
      ExtentY         =   17383
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.OptionButton optFont 
      Caption         =   "���p"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   11040
      TabIndex        =   27
      Top             =   10560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton optFont 
      Caption         =   "���j"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   10560
      TabIndex        =   26
      Top             =   10560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton optFont 
      Caption         =   "Default"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   9480
      TabIndex        =   24
      Top             =   10560
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblPrinter 
      BackStyle       =   0  '�z��
      BorderStyle     =   1  '��u�T�w
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   0
      Top             =   10080
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  '�m�����
      Appearance      =   0  '����
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  '��u�T�w
      Caption         =   "CRS Report Process "
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   11400
      TabIndex        =   35
      Top             =   10560
      Width           =   3735
   End
   Begin VB.Label lblOSVersion 
      Alignment       =   2  '�m�����
      Appearance      =   0  '����
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  '��u�T�w
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8880
      TabIndex        =   19
      Top             =   10560
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '�z��
      BorderStyle     =   1  '��u�T�w
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   8280
      TabIndex        =   17
      Top             =   10560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblPages 
      Alignment       =   2  '�m�����
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  '��u�T�w
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   13
      Top             =   10560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '�z��
      BorderStyle     =   1  '��u�T�w
      Caption         =   "�@"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   7080
      TabIndex        =   16
      Top             =   10560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '�z��
      BorderStyle     =   1  '��u�T�w
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   6600
      TabIndex        =   15
      Top             =   10560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblPage 
      Alignment       =   2  '�m�����
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  '��u�T�w
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   12
      Top             =   10560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '�z��
      BorderStyle     =   1  '��u�T�w
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   5520
      TabIndex        =   14
      Top             =   10560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  '��u�T�w
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   5
      Top             =   10560
      Width           =   7455
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  '��u�T�w
      Caption         =   "�t�ΰT��"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      TabIndex        =   9
      Top             =   10560
      Width           =   1215
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  '��u�T�w
      Caption         =   "�r��j�p"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      TabIndex        =   25
      Top             =   9600
      Visible         =   0   'False
      Width           =   4695
   End
End
Attribute VB_Name = "frmRepPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim defPrinter$
Dim systemDefaultDevice$

Sub cmdMain_Click(Index As Integer)
    Dim source_Printer$, system_print$
    Dim i As Integer
    
    For i = 0 To 6
        cmdMain(i).Enabled = False
    Next
    Select Case Index
        Case 0:
            Unload Me
            Exit Sub
        Case 1:
'            ComDialog.ShowPrinter
        '��v�T�{���i�A��v���W��
'            Load Form1
            If Len(Trim(curr_Record.Dr_report)) <= 0 Then
                MsgBox "�L���i��v��Ƥ��i�W�ǡA�ж�J��A�ǰe!"
            ElseIf Len(Trim(curr_Record.Dr_on)) <= 0 And Need_Dr_On$ <> "YES" Then
                MsgBox "�L�ޮv��Ƥ��i�W�ǡA�ж�J��A�ǰe!"
            Else
                If Len(Trim(curr_Record.Dr_on)) <= 0 Then
                    curr_Record.Dr_on = currForm.cmbDr_Report.Text
                End If
'                DoEvents
                If No_Report_Image$ <> "YES" Then
                    Form1.Show 1
                End If
'                DoEvents
                Call currForm.Record_update(True)
'                DoEvents
                lblStatus(0) = "���i�w��v���A�äW�ǧ���"
            End If
        Case 2: 'wbReport.Navigate path_System & "Tmp\ReportNew.html"
            If No_Report_Image$ <> "YES" Then
                Form1.Show 1
            End If
        Case 3:
        '�ȦC�L
            '���w���˽s��
            'If defPrinter$ <> "" Then GoSub Redirect_Printer_to_ImgSVR_Printer
            
             wbReport.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER, 0, 0
            
'            If Not curr_Record.Type = "�f�z�e�U��" Then
'               Call saveHTMLFile
'            End If
            
'            DoEvents
            'MsgBox wbReport.
'                curr_Record.Status = "�w�C�L"
            'Unload Me
            Call currForm.Record_update(False)
            lblStatus(0) = "�C�L����"
            'If defPrinter$ <> "" Then GoSub Redirect_Printer_to_Default_Printer
        
        Case 4:
        '�ȤW�ǼƾڡA���ܧ���i���A
            'wbReport.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DODEFAULT
            Call currForm.Record_update(True, False)
            lblStatus(0) = "�ȤW�ǼƾڡA�����ܧ���i���A"
        '�x�s�ΦC�L
        Case 5:
            If Len(Trim(curr_Record.Dr_report)) <= 0 Then
                MsgBox "�L���i��v��Ƥ��i�W�ǡA�ж�J��A�ǰe!"
            ElseIf Len(Trim(curr_Record.Dr_on)) <= 0 And Need_Dr_On$ <> "YES" Then
                MsgBox "�L�ޮv��Ƥ��i�W�ǡA�ж�J��A�ǰe!"
            Else
                If Len(Trim(curr_Record.Dr_on)) <= 0 Then
                    curr_Record.Dr_on = currForm.cmbDr_Report.Text
                End If
'                DoEvents
                If No_Report_Image$ <> "YES" Then
                    Form1.Show 1
                End If
'                DoEvents
                Call currForm.Record_update(True)
'                DoEvents
                
                wbReport.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER, 0, 0
                lblStatus(0) = "���i�w�C�L�B��v���A�äW�ǧ���"
            End If
        '���x�s
        Case 6:
            Call currForm.Record_update(False)
            lblStatus(0) = "���i�x�s����"
    End Select
    wbReport.Refresh
    For i = 0 To 6
        cmdMain(i).Enabled = True
    Next
    
    Exit Sub
    
Redirect_Printer_to_ImgSVR_Printer:
    For i% = 0 To lstPrinter.ListCount - 1
         If InStr(lstPrinter.List(i%), Printer.DeviceName) Then
            systemDefaultDevice$ = lstPrinter.List(i%)
         End If
         If InStr(lstPrinter.List(i%), defPrinter$) Then
            prnDevice = lstPrinter.List(i%)
         End If
    Next
'    Length = LenB(StrConv(prnDevice, vbFromUnicode)) + 1
'    If OSVersion = "Windows XP" Then
    
        'Using WMI to set default printer-----------------
'        Const strCls = "Win32_Printer" ' WMI Class
'        GetObject("winmgmts:").InstancesOf(strCls)(strCls & ".DeviceID=""" & systemDefaultDevice$ & """").setDefaultPrinter
        '�����ɦL����ɦ]��IP��}�e����\�ɡA����double�~�ॿ�T����
        prnDevice = Replace(prnDevice, "\", "\\")
        
        strComputer = "."
        Set objWMIService = GetObject("winmgmts:" _
            & "{impersonationLevel=impersonate}!\\" _
            & strComputer & "\root\cimv2")
        Set colInstalledPrinters = objWMIService.ExecQuery _
            ("Select * from Win32_Printer Where Name = '" & prnDevice & "'")
        For Each objPrinter In colInstalledPrinters
            Call objPrinter.SetDefaultPrinter
        Next
        DoEvents
'    Else
'
'        'Using SendMessage to set default printer -------------------------------------------------------------
'        Length = LenB(StrConv(systemDefaultDevice$, vbFromUnicode)) + 1
'        WriteProfileString "windows", "device", systemDefaultDevice$
''        SendMessage HWND_BROADCAST, WM_WININICHANGE, 32767&, ByVal "windows"
'        SendMessage HWND_BROADCAST, WM_WININICHANGE, 0, ByVal "windows"
'        DoEvents
'
'    End If
'    WriteProfileString "windows", "device", prnDevice
''    SendMessage HWND_BROADCAST, WM_WININICHANGE, 32767&, ByVal "windows"
'    SendMessage HWND_BROADCAST, WM_WININICHANGE, 0, ByVal "windows"
'    DoEvents
    Return
    
Redirect_Printer_to_Default_Printer:
    Length = LenB(StrConv(systemDefaultDevice$, vbFromUnicode)) + 1
    WriteProfileString "windows", "device", systemDefaultDevice$
'    SendMessage HWND_BROADCAST, WM_WININICHANGE, 32767&, ByVal "windows"
    SendMessage HWND_BROADCAST, WM_WININICHANGE, 0, ByVal "windows"
    DoEvents
    Return
    
Batch_Print:
            For i% = 0 To lstPrnFile.ListCount - 1
                    wbReport.Navigate lstPrnFile.List(i%)
                    Do While frmRepPreview.wbReport.Busy
                         DoEvents
                    Loop
                    DoEvents
                    Do While frmRepPreview.wbReport.Busy: DoEvents: Loop
                    frmRepPreview.wbReport.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER, Null, Null
            Next
    Return
    
End Sub
Private Sub Redirect_Printer_to_ImgSVR_Printer()
    Dim prnDevice As String
    
    On Error GoTo prnError
    
    For i% = 0 To lstPrinter.ListCount
         If InStr(lstPrinter.List(i%), Printer.DeviceName) Then
            systemDefaultDevice$ = lstPrinter.List(i%)
         End If
         If InStr(lstPrinter.List(i%), defPrinter$) Then
            prnDevice = lstPrinter.List(i%)
         End If
    Next
    

'    If systemDefaultDevice$ = "" Then systemDefaultDevice$ = lstPrinter.List(i%)
    
    If systemDefaultDevice$ = prnDevice$ Then Exit Sub
    
'    If OSVersion = "Windows XP" Then

        'Using WMI to set default printer ---------------------------------------------------------------------
'        Const strCls = "Win32_Printer" ' WMI Class
'        GetObject("winmgmts:").InstancesOf(strCls)(strCls & ".DeviceID=""" & prnDevice & """").SetDefaultPrinter
        '�����ɦL����ɦ]��IP��}�e����\�ɡA����double�~�ॿ�T����
        prnDevice = Replace(prnDevice, "\", "\\")
        
        strComputer = "."
        Set objWMIService = GetObject("winmgmts:" _
            & "{impersonationLevel=impersonate}!\\" _
            & strComputer & "\root\cimv2")
        Set colInstalledPrinters = objWMIService.ExecQuery _
            ("Select * from Win32_Printer Where Name = '" & prnDevice & "'")
        For Each objPrinter In colInstalledPrinters
            Call objPrinter.SetDefaultPrinter
        Next
'    Else
'
'        'Using SendMessage to set default printer -------------------------------------------------------------
'        Length = LenB(StrConv(prnDevice, vbFromUnicode)) + 1
'        WriteProfileString "windows", "device", prnDevice
'        SendMessage HWND_BROADCAST, WM_WININICHANGE, 32767&, ByVal "windows"
'    End If

    'Using WSH script to set default printer --------------------------------------------------------------
    'CreateObject("WScript.Network").setDefaultPrinter prnDevice
    
    DoEvents

    Exit Sub
    
prnError:
    Resume Next
    
End Sub
Private Sub Redirect_Printer_to_Default_Printer()

    
    If systemDefaultDevice$ = lblPrinter Then Exit Sub
    
'    If OSVersion = "Windows XP" Then
    
        'Using WMI to set default printer-----------------
'        Const strCls = "Win32_Printer" ' WMI Class
'        GetObject("winmgmts:").InstancesOf(strCls)(strCls & ".DeviceID=""" & systemDefaultDevice$ & """").SetDefaultPrinter
        '�����ɦL����ɦ]��IP��}�e����\�ɡA����double�~�ॿ�T����
        systemDefaultDevice$ = Replace(systemDefaultDevice$, "\", "\\")
        
        strComputer = "."
        Set objWMIService = GetObject("winmgmts:" _
            & "{impersonationLevel=impersonate}!\\" _
            & strComputer & "\root\cimv2")
        Set colInstalledPrinters = objWMIService.ExecQuery _
            ("Select * from Win32_Printer Where Name = '" & systemDefaultDevice$ & "'")
        For Each objPrinter In colInstalledPrinters
            Call objPrinter.SetDefaultPrinter
        Next
'    Else
'
'        'Using SendMessage to set default printer -------------------------------------------------------------
'        Length = LenB(StrConv(systemDefaultDevice$, vbFromUnicode)) + 1
'        WriteProfileString "windows", "device", systemDefaultDevice$
''        SendMessage HWND_BROADCAST, WM_WININICHANGE, 32767&, ByVal "windows"
'        SendMessage HWND_BROADCAST, WM_WININICHANGE, 0, ByVal "windows"
'        DoEvents
'
'    End If
    
    'Using WSH script to set default printer----------
    'CreateObject("WScript.Network").setDefaultPrinter systemDefaultDevice$
    
    DoEvents

End Sub

Private Sub Redirect_Printer_to_ImgSVR_PrinterOld()
    
    On Error GoTo prnError
    
    For i% = 0 To lstPrinter.ListCount
         If InStr(lstPrinter.List(i%), Printer.DeviceName) Then
            systemDefaultDevice$ = lstPrinter.List(i%)
         End If
         If InStr(lstPrinter.List(i%), defPrinter$) Then
            prnDevice = lstPrinter.List(i%)
         End If
    Next
    
    If systemDefaultDevice$ = "" Then systemDefaultDevice$ = lstPrinter.List(i%)
    
    Length = LenB(StrConv(prnDevice, vbFromUnicode)) + 1
    WriteProfileString "windows", "device", prnDevice
    SendMessage HWND_BROADCAST, WM_WININICHANGE, 32767&, ByVal "windows"
    DoEvents
    On Error GoTo 0
    Exit Sub
    
prnError:
    Resume Next
    
End Sub
    
Private Sub Redirect_Printer_to_Default_PrinterOld()

    Length = LenB(StrConv(systemDefaultDevice$, vbFromUnicode)) + 1
    WriteProfileString "windows", "device", systemDefaultDevice$
    SendMessage HWND_BROADCAST, WM_WININICHANGE, 32767&, ByVal "windows"
    DoEvents

End Sub

Private Sub saveHTMLFile()
    Dim sourceHTML$, targetHTML$
    
    targetHTML$ = path_Images & "\Img" & Format(curr_Record.Date, "yyMM")
    If Not isFileExist(targetHTML$, vbDirectory) Then MkDir targetHTML$
    targetHTML$ = targetHTML$ & "\" & Trim(curr_Record.uni_key) & ".html"

'   ���i���e���åi��y���W�� Server �����i�ɮװ��D---------------------------------------------
'    FileCopy App.Path & "\Report" & curr_Record.ChartNo & "_1.html", targetHTML$
    
    tmp$ = lstPrnFile.List(0)
    FileCopy tmp$, targetHTML$
'----------------------------------------------------------------------------------------------


    'curr_Record.LastUpdateDate = Format(Date, "yyyy/MM/dd")
    'curr_Record.LastUpdateTime = Format(Time, "hh:NN:ss")
    'curr_Record.ReportDate = IIf(Len(curr_Record.ReportDate) < 10, Format(Date, "yyyy/MM/dd"), curr_Record.ReportDate)
    'curr_Record.ReportTime = IIf(Len(curr_Record.ReportTime) < 8, Format(Time, "hh:NN:ss"), curr_Record.ReportTime)
    'curr_Record.Status = "�w���i"

End Sub

Private Sub cmdPage_Click(Index As Integer)
    
    If Index = 0 Then
        If Val(lblPage) > 1 Then
            lblPage = Trim(str(Val(lblPage) - 1))
            wbReport.Navigate lstPrnFile.List(Val(lblPage) - 1)
        End If
    Else
        If Val(lblPage) < Val(lblPages) Then
            lblPage = Trim(str(Val(lblPage) + 1))
            wbReport.Navigate lstPrnFile.List(Val(lblPage) - 1)
        End If
    End If
    
    If Val(lblPage) = Val(lblPages) Then
        cmdPage(1).Enabled = False
    Else
        cmdPage(1).Enabled = True
    End If
    
    If Val(lblPage) = 1 Then
        cmdPage(0).Enabled = False
    Else
        cmdPage(0).Enabled = True
    End If
    
End Sub

Private Sub Form_Activate()

    If Val(lblPage) = Val(lblPages) Then
        cmdPage(1).Enabled = False
    Else
        cmdPage(1).Enabled = True
    End If
    
    If Val(lblPage) = 1 Then
        cmdPage(0).Enabled = False
    Else
        cmdPage(0).Enabled = True
    End If

End Sub

Private Sub Form_Load()
    Dim tmpPrinter$
    Dim rtn As Long, tmpA As String * 260, nRet As Long
    Dim tmp$, leftString$, i%
        
    currForm.Enabled = False 'currForm.Enabled = False
    Me.ZOrder 0

'    For i% = 0 To Printers.count - 1
'        lstPrinter.AddItem Printers(i%).DeviceName & "," & _
'        Printers(i%).DriverName & "," & _
'        Printers(i%).Port
'    Next
    For i% = 0 To Printers.Count - 1
        lstPrinter.AddItem Printers(i%).DeviceName
    Next
    
    If xall_up_button$ = "YES" Then
        cmdMain(5).Visible = True
    End If
    
    For i% = 0 To imgOptionFalse.Count - 1
        imgOption(i%).Picture = imgOptionFalse(i%).Picture
    Next
    '/*�ª��A�Ҧ�����������w�]���O�a�ĤG�ӿﶵ*/
    'imgOption(1).Picture = imgOptionTrue(1).Picture
    '/*�s���A�߾ɺޤ�N���n�a�Ĥ@�ӡA��l����*/
    If Login_LastOpenReportType = "�߾ɺޤ�N" Then
        imgOption(0).Picture = imgOptionTrue(0).Picture
    Else
        imgOption(1).Picture = imgOptionTrue(1).Picture
    End If
    '/*�p�حק諸(20100413)*/
    'frmRepPreview.wbReport.Refresh
    'frmRepPreview.Show
    
    tmp$ = currForm.txtExamType & "���i"
'    rtn = ReadINI("Default Printer", "Device_" & tmp$, "", tmpA, Len(tmpA), App.Path & "\ExamSVR.ini")
'    If Not rtn = 0 Then tmp$ = Trim(Left(tmpA, rtn))
'    tmpPrinter$ = ini_Purge(tmp$, rtn)
    tmpPrinter$ = InputINI("Default Printer", "Device_" & tmp$, App.Path & "\ExamSVR.ini")
    
    If Len(tmpPrinter$) = 0 Then
'       rtn = ReadINI("Default Printer", "Device_��L", "", tmpA, Len(tmpA), App.Path & "\ExamSVR.ini")
'       If Not rtn = 0 Then tmp$ = Trim(Left(tmpA, rtn))
'       tmpPrinter$ = ini_Purge(tmp$, rtn)
       tmpPrinter$ = InputINI("Default Printer", "Device_��L", App.Path & "\ExamSVR.ini")
    End If
    
    defPrinter$ = ""
    For i% = 0 To Printers.Count - 1
        If Printers(i%).DeviceName = tmpPrinter$ Then
           defPrinter$ = tmpPrinter$
        End If
    Next
    lblPrinter = defPrinter$
    
    '���s��s�W�w�A�o��v�~�i�H�T�{���i
    If Need_Dr_Confirm$ = "YES" And frmQueue.lblUserType <> "��v" Then
        cmdMain(1).Enabled = False
    Else
        cmdMain(1).Enabled = True
    End If
    
    Call Redirect_Printer_to_ImgSVR_Printer
    
    If No_Report_Image$ = "YES" Then
        cmdMain(2).Enabled = False
    Else
        cmdMain(2).Enabled = True
    End If

    lblOSVersion = OSVersion$
    
    '���s��M�סA�u���B�ʤ߹q�ϳ���~�}��N���i��v���\��
'    If UCase(Trim(curr_Record.TemplateFile)) = "R341007.RPS" Or UCase(Trim(curr_Record.TemplateFile)) = "R341005.RPS" Then
'        cmdMain(1).Enabled = True
'    Else
'        cmdMain(1).Enabled = False
'    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'frmReport.Enabled = True
    'frmReport.SetFocus
    
    '/**/
    Call Redirect_Printer_to_Default_Printer
    '/**/
    If FSO.FileExists(App.Path & "\" & curr_Record.uni_key & ".HTML") Then
        FSO.DeleteFile (App.Path & "\" & curr_Record.uni_key & ".HTML")
    End If
    If FSO.FileExists(App.Path & "\" & curr_Record.uni_key & ".TB6") Then
        FSO.DeleteFile (App.Path & "\" & curr_Record.uni_key & ".TB6")
    End If
    If FSO.FileExists(App.Path & "\" & curr_Record.uni_key & ".TXT") Then
        FSO.DeleteFile (App.Path & "\" & curr_Record.uni_key & ".TXT")
    End If
    If FSO.FileExists(App.Path & "\Report" & curr_Record.uni_key & ".html") Then
        FSO.DeleteFile (App.Path & "\Report" & curr_Record.uni_key & ".html")
    End If
    currForm.Enabled = True
    currForm.Visible = True
    currForm.SetFocus
    
End Sub

Public Sub imgOption_Click(Index As Integer)
    Dim fontSize$
    
    '�j�z���z�~�쪺�j�z���ˬd���T�w�����A���i�D��
    If UCase(Trim(curr_Record.TemplateFile)) = "COLON_OUT.RPS" Then
        Exit Sub
    End If
    
    For i% = 0 To imgOptionFalse.Count - 1
        imgOption(i%).Picture = imgOptionFalse(i%).Picture
    Next
    imgOption(Index).Picture = imgOptionTrue(Index).Picture
    
    For i% = 0 To optFont.Count - 1
        If optFont(i%) Then fontSize$ = optFont(i%).Caption
    Next
    
    'fontSize$ = "���p"
    Select Case Index
           Case 0: Call currForm.ReportPrnOption(2, fontSize$)
           Case 1: Call currForm.ReportPrnOption(3, fontSize$)
           Case 2: Call currForm.ReportPrnOption(4, fontSize$)
           Case 3: Call currForm.ReportPrnOption(1, fontSize$)
           
           Case 4: Call currForm.ReportPrnOption(9, fontSize$) 'Only Image
           Case 5: Call currForm.ReportPrnOption(0, fontSize$) 'Only Text
    End Select
    frmRepPreview.wbReport.Refresh
    
End Sub

Private Sub optFont_Click(Index As Integer)
    Dim fontSize$
    
    For i% = 0 To imgOptionFalse.Count - 1
        imgOption(i%).Picture = imgOptionFalse(i%).Picture
    Next
    imgOption(Index).Picture = imgOptionTrue(Index).Picture
    
    For i% = 0 To optFont.Count - 1
        If optFont(i%) Then fontSize$ = optFont(i%).Caption
    Next
    
    'fontSize$ = "���p"
    Select Case Index
           Case 0: Call currForm.ReportPrnOption(2, fontSize$)
           Case 1: Call currForm.ReportPrnOption(3, fontSize$)
           Case 2: Call currForm.ReportPrnOption(4, fontSize$)
           Case 3: Call currForm.ReportPrnOption(5, fontSize$)
    End Select
    frmRepPreview.wbReport.Refresh

End Sub
Private Sub setDefaultPrinterbyClass(xDeviceName$)
        Dim sMsg As String
        Dim DeviceName As String
    
        If cSetPrinter.SetPrinterAsDefault(xDeviceName$) Then
            sMsg = DeviceName & " has successfully been set as the default printer."
        Else
            sMsg = DeviceName & " has failed to be set as the default printer."
        End If
        'MsgBox sMsg, vbExclamation, App.Title

End Sub

