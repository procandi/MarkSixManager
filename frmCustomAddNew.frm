VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCustomAddNew 
   BackColor       =   &H00808080&
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�Ȥ��Ʃ���"
   ClientHeight    =   6270
   ClientLeft      =   2160
   ClientTop       =   3660
   ClientWidth     =   8775
   Icon            =   "frmCustomAddNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   8775
   Begin Threed.SSPanel pnlRegist 
      Height          =   5535
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   8415
      _Version        =   65536
      _ExtentX        =   14843
      _ExtentY        =   9763
      _StockProps     =   15
      ForeColor       =   0
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Outline         =   -1  'True
      FloodColor      =   0
      Alignment       =   6
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5520
         MaxLength       =   256
         TabIndex        =   27
         Top             =   2520
         Width           =   2655
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5520
         MaxLength       =   256
         TabIndex        =   26
         Top             =   2160
         Width           =   2655
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5520
         MaxLength       =   256
         TabIndex        =   25
         Top             =   1800
         Width           =   2655
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5520
         MaxLength       =   256
         TabIndex        =   24
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5520
         MaxLength       =   256
         TabIndex        =   23
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5520
         MaxLength       =   256
         TabIndex        =   22
         Top             =   720
         Width           =   2655
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   1320
         TabIndex        =   21
         Top             =   1560
         Width           =   3015
      End
      Begin VB.ComboBox cmbDiagnosisClassM 
         Height          =   300
         Left            =   1320
         TabIndex        =   20
         Top             =   1920
         Width           =   3015
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   240
         MultiSelect     =   1  '²���h�����
         TabIndex        =   19
         Top             =   3240
         Width           =   7935
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1320
         MaxLength       =   256
         TabIndex        =   18
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5520
         MaxLength       =   256
         TabIndex        =   17
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&U �T�w�ק�"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   3
         Tag             =   "Edit"
         Top             =   4920
         Width           =   2535
      End
      Begin VB.ListBox lstExamDetail 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   240
         MultiSelect     =   1  '²���h�����
         TabIndex        =   1
         Top             =   4080
         Width           =   7935
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&X ����"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5520
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   4
         Top             =   4920
         Width           =   2655
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&R �T�w�s�W"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   2
         Tag             =   "Insert"
         Top             =   4920
         Width           =   2535
      End
      Begin VB.TextBox txtAccessionNo 
         Enabled         =   0   'False
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
         Left            =   1440
         TabIndex        =   13
         Top             =   240
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker dtpOrderDate 
         Height          =   375
         Left            =   2760
         TabIndex        =   0
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy/MM/dd"
         Format          =   91422723
         CurrentDate     =   42267
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '�a�k���
         BackColor       =   &H00808080&
         BorderStyle     =   1  '��u�T�w
         Caption         =   "�Ȧ�b��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   10
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '�a�k���
         BackColor       =   &H00808080&
         BorderStyle     =   1  '��u�T�w
         Caption         =   "�q��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   2
         Left            =   4440
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblBasic 
         Alignment       =   1  '�a�k���
         BackColor       =   &H00808080&
         BorderStyle     =   1  '��u�T�w
         Caption         =   "�ȧO���O"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblRegist 
         Alignment       =   1  '�a�k���
         BackColor       =   &H00808080&
         BorderStyle     =   1  '��u�T�w
         Caption         =   "�}����"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   2505
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '�a�k���
         BackColor       =   &H00808080&
         BorderStyle     =   1  '��u�T�w
         Caption         =   "�Ƶ�"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Index           =   8
         Left            =   240
         TabIndex        =   11
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '�a�k���
         BackColor       =   &H00808080&
         BorderStyle     =   1  '��u�T�w
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   6
         Left            =   240
         TabIndex        =   10
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '�a�k���
         BackColor       =   &H00808080&
         BorderStyle     =   1  '��u�T�w
         Caption         =   "�m�W"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   5
         Left            =   4440
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '�a�k���
         BackColor       =   &H00808080&
         BorderStyle     =   1  '��u�T�w
         Caption         =   "�a�}"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '�a�k���
         BackColor       =   &H00808080&
         BorderStyle     =   1  '��u�T�w
         Caption         =   "�h��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   4
         Left            =   240
         TabIndex        =   7
         Top             =   1920
         Width           =   1095
      End
   End
   Begin MSAdodcLib.Adodc datBasic 
      Height          =   495
      Left            =   120
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   2
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
      LockType        =   1
      CommandType     =   1
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
      DataSourceName  =   "EndoSVR"
      OtherAttributes =   ""
      UserName        =   "alantso"
      Password        =   "5682"
      RecordSource    =   ""
      Caption         =   "datBasic"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      BackStyle       =   0  '�z��
      Caption         =   "�Ȥ��Ʃ���"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   24
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   1
      Left            =   5280
      TabIndex        =   16
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmCustomAddNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Call Form_Unload(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmCustom.Show
    Unload Me
End Sub

