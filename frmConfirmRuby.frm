VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmConfirmRuby 
   BorderStyle     =   1  '��u�T�w
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
         Name            =   "�s�ө���"
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
      Caption         =   "&N ���@��"
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
      Left            =   2400
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   2
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton cmdConfirm 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&O �T�@�w"
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
      Left            =   0
      Style           =   1  '�Ϥ��~�[
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
         Name            =   "�s�ө���"
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
      Alignment       =   1  '�a�k���
      BorderStyle     =   1  '��u�T�w
      Caption         =   "���~�W��"
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
      Index           =   2
      Left            =   720
      TabIndex        =   9
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblEntry 
      Alignment       =   1  '�a�k���
      BorderStyle     =   1  '��u�T�w
      Caption         =   "�Ȥ�W��"
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
      Index           =   0
      Left            =   720
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  '�m�����
      BackColor       =   &H00808080&
      BackStyle       =   0  '�z��
      Caption         =   "�����(�`�b)�H��w�����ѳ����X�C�g�B��B�~����(�`�b)�H��w�����Ѫ���g�B��B�~�����X�C"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   1  '�a�k���
      BorderStyle     =   1  '��u�T�w
      Caption         =   "������"
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
      Index           =   1
      Left            =   720
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  '�m�����
      BackColor       =   &H00808080&
      BackStyle       =   0  '�z��
      Caption         =   "����C�L"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H00808080&
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
        MsgBox "�Х���ܭn�C�L���ɶ��I"
    ElseIf (basVariable.Parameter = "CustomProductDayReport" Or basVariable.Parameter = "CustomProductWeekReport") And cmbCName.Text = "" And cmbPName.Text = "" Then
        MsgBox "�|����ܫȤ�β��~�I"
    ElseIf (basVariable.Parameter = "CustomWeekReport" Or basVariable.Parameter = "CustomMonthReport" Or basVariable.Parameter = "CustomYearReport") And cmbCName.Text = "" Then
        MsgBox "�|����ܫȤ�I"
    Else
        Dim CData() As String
        Dim PData() As String
            
        Select Case basVariable.Parameter
        Case "CustomDailyTransactionDetail"
            'Label1(0).Caption = "�Ȥ�C��������"
            
            CData = Split(cmbCName.Text, " ")
            PData = Split(cmbPName.Text, " ")
            Call Shell("C:\Ruby22-x64\bin\ruby.exe -Ku main.rb" & " DailyTransactionCounting " & PData(0) & " " & CData(0) & " " & txtCurrentDate.Text)
            
        Case "DailyTransactionCounting"
            'Label1(0).Caption = "�C�����[�`��"
            
            PData = Split(cmbPName.Text, " ")
            Call Shell("C:\Ruby22-x64\bin\ruby.exe -Ku main.rb" & " DailyTransactionCounting " & PData(0) & " " & txtCurrentDate.Text)
            
        Case "AllDailyTransactionCounting"
            'Label1(0).Caption = "�����~�C�����[�`��"
            
            Call Shell("C:\Ruby22-x64\bin\ruby.exe -Ku main.rb" & " AllDailyTransactionCounting " & txtCurrentDate.Text)
            
        Case "AllWeekTransactionCounting"
            'Label1(0).Caption = "�����~�@�g����[�`��"
            
            Call Shell("C:\Ruby22-x64\bin\ruby.exe -Ku main.rb" & " AllWeekTransactionCounting " & txtCurrentDate.Text)
            
        Case "AllMonthTransactionCounting"
            'Label1(0).Caption = "�����~������[�`��"
            
            Call Shell("C:\Ruby22-x64\bin\ruby.exe -Ku main.rb" & " AllMonthTransactionCounting " & txtCurrentDate.Text)
            
        Case "MonthTransactionCounting"
            'Label1(0).Caption = "������[�`��"
            
            PData = Split(cmbPName.Text, " ")
            Call Shell("C:\Ruby22-x64\bin\ruby.exe -Ku main.rb" & " MonthTransactionCounting " & PData(0) & " " & txtCurrentDate.Text)
            
        Case "AllDaily4KTransactionCounting"
            'Label1(0).Caption = "�����~4K�C�����[�`��"
            
            Call Shell("C:\Ruby22-x64\bin\ruby.exe -Ku main.rb" & " AllDaily4KTransactionCounting " & txtCurrentDate.Text)
            
        Case "AllMonth4KTransactionCounting"
            'Label1(0).Caption = "�����~4K������[�`��"
            
            Call Shell("C:\Ruby22-x64\bin\ruby.exe -Ku main.rb" & " AllMonth4KTransactionCounting " & txtCurrentDate.Text)
            
        Case "CustomDailyPriceDetail"
            'Label1(0).Caption = "�Ȥ�C���������"
            
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
        Label1(0).Caption = "�Ȥ�C��������"
        
        lblEntry(0).Visible = True
        lblEntry(2).Visible = True
        cmbCName.Visible = True
        cmbPName.Visible = True
        Call ComboBox_LoadFrom_DataBase_ByFile(cmbCName, "CID,CName", "custom", "", "", "")
        
        Call cmbPName.AddItem("100 539_��")
        Call cmbPName.AddItem("110 �丹_��")
        Call cmbPName.AddItem("120 �j�ֳz_��")
        
    Case "DailyTransactionCounting"
        Label1(0).Caption = "�C�����[�`��"
        
        lblEntry(2).Visible = True
        cmbPName.Visible = True
        
        Call cmbPName.AddItem("100 539_��")
        Call cmbPName.AddItem("110 �丹_��")
        Call cmbPName.AddItem("120 �j�ֳz_��")
        
    Case "AllDailyTransactionCounting"
        Label1(0).Caption = "�����~�C�����[�`��"
        
    Case "AllWeekTransactionCounting"
        Label1(0).Caption = "�����~�@�g����[�`��"
        
    Case "AllMonthTransactionCounting"
        Label1(0).Caption = "�����~������[�`��"
        
    Case "MonthTransactionCounting"
        Label1(0).Caption = "������[�`��"
        
        lblEntry(2).Visible = True
        cmbPName.Visible = True
        
        Call cmbPName.AddItem("100 539_��")
        Call cmbPName.AddItem("110 �丹_��")
        Call cmbPName.AddItem("120 �j�ֳz_��")
        
    Case "AllDaily4KTransactionCounting"
        Label1(0).Caption = "�����~4K�C�����[�`��"
        
    Case "AllMonth4KTransactionCounting"
        Label1(0).Caption = "�����~4K������[�`��"
        
    Case "CustomDailyPriceDetail"
        Label1(0).Caption = "�Ȥ�C���������"

    End Select
    
    
    dtpCurrentDate.Value = Format(DateTime.Now, "yyyy/MM/dd")
    txtCurrentDate.Text = Format(DateTime.Now, "yyyy/MM/dd")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmProveNew.Show
    Unload Me
End Sub
