VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmConfirm 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "Report Confirm"
   ClientHeight    =   3105
   ClientLeft      =   12510
   ClientTop       =   7530
   ClientWidth     =   4680
   Icon            =   "frmConfirm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4680
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
      Top             =   840
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
      Top             =   2520
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
      Left            =   120
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   1
      Top             =   2520
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker dtpCurrentDate 
      Height          =   360
      Left            =   1800
      TabIndex        =   5
      Top             =   840
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
      Format          =   94109699
      CurrentDate     =   37058
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
      Top             =   1440
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
      Top             =   840
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
      Height          =   2295
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Call Form_Unload(0)
End Sub

Sub DayReport(ByVal TargetPath As String)
    Dim Body As String
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    
    SQL = "select * from product;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    SQL = "select * from price where CID='" & basVariable.SelectCID & "' and CurrentDate='" & Format(DateTime.Now, "yyyy/MM/dd") & "';"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
    
    
    Open TargetPath For Output As #1
        Print #1, "���", vbTab & Format(DateTime.Now, "yyyy/MM/dd")
        
        Body = "���~"
        Print #1, Body
    Close #1
End Sub

Sub WeekReport(ByVal TargetPath As String)

End Sub

Sub MonthReport(ByVal TargetPath As String)

End Sub

Sub YearReport(ByVal TargetPath As String)

End Sub

Sub DayAccount(ByVal TargetPath As String)

End Sub

Sub WeekAccount(ByVal TargetPath As String)

End Sub

Sub MonthAccount(ByVal TargetPath As String)

End Sub

Sub YearAccount(ByVal TargetPath As String)

End Sub

Sub FourKDayReport(ByVal TargetPath As String)

End Sub

Sub FourKWeekReport(ByVal TargetPath As String)

End Sub

Sub FourKMonthReport(ByVal TargetPath As String)

End Sub

Sub FourKYearReport(ByVal TargetPath As String)

End Sub

Sub FourKDayAccount(ByVal TargetPath As String)

End Sub

Sub FourKWeekAccount(ByVal TargetPath As String)

End Sub

Sub FourKMonthAccount(ByVal TargetPath As String)

End Sub

Sub FourKYearAccount(ByVal TargetPath As String)

End Sub

Private Sub cmdConfirm_Click()
    If txtCurrentDate.Text = "" Then
        MsgBox "�Х���ܭn�C�L���ɶ��I"
    Else
        Dim TargetPath As String
        
        TargetPath = App.Path
        If Right(TargetPath, 1) <> "\" Then
            TargetPath = TargetPath & "\report\"
        Else
            TargetPath = TargetPath & "report\"
        End If
        Call CreatePath(TargetPath)
        
        Select Case basVariable.Parameter
        Case "DayReport"
            TargetPath = TargetPath & Format(DateTime.Now, "yyyyMMdd") & "_�����.xls"
            Call DayReport(TargetPath)
        Case "WeekReport"
            TargetPath = TargetPath & Format(DateTime.Now, "yyyyMMdd") & "_�g����.xls"
            Call WeekReport(TargetPath)
        Case "MonthReport"
            TargetPath = TargetPath & Format(DateTime.Now, "yyyyMMdd") & "_�����.xls"
            Call MonthReport(TargetPath)
        Case "YearReport"
            TargetPath = TargetPath & Format(DateTime.Now, "yyyyMMdd") & "_�~����.xls"
            Call YearReport(TargetPath)
        Case "DayAccount"
            TargetPath = TargetPath & Format(DateTime.Now, "yyyyMMdd") & "_���`��.xls"
            Call DayAccount(TargetPath)
        Case "WeekAccount"
            TargetPath = TargetPath & Format(DateTime.Now, "yyyyMMdd") & "_�g�`��.xls"
            Call WeekAccount(TargetPath)
        Case "MonthAccount"
            TargetPath = TargetPath & Format(DateTime.Now, "yyyyMMdd") & "_���`��.xls"
            Call MonthAccount(TargetPath)
        Case "YearAccount"
            TargetPath = TargetPath & Format(DateTime.Now, "yyyyMMdd") & "_�~�`��.xls"
            Call YearAccount(TargetPath)
        Case "FourKDayReport"
            TargetPath = TargetPath & Format(DateTime.Now, "yyyyMMdd") & "_4K�����.xls"
            Call FourKDayReport(TargetPath)
        Case "FourKWeekReport"
            TargetPath = TargetPath & Format(DateTime.Now, "yyyyMMdd") & "_4K�g����.xls"
            Call FourKWeekReport(TargetPath)
        Case "FourKMonthReport"
            TargetPath = TargetPath & Format(DateTime.Now, "yyyyMMdd") & "_4K�����.xls"
            Call FourKMonthReport(TargetPath)
        Case "FourKYearReport"
            TargetPath = TargetPath & Format(DateTime.Now, "yyyyMMdd") & "_4K�~����.xls"
            Call FourKYearReport(TargetPath)
        Case "FourKDayAccount"
            TargetPath = TargetPath & Format(DateTime.Now, "yyyyMMdd") & "_4K���`��.xls"
            Call FourKDayAccount(TargetPath)
        Case "FourKWeekAccount"
            TargetPath = TargetPath & Format(DateTime.Now, "yyyyMMdd") & "_4K�g�`��.xls"
            Call FourKWeekAccount(TargetPath)
        Case "FourKMonthAccount"
            TargetPath = TargetPath & Format(DateTime.Now, "yyyyMMdd") & "_4K���`��.xls"
            Call FourKMonthAccount(TargetPath)
        Case "FourKYearAccount"
            TargetPath = TargetPath & Format(DateTime.Now, "yyyyMMdd") & "_4K�~�`��.xls"
            Call FourKYearAccount(TargetPath)
        End Select
    End If
    
    MsgBox "�w��X�����" & TargetPath & "�I"
End Sub

Private Sub dtpCurrentDate_CloseUp()
    txtCurrentDate.Text = Format(dtpCurrentDate.Value, "yyyy/MM/dd")
End Sub

Private Sub Form_Load()
    Select Case basVariable.Parameter
    Case "DayReport"
        Label1(0).Caption = "�����C�L"
    Case "WeekReport"
        Label1(0).Caption = "�g����C�L"
    Case "MonthReport"
        Label1(0).Caption = "�����C�L"
    Case "YearReport"
        Label1(0).Caption = "�~����C�L"
    Case "DayAccount"
        Label1(0).Caption = "���`��C�L"
    Case "WeekAccount"
        Label1(0).Caption = "�g�`��C�L"
    Case "MonthAccount"
        Label1(0).Caption = "���`��C�L"
    Case "YearAccount"
        Label1(0).Caption = "�~�`��C�L"
    Case "FourKDayReport"
        Label1(0).Caption = "4K�����C�L"
    Case "FourKWeekReport"
        Label1(0).Caption = "4K�g����C�L"
    Case "FourKMonthReport"
        Label1(0).Caption = "4K�����C�L"
    Case "FourKYearReport"
        Label1(0).Caption = "4K�~����C�L"
    Case "FourKDayAccount"
        Label1(0).Caption = "4K���`��C�L"
    Case "FourKWeekAccount"
        Label1(0).Caption = "4K�g�`��C�L"
    Case "FourKMonthAccount"
        Label1(0).Caption = "4K���`��C�L"
    Case "FourKYearAccount"
        Label1(0).Caption = "4K�~�`��C�L"
    End Select
    
    
    dtpCurrentDate.Value = Format(DateTime.Now, "yyyy/MM/dd")
    txtCurrentDate.Text = Format(DateTime.Now, "yyyy/MM/dd")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmProve.Show
    Unload Me
End Sub
