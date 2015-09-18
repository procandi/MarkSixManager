VERSION 5.00
Begin VB.Form frmDoctorConfirm 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "確認報告醫師"
   ClientHeight    =   5790
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  '暫止
      Left            =   4080
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   720
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   2160
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   727
      Width           =   2895
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      TabIndex        =   1
      Top             =   4800
      Width           =   2535
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "確認"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   4800
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H0080FFFF&
      Caption         =   "請確定報告內容是否已完整"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   3840
      Width           =   5535
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "請點選醫師姓名及輸入密碼"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   5535
   End
   Begin VB.Label Label1 
      Caption         =   "密碼 :"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3120
      TabIndex        =   8
      Top             =   780
      Width           =   975
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3360
      TabIndex        =   5
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "確認報告醫師姓名 :"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "打報告醫師姓名 :"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmDoctorConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'此常數用於定義下拉選單的醫師的科別
Const system_type = "腸胃科"

Option Explicit

Private Sub CancelButton_Click()
    frmDoctorConfirm.Visible = False
    frmSpread.Show
    Unload frmDoctorConfirm
End Sub

Private Sub Form_Load()
    'Call ComboBox_LoadFrom_DataBase(Combo1, "Name", "CRIS_USER with(nolock)", "System='" & system_type & "'", "", EXAMSVR_INI)
    Call cmb_DR_Initial(db_Name$, "CRIS_User", "Name", "System='" & system_type & "'", Combo1)
    Call cmb_DR_Initial(db_Name$, "CRIS_User", "Name", "System='" & system_type & "'", Combo2)
    'Call ComboBox_LoadFrom_DataBase(Combo2, "Name", "CRIS_USER with(nolock)", "System='" & system_type & "'", "", EXAMSVR_INI)
    Combo1.Text = curr_Record.Dr_on
    Combo2.Text = curr_Record.Dr_report
End Sub

Private Sub OKButton_Click()
    If Password_Get(Combo1, Text1.Text) Then
        curr_Record.Dr_on = Combo1.Text
        curr_Record.Dr_report = Combo2.Text
        
        Dim SQL$, dbS As New adoDB.Connection

        dbS.Open dbConnection$
        SQL$ = "UPDATE CRIS_Exam_Online SET " & _
               "Dr_On='" & curr_Record.Dr_on & "', " & _
               "Dr_report='" & curr_Record.Dr_report & "', " & _
               "item6='" & frmSpread.Temp_item6 & "' " & _
               "WHERE ChartNo='" & curr_Record.ChartNo & "' and Uni_Key='" & curr_Record.Uni_key & "'"
        dbS.Execute SQL$
        
        dbS.Close
        Set dbS = Nothing
        
        frmDoctorConfirm.Visible = False
        frmPatientTrack.Show
        Unload frmDoctorConfirm
    Else
        MsgBox "密碼錯誤，請注意大小寫與全半形的不同!"
    End If
End Sub

'取得密碼與比對，比對正確傳回true，錯誤傳回false
Private Function Password_Get(dr_name As String, dr_pass As String) As Boolean
    Dim SQL$, conn$, rec$, Exam_Type$
    Dim adoDB As New adoDB.Connection
    Dim adoCode As New adoDB.Recordset
    Dim adoClass As New adoDB.Recordset
    Dim result As Boolean
    
    result = False
    adoDB.Open dbConnection$
    
    SQL$ = " UserID+Name='" & dr_name & "' "
    '/**/
    'rec$ = "SELECT DISTINCT Name, UserID, Password, Type, System, Phone FROM CRIS_User WHERE " & SQL$
    '/**/
    rec$ = "SELECT DISTINCT Name, UserID, Password, Type, System, Phone FROM CRIS_User with(nolock) WHERE " & SQL$
    '/**/
    adoCode.Open rec$, adoDB, adOpenForwardOnly, adLockReadOnly
    If Not adoCode.EOF Then
        If dr_pass = adoCode!Password Then
            result = True
        End If
    End If
    
    adoCode.Close
    adoDB.Close
    Set adoCode = Nothing
    Set adoDB = Nothing
    Password_Get = result
End Function
