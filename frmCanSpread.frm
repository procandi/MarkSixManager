VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Begin VB.Form frmCanSpread 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�M�\�d���s��/���"
   ClientHeight    =   11175
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   15270
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11175
   ScaleWidth      =   15270
   StartUpPosition =   2  '�ù�����
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "�R��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6840
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   13
      ToolTipText     =   "�R���ثe������M�\�d��"
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   12
      ToolTipText     =   "�ھڥثe��v�N�����sŪ���M�\�d�����"
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtDr 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   11
      ToolTipText     =   "�ܧ���v��00000�ɡA�i�t�s�s�ɬ��q�νd��"
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "�Ǧ^�ثe���M�\�d��"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11760
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   10
      ToolTipText     =   "�N������M�\�d���Ǧ^"
      Top             =   10200
      Width           =   3375
   End
   Begin MSComDlg.CommonDialog CmnDlg 
      Left            =   0
      Top             =   8520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox lstCanTemplate 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6780
      Left            =   11760
      TabIndex        =   9
      ToolTipText     =   "�����H���J������M�\�d��"
      Top             =   3360
      Width           =   3375
   End
   Begin VB.CommandButton cmdAddCanTemplate 
      BackColor       =   &H00FFC0C0&
      Caption         =   "�t�s�s��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9960
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   7
      ToolTipText     =   "�N�ثe�s�褤���d���s�J�s���M�\�d����"
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdSaveCanTemplate 
      BackColor       =   &H00FFC0C0&
      Caption         =   "�x�s"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8280
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   6
      ToolTipText     =   "�N�ثe�s�褤���d���s�^������M�\�d���ɤ�"
      Top             =   120
      Width           =   1575
   End
   Begin VB.ListBox lstSourceTemplate 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      Left            =   11760
      TabIndex        =   5
      ToolTipText     =   "�����H���J������ťսd��"
      Top             =   480
      Width           =   3375
   End
   Begin FPSpreadADO.fpSpread fpSpread1 
      Height          =   10125
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   11520
      _Version        =   458752
      _ExtentX        =   20320
      _ExtentY        =   17859
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frmCanSpread.frx":0000
   End
   Begin VB.Label lblSpread 
      Height          =   375
      Index           =   3
      Left            =   10800
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblSpread 
      Height          =   375
      Index           =   2
      Left            =   10800
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblSpread 
      Height          =   375
      Index           =   1
      Left            =   10800
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblSpread 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   14
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '�m�����
      BorderStyle     =   1  '��u�T�w
      Caption         =   "�M�\�d��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   11760
      TabIndex        =   8
      Top             =   3000
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  '�m�����
      BorderStyle     =   1  '��u�T�w
      Caption         =   "��l�ťսd��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   11760
      TabIndex        =   4
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label lbType 
      BackColor       =   &H00808080&
      BorderStyle     =   1  '��u�T�w
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   1200
      TabIndex        =   3
      Top             =   480
      Width           =   4485
   End
   Begin VB.Label Label1 
      Alignment       =   2  '�m�����
      BorderStyle     =   1  '��u�T�w
      Caption         =   "���O"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1125
   End
   Begin VB.Label Label1 
      Alignment       =   2  '�m�����
      BorderStyle     =   1  '��u�T�w
      Caption         =   "��v"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1125
   End
End
Attribute VB_Name = "frmCanSpread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xSpreadCanName As String     '�Ω�Ǧ^�s�W�d���W��
Public txtExamType As String        '�Ω���y�\��ϥ�
Public is_Open As Boolean           '�Ω���y�\��ϥ�
Dim xTemplateFile() As String       '�Ω�O����l�d�����
Dim currTemplate(8) As String       '�Ω�O���ثe�ϥΪ��d�����
'0  :Source_filePath        :��Ӫ��d�������|
'1  :SOURCE_FILENAME        :��Ӫ��d�����ɦW
'2  :SOURCE_EXAMNAME        :��Ӫ��d��ExamName
'3  :SOURCE_EXAMID          :��Ӫ��d��ExamID
'4  :SOURCE_EXAMDESCRIPTION :��Ӫ��d��EXAMDescription
'5  :SPREAD_FILENAME        :�M�\�d�����ɮצW��
'6  :SPREAD_NAME            :�M�\�d���W��
'7  :USERID                 :��v�N�X
'8  :template_type          :�ثe�d���ӷ��A0����l�d���A1���M�\�d��
Dim FcurrTemplate As Boolean        '�Ω�O���ثe�d���O�_�ק��w�s
Dim yTemplateFile() As String       '�Ω����USER�M�\�d�����
Dim has_yTemplate As Boolean        '�Ω�O���O�_��USER�M�\�d�����
Dim xCurrForm
'Public lblSpread(3) As String          '�Ω�O�����y���^�Ǧ�m

'�]�wspread��
Public Sub setSPtext(ByVal sheet As Long, ByVal Col As Long, ByVal row As Long, ByVal tvalue As String)
    fpSpread1.sheet = sheet
    fpSpread1.row = row
    fpSpread1.Col = Col
    If Trim(tvalue) = "" Then
        fpSpread1.Text = ""
    Else
        fpSpread1.Text = tvalue
    End If
End Sub

'���ospread������r
Public Function getSPtext(ByVal sheet As Long, ByVal Col As Long, ByVal row As Long) As String
    fpSpread1.sheet = sheet
    fpSpread1.Col = Col
    fpSpread1.row = row
    getSPtext = Trim(fpSpread1.Text)
End Function

Private Sub cmdPhrase_Click()
    Dim t$
    Dim i As Integer
    
    If Val(lblSpread(0)) + Val(lblSpread(1)) > 0 Then
        Set xCurrForm = currForm
        Set currForm = Me
        Load frmPhraseSpread
        Me.Enabled = False
        frmPhraseSpread.Show
        
    End If

End Sub

'�s�W�d���W��
Private Sub cmdAddCanTemplate_Click()
    Dim xflag As Boolean
    Dim sql$
    
    cmdAddCanTemplate.Enabled = False
    xSpreadCanName = ""
    Me.Hide
    frmCanSpreadDialog.Show 1
    Me.Show

    If xSpreadCanName <> "" Then
        xflag = True
        '���ˬd�O�_��USER�M�\�d�����
        If has_yTemplate Then
            '�ˬd�O�_�����ƦW�٪��d��
            For i = 0 To UBound(yTemplateFile)
                If yTemplateFile(i, 5) = xSpreadCanName Then
                    xflag = False
                    Exit For
                End If
            Next
        End If
        
        '�Y�L���ƦW�١A�s�WUSER�d��
        If xflag Then
            currTemplate(6) = xSpreadCanName
            '�s���M�\�d���ɦW��USER�R�W���W��+�~���.rps
            currTemplate(5) = xSpreadCanName & Format(Now, "YYYYMMDDhhmmss") & ".rps"
            currTemplate(7) = Trim(txtDr.Text)
            If Not FSO.FolderExists(currTemplate(0) & currTemplate(7)) Then
                Call FSO.CreateFolder(currTemplate(0) & currTemplate(7))
            End If
            '�M�\�d���O�x�s�bsource_filepath & userid\���ؿ����U
            Call fpSpread1.SaveToFile(currTemplate(0) & currTemplate(7) & "\" & currTemplate(5), False)
            
            sql$ = "Insert into cris_SpreadForm_List ( "
            sql$ = sql$ & " USERID, SPREAD_NAME, SPREAD_FILENAME, EXAM_TYPE, DIVISION_NAME, SOURCE_FILENAME, "
            sql$ = sql$ & " SOURCE_FILEPATH, SOURCE_EXAMNAME, SOURCE_EXAMID, SOURCE_EXAMDESCRIPTION "
            sql$ = sql$ & ") values ( "
            sql$ = sql$ & " '" & currTemplate(7) & "', '" & currTemplate(6) & "', '" & currTemplate(5) & "', "
            sql$ = sql$ & " '" & lbType.Caption & "', '" & Trim(curr_Record.Division_on) & "', "
            sql$ = sql$ & " '" & currTemplate(1) & "', '" & currTemplate(0) & "', '" & currTemplate(2) & "', "
            sql$ = sql$ & " '" & currTemplate(3) & "', '" & currTemplate(4) & "' )"
            Call Connection.Execute(sql$)
            currTemplate(8) = "1"
            '���s���J�M�\�d��
            Call Load_Can_Tempalte
            frmCanSpread.Caption = "�M�\�d���s��/��� - �M�\�d��: " & currTemplate(6)
        Else
            MsgBox "�ɦW���ơA�Хt�~���W"
        End If
        
    End If
    cmdAddCanTemplate.Enabled = True
End Sub

'�s�^USER�M�\�d��
Private Sub cmdSaveCanTemplate_Click()
    cmdSaveCanTemplate.Enabled = False
    '���ˬd�ثe���A�O�_�����JUSER�M�\�d���F�Y����l�d���h���s�W�M�\�d���W��
    If currTemplate(8) = "0" Then
        Call cmdAddCanTemplate_Click
    Else
        '�O�_���ק�L�A���ק�L�ɤ~�ݶץX�ɮסA���ݭק��ƪ�
        If Not FcurrTemplate Then
            '�M�\�d���O�x�s�bsource_filepath & userid\���ؿ����U
            Call fpSpread1.SaveToFile(currTemplate(0) & currTemplate(7) & "\" & currTemplate(5), False)
        End If
    End If
    cmdSaveCanTemplate.Enabled = True
End Sub

'���s���J�M�\�d�����
Private Sub Command1_Click()
    Command1.Enabled = False
    Call Load_Can_Tempalte
    
    Command1.Enabled = True
End Sub

'�R���ثe������d���O��
Private Sub Command2_Click()
    Dim sql$
    
    Command2.Enabled = False
    '�u���M�\�d���~�i�H�R���A��l�d�����i�R��
    If currTemplate(8) = "1" Then
        sql$ = "delete from cris_SpreadForm_List "
        sql$ = sql$ & " where division_name='" & Trim(curr_Record.Division_on) & "' "
        sql$ = sql$ & " and userID = '" & currTemplate(7) & "' "
        sql$ = sql$ & " and exam_type = '" & lbType.Caption & "' "
        sql$ = sql$ & " and Spread_Name = '" & currTemplate(6) & "' "
        Call Connection.Execute(sql$)
        Call Load_Can_Tempalte
        If lstCanTemplate.ListCount > 0 Then
            lstCanTemplate.ListIndex = 0
            Call lstCanTemplate_DblClick
        End If
    End If
    Command2.Enabled = True
End Sub

'�Ǧ^������M�\�d��
Private Sub Command3_Click()
    Dim tPath$
    
    Command3.Enabled = False
    '�N�ثe�}�Ҫ��d���s�Jcurr_record.uni_key+.rps�ɮפ�
    tPath$ = App.Path
    If Right(tPath$, 1) <> "\" Then
        tPath$ = tPath$ & "\"
    End If
    tPath$ = tPath$ & "temp"
    If Not FSO.FolderExists(tPath$) Then
        Call FSO.CreateFolder(tPath$)
    End If
    Call fpSpread1.SaveToFile(tPath$ & "\" & Trim(curr_Record.uni_key) & ".rps", False)
    Spread_ID = currTemplate(3)
    Spread_Name = currTemplate(2)
    curr_Record.TemplateFile = currTemplate(1)
    curr_Record.TemplateName = currTemplate(4)
    Call currForm.fpSpread1.LoadFromFile(tPath$ & "\" & Trim(curr_Record.uni_key) & ".rps")
    '�]�wcurr_form�����A�Pcurr_record���]�w��
    Command3.Enabled = True
    currForm.fpSpread1.ChangeMade = True
    currForm.Enabled = True
    Unload Me
End Sub

'���J��l�d��
Private Sub Load_Source_Tempalte()
    Dim sql$
    Dim rcount As Integer
    Dim i As Integer
    
    '���J��l�d��
    lstSourceTemplate.Clear
    sql$ = "SELECT * FROM CRIS_ReportTemplate "
    sql$ = sql$ & " where divisionname='" & Trim(curr_Record.Division_on) & "' "
    sql$ = sql$ & " ORDER BY DivisionName, ExamID"
    Call OpenRecordset(sql$, Connection, Recordset)
    rcount = 0
    If Not Recordset.EOF Then
        While Not Recordset.EOF
            rcount = rcount + 1
            Recordset.MoveNext
        Wend
        Recordset.MoveFirst
        ReDim xTemplateFile(rcount - 1, 4)
        i = 0
        While Not Recordset.EOF
            If FSO.FileExists(Recordset("templatefilesource") & Recordset("templatefilename")) Then
                lstSourceTemplate.AddItem Recordset("examID") & " - " & Recordset("examname")
                '��Ӫ��d�������|
                xTemplateFile(i, 0) = Recordset("templatefilesource")
                '��Ӫ��d��ExamID
                xTemplateFile(i, 1) = Recordset("ExamID")
                '��Ӫ��d��ExamName
                xTemplateFile(i, 2) = Recordset("ExamName")
                '��Ӫ��d��EXAMDescription
                xTemplateFile(i, 3) = Recordset("ExamDescription")
                '��Ӫ��d�����ɦW
                xTemplateFile(i, 4) = Recordset("templatefilename")
                i = i + 1
            End If
            Recordset.MoveNext
        Wend
    End If
End Sub

'���J�M�\�d��
Private Sub Load_Can_Tempalte()
    Dim sql$
    Dim rcount As Integer
    Dim i As Integer
    
    lstCanTemplate.Clear
    has_yTemplate = False
    sql$ = "SELECT * FROM cris_SpreadForm_List "
    sql$ = sql$ & " where division_name='" & Trim(curr_Record.Division_on) & "' "
    sql$ = sql$ & " and (userID = '" & Trim(txtDr.Text) & "' or userID = '00000') "
    sql$ = sql$ & " and exam_type = '" & lbType.Caption & "' "
    sql$ = sql$ & " ORDER BY Division_Name, userID, Spread_Name"
    Call OpenRecordset(sql$, Connection, Recordset)
    rcount = 0
    If Not Recordset.EOF Then
        While Not Recordset.EOF
            rcount = rcount + 1
            Recordset.MoveNext
        Wend
        Recordset.MoveFirst
        ReDim yTemplateFile(rcount - 1, 9)
        i = 0
        While Not Recordset.EOF
            'USER�N�X
            yTemplateFile(i, 0) = NoNull(Recordset("UserID"))
            '��Ӫ��d��EXAMDescription
            yTemplateFile(i, 1) = NoNull(Recordset("Source_ExamDescription"))
            '��Ӫ��d��ExamName
            yTemplateFile(i, 2) = NoNull(Recordset("Source_ExamName"))
            '��Ӫ��d��ExamID
            yTemplateFile(i, 3) = NoNull(Recordset("Source_ExamID"))
            '��Ӫ��d�����ɦW
            yTemplateFile(i, 4) = NoNull(Recordset("Source_FileName"))
            '��Ӫ��d�����ɮ׸��|
            yTemplateFile(i, 5) = NoNull(Recordset("Source_FilePath"))
            '��ǦW��
            yTemplateFile(i, 6) = NoNull(Recordset("Division_Name"))
            '�ˬd���O
            yTemplateFile(i, 7) = NoNull(Recordset("Exam_Type"))
            '�M�\�d���W��
            yTemplateFile(i, 8) = NoNull(Recordset("Spread_Name"))
            '�M�\�d�����ɮצW��
            yTemplateFile(i, 9) = NoNull(Recordset("Spread_FileName"))
            lstCanTemplate.AddItem yTemplateFile(i, 0) & " - " & yTemplateFile(i, 8)
            i = i + 1
            Recordset.MoveNext
        Wend
        has_yTemplate = True
    End If
End Sub

Private Sub Form_Activate()
    Set currForm = xCurrForm
    
End Sub

Private Sub Form_Load()
    Dim sql$
    Dim rcount As Integer
    Dim i As Integer
    
    Set xCurrForm = currForm
    txtDr.Text = frmQueue.lblUser.Caption
    lbType.Caption = curr_Record.Type
    txtExamType = curr_Record.Type
    is_Open = False
    
    '���J��l�d��
    Call Load_Source_Tempalte
    
    '���JUSER���M�\�d��
    Call Load_Can_Tempalte
            
    '���J�w�]����l�d���P�]�w���
    If lstSourceTemplate.ListCount > 0 Then
        lstSourceTemplate.ListIndex = 0
        Call lstSourceTemplate_DblClick
    ElseIf lstCanTemplate.ListCount > 0 Then
        lstCanTemplate.ListIndex = 0
        Call lstCanTemplate_DblClick
    End If

    FcurrTemplate = True
    Command3.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    currForm.Enabled = True
    currForm.Show
    Unload Me
End Sub

Private Sub fpSpread1_Change(ByVal Col As Long, ByVal row As Long)
    FcurrTemplate = False
End Sub

Private Sub fpSpread1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    Dim tmpTxtRowCol$, txtCol&, txtRow&
           
    fpSpread1.Col = Col
    fpSpread1.row = row
    fpSpread1.sheet = fpSpread1.ActiveSheet
    tmpTxtRowCol$ = fpSpread1.CellTag
    If Len(tmpTxtRowCol$) = 4 Then
       txtRow& = Val(Left(tmpTxtRowCol$, 2))
       txtCol& = Val(Right(tmpTxtRowCol$, 2))
        
       fpSpread1.Col = txtCol&
       fpSpread1.row = txtRow&
       
       lblSpread(0) = fpSpread1.row
       lblSpread(1) = fpSpread1.Col
       lblSpread(2) = fpSpread1.CellNote
       'lblSpread(3) = fpSpread1.SelStart
    Else
       lblSpread(0) = ""
       lblSpread(1) = ""
       lblSpread(2) = ""
       'lblSpread(3) = ""
    
    End If
        '--------------------------------------------------------------

    If fpSpread1.CellType = CellTypeEdit Then
       DoEvents
       
       Call cmdPhrase_Click
    End If
End Sub

'USER�����M�\�d��
Private Sub lstCanTemplate_DblClick()
    Dim temp$
    
    lstCanTemplate.Enabled = False
    '���J�M�\�d��
    If lstCanTemplate.ListIndex >= 0 Then
        If FSO.FileExists(yTemplateFile(lstCanTemplate.ListIndex, 5) & yTemplateFile(lstCanTemplate.ListIndex, 0) & "\" & yTemplateFile(lstCanTemplate.ListIndex, 9)) Then
            Call fpSpread1.LoadFromFile(yTemplateFile(lstCanTemplate.ListIndex, 5) & yTemplateFile(lstCanTemplate.ListIndex, 0) & "\" & yTemplateFile(lstCanTemplate.ListIndex, 9))
            fpSpread1.ChangeMade = True
            fpSpread1.ColHeadersShow = False
            fpSpread1.RowHeadersShow = False
            '�]�w���
            currTemplate(0) = yTemplateFile(lstCanTemplate.ListIndex, 5)
            currTemplate(1) = yTemplateFile(lstCanTemplate.ListIndex, 4)
            currTemplate(2) = yTemplateFile(lstCanTemplate.ListIndex, 2)
            currTemplate(3) = yTemplateFile(lstCanTemplate.ListIndex, 3)
            currTemplate(4) = yTemplateFile(lstCanTemplate.ListIndex, 1)
            currTemplate(5) = yTemplateFile(lstCanTemplate.ListIndex, 9)
            currTemplate(6) = yTemplateFile(lstCanTemplate.ListIndex, 8)
            currTemplate(7) = yTemplateFile(lstCanTemplate.ListIndex, 0)
            currTemplate(8) = "1"
            frmCanSpread.Caption = "�M�\�d���s��/��� - �M�\�d��: " & currTemplate(6)
            Command2.Enabled = True
            FcurrTemplate = True
        Else
            MsgBox "�䤣��M�\�d���ɮסA���p���a�Q�u�{�v!"
            temp$ = "�򥢮M�\�d���ɮ�: "
            temp$ = temp$ & yTemplateFile(lstCanTemplate.ListIndex, 5) & yTemplateFile(lstCanTemplate.ListIndex, 0) & "\" & yTemplateFile(lstCanTemplate.ListIndex, 9)
            Log (temp$)
        End If
    End If
    lstCanTemplate.Enabled = True
End Sub

'USER������l�d��
Private Sub lstSourceTemplate_DblClick()
    lstSourceTemplate.Enabled = False
    '���J��l�d��
    If lstSourceTemplate.ListIndex >= 0 Then
        If FSO.FileExists(xTemplateFile(lstSourceTemplate.ListIndex, 0) & xTemplateFile(lstSourceTemplate.ListIndex, 4)) Then
            Call fpSpread1.LoadFromFile(xTemplateFile(lstSourceTemplate.ListIndex, 0) & xTemplateFile(lstSourceTemplate.ListIndex, 4))
            fpSpread1.ChangeMade = True
            fpSpread1.ColHeadersShow = False
            fpSpread1.RowHeadersShow = False
            '�]�w���
            currTemplate(0) = xTemplateFile(lstSourceTemplate.ListIndex, 0)
            currTemplate(1) = xTemplateFile(lstSourceTemplate.ListIndex, 4)
            currTemplate(2) = xTemplateFile(lstSourceTemplate.ListIndex, 2)
            currTemplate(3) = xTemplateFile(lstSourceTemplate.ListIndex, 1)
            currTemplate(4) = xTemplateFile(lstSourceTemplate.ListIndex, 3)
            currTemplate(5) = ""
            currTemplate(6) = ""
            currTemplate(7) = Trim(txtDr.Text)
            currTemplate(8) = "0"
            frmCanSpread.Caption = "�M�\�d���s��/��� - ��l�d��: " & currTemplate(3) & " - " & currTemplate(2)
            FcurrTemplate = True
            '��l�d�����i�R��
            Command2.Enabled = False
        Else
            MsgBox "�䤣���l�d���ɮסA���p���a�Q�u�{�v!"
            temp$ = "�򥢭�l�d���ɮ�: "
            temp$ = temp$ & xTemplateFile(lstSourceTemplate.ListIndex, 0) & xTemplateFile(lstSourceTemplate.ListIndex, 4)
            Log (temp$)
        End If
    End If
    lstSourceTemplate.Enabled = True
End Sub
