Attribute VB_Name = "basUserInfo"
'/******************************************************************/
'/*�����G�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*��m��ϥΪ̭ӤH�]�w�B��Ʀ������a��C                          */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�ҲաG�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@    �@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�ѦҡG    �@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*msscript.dll�C                                                  */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*����G    �@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*FPSPR70.OCX�C                                                   */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�`�N�ƶ��G�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*                                           Edit By Edward Hsieh */
'/*                                      Last Edit Date 2009/02/26 */
'/******************************************************************/
Option Explicit



'/**************************���Ϊ��n�J�ϥΪ̱`�Ƹ��***********************************/

'/**/
Public Const CHARTNO_LENGTH As Integer = 10 '���t�Ϊ��f��������
Public Const EMERGENCY_ID As String = "55555" '�ߤ�²�T�|�Ψ쪺�A���H²�T���N�X
'/**/


'/*�]��ni_svr_queue�w�g���ӥi��ϥΨ�F�A�G�o��������������*/
'Public Const NI_REPORT_TEMPLATE_PATH_A As String = "./JS_Template_A.txt" '�V�W�Ψ�A�����Ҫ����|
'Public Const NI_REPORT_TEMPLATE_PATH_B As String = "./JS_Template_B.txt" '�V�W�Ψ�B�����Ҫ����|
'Public Const NI_REPORT_TEMPLATE_PATH_C As String = "./JS_Template_C.txt" '�V�W�Ψ�C�����Ҫ����|
'Public Const NI_REPORT_TEMPLATE_PATH_D As String = "./JS_Template_D.txt" '�V�W�Ψ�D�����Ҫ����|
'Public Const NI_REPORT_TEMPLATE_PATH_E As String = "./JS_Template_E.txt" '�V�W�Ψ�E�����Ҫ����|
'Public Const NI_REPORT_TEMPLATE_PATH_F As String = "./JS_Template_F.txt" '�V�W�Ψ�F�����Ҫ����|
'Public Const MAX_NI_FIELD_A As Integer = 50 '�V�W�Ψ�A�������ƶq
'Public Const MAX_NI_FIELD_B As Integer = 48 '�V�W�Ψ�B�������ƶq
'Public Const MID_NI_FIELD_C_AND_D As Integer = 28 '�V�W�Ψ�C�BD�������ƶq
'Public Const MAX_NI_FIELD_E As Integer = 16 '�V�W�Ψ�E�������ƶq
'Public Const MAX_NI_FIELD_F As Integer = 28 '�V�W�Ψ�F�������ƶq
'/**/


'/**/
Public Const NI_REPORT_TEMPLATE_PATH_NECK_TCI As String = "./NI_Template_NECK_TCI.ini" '�V�W�Ψ�NECK_TCI�����Ҫ����|
Public Const NI_REPORT_TEMPLATE_PATH_NECK As String = "./NI_Template_NECK.ini" '�V�W�Ψ�NECK�����Ҫ����|
Public Const NI_REPORT_TEMPLATE_PATH_TCI As String = "./NI_Template_TCI.ini" '�V�W�Ψ�TCI�����Ҫ����|
Public Const NI_REPORT_TEMPLATE_PATH_LimpUpper As String = "./NI_Template_LimpUpper.ini" '�V�W�Ψ�LimpUpper�����Ҫ����|
Public Const NI_REPORT_TEMPLATE_PATH_LimpLower As String = "./NI_Template_LimpLower.ini" '�V�W�Ψ�LimpLower�����Ҫ����|
'/**/


'/*��ini�B�z�������`��*/
Public Const EXAMSVR_INI As String = "./ExamSVR.ini" '�O�����Ω�s�u��HIS���ƪ��s���r�굥��ƪ�ini�ɸ��|(�����)
Public Const EXAMHCR_INI As String = "./ExamHCR.ini" '�O�����Ω�s�u��HIS���ƪ��s���r�굥��ƪ�ini�ɸ��|(���������)
Public Const SCHEDULE_PRINTER_INI As String = "./Schedule_Printer.ini" '�O��Schdule�ӵ{���A�Ҧ������L������]�w��ƪ�ini�ɸ��|(��������)
Public Const TIF_MDI_2_JPG_INI As String = "./TIF_MDI_2_JPG.ini" '�O��TIF_MDI_2_JPG�ӵ{���A�Ҧ��w�]���ɮצs�����|�A�H�έn�����dialog�������W�٦b���]�w(��������)
Public Const TIF_MDI_2_DCM_INI As String = "./TIF_MDI_2_DCM.ini" '�O��TIF_MDI_2_DCM�ӵ{���A�Ҧ��w�]���ɮצs�����|�A�H�έn�����dialog�������W�٦b���]�w(��������)
Public Const PDF_2_JPG_INI As String = "./PDF_2_JPG.ini" '�O��PDF2JPG�ӵ{���A�Ҧ��w�]���ɮצs�����|�A�H�έn�����dialog�������W�٦b���]�w(��������)
Public Const SVR_HC_PDF_2_JPG_INI As String = "./SVR_HC_PDF_2_JPG.ini" '�O��SVR_HC_PDF2JPG�ӵ{���A�Ҧ��w�]���ɮצs�����|�A�H�έn�����dialog�������W�٦b���]�w(�s�˪���)
Public Const REFERENCE_INI As String = "./Reference.ini" '�O���@�Ǹ��^�����������(��������)
Public Const HISSYNC_INI As String = "./HISSync.ini" '�O���@�Ǹ�HISSync�n�U�����Ǹ�Ʀ��������(��������)
Public Const SWCONFIG_TAIAN_INI As String = "./SWConfig_TAIAN.ini" '�O���@�Ǹ�HISSync�n�U�����Ǹ�Ʀ��������(��������)
Public Const CATH_XML_INI As String = "./Cath_XML.ini" '�O���@��XML�ഫ�t�η|�Ψ쪺XML�w�]�s���m
Public Const PACS_UPLOAD_INI As String = "./Pacs_Upload.ini" '�O���@��Dicom_Upload�ӵ{���A�w�]�s��Ʈw����m��
Public Const OFFLINE_REPORT_INI As String = "./OfflineReport.ini" 'OfflineReport����{���|�Ψ쪺�]�w���
Public Const DMMHE_INI As String = "./DMMHE.ini" '�O���@��Delete_MMH_Exam�ӵ{���A�n�R�h�֤���H�W����Ƨ��θ�Ʈw��Ƶ�
Public Const ERASE_LOG_INI As String = "./Erase_Log.ini" '�O���@��Erase_Log�ӵ{���A�n���}�h�֤���H�W���S�w���ɦW��Ƶ�
Public Const DRAFT_INI As String = "./Draft.ini" '�O���@��Draft�һݪ��s�ɦ�m�����
Public Const COMMANDFROM_INI As String = "./CommandFrom.ini" '�O���@��CommandFrom�һݪ��s�ɦ�m�����
Public Const CHECKPRINTERSTATUS_INI As String = "./CheckPrinterStatus.ini" '�O���@��CheckPrinterStatus�һݪ��s�ɦ�m�����
'/*�p�حק諸(2009/02/04)*/

'/**************************�p�حק諸(2009/04/14)***********************************/


'/**************************���Ϊ��n�J�ϥΪ��ܼƸ��***********************************/

'/*�O���n�J�̪���T*/
Public Login_ID As String '�ϥΪ̱b���A������Ʈwcris_user��logonuser
Public Login_PW As String '�ϥΪ̱K�X�A������Ʈwcris_user��password
Public Login_Name As String '�ϥΪ̦W�١A������Ʈwcris_user��name
Public Login_Position As String '�ϥΪ�¾�١A������Ʈwcris_user��type
Public Login_No As String '�ϥΪ̽s���A������Ʈwcris_user��userid
Public Login_Power As Integer '�ϥΪ��v���A������Ʈwcris_user��authorid
Public Login_Phone As String '�ϥΪ̹q�ܡA������Ʈwcris_user��phone
Public Login_HostName As String '�ϥΪ̥D���W�١A����ini�̪�hostname
Public Login_System As String '�ϥΪ̬�O�A������Ʈwcris_user��system
'/*�p�حק諸(2009/03/18)*/


'/*�O���}�Ҫ�������T*/
Public Login_LastOpen As String '�w�L�k�o��
Public Login_LastOpenReportList As String ''�ϥΪ̳̫�}�������e�N��(�������D�e����)
Public Login_LastOpenReportBody As String '�ϥΪ̳̫�}�������e�A������Ʈwcris_exam_online��item6
Public Login_LastOpenReportType As String '�ϥΪ̳̫�}�������ˬd���O�A������Ʈwcris_exam_online��type
Public Login_LastOpenReportStatus As String '�ϥΪ̳̫�}���������i���A�A������Ʈwcris_exam_online��status
Public Login_LastOpenReportUnikey As String '�ϥΪ̳̫�}���������iunikey�A������Ʈwcris_exam_online��uni_key
Public Login_LastOpenReportChartNO As String '�ϥΪ̳̫�}���������i���f�����A������Ʈwcris_exam_online��chartno
Public Login_LastOpenReportChartName As String '�ϥΪ̳̫�}���������i�f�w�m�W
Public Login_LastOpenReportDrFrom As String '�ϥΪ̳̫�}���������i���E�_�ӷ��A������Ʈwcris_exam_online��dr_from
Public Login_LastOpenReportModifyDate As String '�ϥΪ̳̫�}�������T�{����A������Ʈwcris_exam_online��examdate
Public Login_LastOpenReportModifyTime As String '�ϥΪ̳̫�}�������T�{�ɶ��A������Ʈwcris_exam_online��examtime
Public Login_LastOpenReportReturnDate As String '�ϥΪ̳̫�}���������i����A������Ʈwcris_exam_online��reportdate
Public Login_LastOpenReportReturnTime As String '�ϥΪ̳̫�}���������i�ɶ��A������Ʈwcris_exam_online��reporttime
Public Login_LastOpenReportCreateDRName As String '�ϥΪ̳̫�}�������}����ͦW�r�A������Ʈwcris_exam_online��dr_on
Public Login_LastOpenReportCreateDRPhone As String '�ϥΪ̳̫�}�������}����͹q��
Public Login_LastOpenReportFieldList As String '�ϥΪ̳̫�}��������ܪ����
Public Login_LastOpenReportSystem As String '�ϥΪ̳̫�}��������O�A������Ʈwcris_exam_online��system
'/*�p�حק諸(2009/03/18)*/



'/*�Ω�O��frmQueue����ơA�H�Ǩ�frmReport�ӭ��ϥ�*/
Public Report_DrReport As String '�Ω�O��frmQueue�����i��v
Public Report_DrOn As String '�Ω�O��frmQueue���ˬd��v
'/*�p�حק諸(2009/03/18)*/


'/*�Ω�O��frmSpread����ơA�H�Ǩ�frmSpread�ӭ��ϥ�*/
Public Spread_ID As String
Public Spread_Name As String
'/*�p�حק諸(2009/09/14)*/


'/*�O���ϥΪ̦b�ϥ�frmImageSelect���]�w*/
Public Img_Sel_Cols As Long '�x�s�bfrmImageSelect���ϥΪ̪���ƹw�]�Ӧ��X�i���]�w
Public Img_Sel_Rows As Long '�x�s�bfrmImageSelect���ϥΪ̪����ƹw�]�Ӧ��X�i���]�w
'/*�p�حק諸(2009/03/18)*/


'/*�O���ϥΪ̦b�ϥ�frmRepPreview���]�w*/
Public Img_Sel_Opt As Long '�x�s�bfrmRepPreview���C�L�榡�w�]�n�]�X���X
'/*�p�حק諸(2009/03/18)*/


'/*�O���ϥΪ̦b�ϥ�frmReport���]�w*/
Public defaultImgZoomLeft As Long '�O���bfrmReport���Ϥ��j�p���ܫe���w�]X�b��m
Public defaultImgZoomTop As Long '�O���bfrmReport���Ϥ��j�p���ܫe���w�]Y�b��m
Public defaultImgZoomWidth As Long '�O���bfrmReport���Ϥ��j�p���ܫe���w�]�e��
Public defaultImgZoomHeight As Long '�O���bfrmReport���Ϥ��j�p���ܫe���w�]����
'/*�p�حק諸(2009/04/15)*/


'/*�O���ϥΪ̦b�ϥ�frmRepPreview���]�w*/
Public ImgSvr_Path As String '���ݭn�n�J�i�h��Image Server�����|
'/*�p�حק諸(2009/04/16)*/


'/*�ΨӨM�w�D���O�_��ݳo�����i*/
Public MasterCanNotSee As Boolean
'/*�p�حק諸(2009/06/26)*/


'/*�Ω��frmViewDcmList���q���ܼ�*/
Public imgList As Integer
Public imgFilePath() As String
Public imgFileName() As String
'/**/

'/**************************�p�حק諸(2009/02/03)***********************************/



'/*************************���Ϊ��n�J�ϥΪ̵��c���**********************************/

'/*�Ω�W��HIS���ӰƵ{���A��e���W�Ҧ�������S�����ݩʳ�COPY�U�ӥΪ�TYPE*/
Private Type SortObject
    ID As Long
    Name As String
    Text As String
    Left As Double
    Top As Double
    TrueLeft As Long
    TrueTop As Long
End Type
'/**/

'/**************************�p�حק諸(2009/02/03)***********************************/



'/*                  ����w�����Ψ�W������A���ഫ����r�ɡA�äW�ǳ��i��HIS                                     */
Public Sub SaveToHIS(ByRef frm As Form)
    Dim i As Long
    Dim j As Long
    
    Dim diff As Double
    
    Dim obj As Object
    
    Dim SortObj() As SortObject
    Dim Pic1_Left() As Long
    Dim Pic1_Top() As Long
    Dim Pic3_Left() As Long
    Dim Pic3_Top() As Long
    
    Dim All_Count As Long
    Dim Pic1_Count As Long
    Dim Pic3_Count As Long
    
    '/*��X�D�������B�R�O���s�ιϤ�������@���X��*/
    All_Count = 0
    Pic1_Count = 0
    Pic3_Count = 0
    For Each obj In frm.Controls
        If VarType(obj) <> VARTYPE_COMMONDIALOG And VarType(obj) <> VARTYPE_COMMAND And VarType(obj) <> VARTYPE_PICTUREBOX Then
            All_Count = All_Count + 1
        ElseIf VarType(obj) = VARTYPE_PICTUREBOX Then
            If Right(obj.Name, 1) = "1" Then
                If obj.Index > Pic1_Count Then
                    Pic1_Count = obj.Index
                End If
            ElseIf Right(obj.Name, 1) = "3" Then
                If obj.Index > Pic3_Count Then
                    Pic3_Count = obj.Index
                End If
            End If
        End If
    Next
    '/**/
    
    '/*�w�q�@�U�n�쪺���󪺤j�p*/
    ReDim SortObj(All_Count)
    ReDim Pic1_Left(Pic1_Count)
    ReDim Pic1_Top(Pic1_Count)
    ReDim Pic3_Left(Pic3_Count)
    ReDim Pic3_Top(Pic3_Count)
    '/**/
    
    '/*��Ҧ���picturebox����Ƨ�X��*/
    For Each obj In frm.Controls
        If VarType(obj) = VARTYPE_PICTUREBOX Then
            If Right(obj.Name, 1) = "1" Then
                Pic1_Left(obj.Index) = obj.Left
                Pic1_Top(obj.Index) = obj.Top
            ElseIf Right(obj.Name, 1) = "3" Then
                Pic3_Left(obj.Index) = obj.Left
                Pic3_Top(obj.Index) = obj.Top
            End If
        End If
    Next
    '/**/
    
    '/*��Ҧ��D�������B�R�O���s�ιϤ�����Ƨ�X��*/
    Dim temp() As String
    i = 0
    For Each obj In frm.Controls
        If VarType(obj) <> VARTYPE_COMMONDIALOG And VarType(obj) <> VARTYPE_COMMAND And VarType(obj) <> VARTYPE_PICTUREBOX Then
            SortObj(i).ID = i
            SortObj(i).Name = obj.Name
            Select Case UCase(Left(obj.Name, 3))
            Case "TXT", "TEX"
                SortObj(i).Text = obj.Text
            Case "LBL", "LAB"
                SortObj(i).Text = obj.Caption
            Case "OPT", "CHE"
                If obj.Value Then
                    SortObj(i).Text = "��" & obj.Caption
                Else
                    SortObj(i).Text = "��" & obj.Caption
                End If
            End Select
            


           SortObj(i).Left = obj.Left
           SortObj(i).Top = obj.Top
           If UCase(Left(obj.Name, 3)) = "TXT" Or UCase(Left(obj.Name, 3)) = "LAB" Then
                If obj.LinkItem <> "" Then
                    temp = Split(obj.LinkItem, "_")
                    
                    If temp(0) = "1" Then
                        SortObj(i).TrueLeft = obj.Left + Pic1_Left(temp(1))
                        SortObj(i).TrueTop = obj.Top + Pic1_Top(temp(1))
                    ElseIf temp(0) = "3" Then
                        SortObj(i).TrueLeft = obj.Left + Pic3_Left(temp(1))
                        SortObj(i).TrueTop = obj.Top + Pic3_Top(temp(1))
                    End If
                Else
                    SortObj(i).TrueLeft = obj.Left
                    SortObj(i).TrueTop = obj.Top
                End If
            Else
                SortObj(i).TrueLeft = obj.Left
                SortObj(i).TrueTop = obj.Top
            End If
            
            
            i = i + 1
        End If
    Next
    '/**/
    
    '/*�w��o�Ǫ��󪺸�ơA���쥻�b���W����m���ӱƧ�*/
    For i = 0 To All_Count - 1
        For j = i + 1 To All_Count - 1
            If SortObj(i).TrueTop > SortObj(j).TrueTop Or (SortObj(i).TrueLeft > SortObj(j).TrueLeft And SortObj(i).TrueTop >= SortObj(j).TrueTop) Then
                Call swap(SortObj(i).Name, SortObj(j).Name)
                Call swap(SortObj(i).Left, SortObj(j).Left)
                Call swap(SortObj(i).Top, SortObj(j).Top)
                Call swap(SortObj(i).Text, SortObj(j).Text)
                Call swap(SortObj(i).ID, SortObj(j).ID)
                Call swap(SortObj(i).TrueLeft, SortObj(j).TrueLeft)
                Call swap(SortObj(i).TrueTop, SortObj(j).TrueTop)
            End If
        Next
    Next
    '/**/
    
    '/*�̤W�U���k�����ǡA�N��g���@���r�ꪫ�󤤡A���ݭn���ܤ]�i�H�g�i�����(�C�@�Ӫ���b���W���רC�Z��500���A�b��r�ɤ��Y�_�@����)*/
    Login_LastOpenReportBody = ""
    'Open "test.txt" For Output As #1
        For i = 0 To All_Count - 2
            Call Replace(SortObj(i).Text, " ", "")
            
            'Print #1, SortObj(i).Text, ;
            If SortObj(i).Text <> "" Then
                Login_LastOpenReportBody = Login_LastOpenReportBody & SortObj(i).Text
            End If
            
            If SortObj(i + 1).TrueTop > SortObj(i).TrueTop Then
                diff = SortObj(i + 1).TrueTop - SortObj(i).TrueTop
                
                'If SortObj(i).name <> "lblImgTime" Then
                    Do Until diff <= 0
                        'Print #1, ""
                        Login_LastOpenReportBody = Login_LastOpenReportBody & vbCrLf
                        
                        diff = diff - 500
                    Loop
                'End If
            End If
            
        Next
        
        'Print #1, SortObj(i).Text
        Login_LastOpenReportBody = Login_LastOpenReportBody & SortObj(i).Text & vbCrLf
    'Close #1
    'Debug.Print Len(Login_LastOpenReportBody)
    '/**/
    
    
    '/*�g�J�ڭ̪���Ʈw�A�H��MPPS�W�Ǩ�HIS����Ʈw��*/
    If Connection.State Then
        i = InStr(1, Login_LastOpenReportBody, "SPIROMETRIC INTERPRETATION")
        Login_LastOpenReportBody = Right(Login_LastOpenReportBody, Len(Login_LastOpenReportBody) - i + 1)
        Call DBRecordLog("update", "update cris_exam_online set item1='" & Login_LastOpenReportBody & "' where status<>'�w�R��' and uni_key='" & Login_LastOpenReportUnikey & "' ", "��MPPS�W�Ǩ�HIS����Ʈwcris_exam_online")
        Connection.Execute "update cris_exam_online set item1='" & Login_LastOpenReportBody & "' where status<>'�w�R��' and uni_key='" & Login_LastOpenReportUnikey & "' "
    End If
    '/**/
End Sub
'/**/
'/**************************�p�حק諸(2009/07/03)***********************************/




'/*�ΥHŪ����r�ɤ���template�λ����ɡA�ö����W���Ƶ{��*/
Public Function NI_Report_Load_From_File(ByVal ReportFilePath As String, ByVal TemplateFilePath As String, ByVal MAX_FIELD As Integer) As String()
    On Error GoTo errout:

    '/*���O�O�A��z�Ǧ^�ɥΪ�report�t�C�ܼơAŪ�ɮ׮ɪ��ȥ��ܼơA�����ܼ�*/
    Dim ReportValue() As String
    Dim ReportValue_PSV() As String
    Dim ReportValue_EDV() As String
    ReDim ReportValue(MAX_FIELD) As String
    ReDim ReportValue_PSV(MAX_FIELD) As String
    ReDim ReportValue_EDV(MAX_FIELD) As String
    
    
    Dim InputValue As String
    Dim FieldValue() As String
    Dim FieldWord() As String
    
    Dim i As Integer
    '/**/
    
    '/*script����A�y���ĥ�java script*/
    Dim MSSC As New MSScriptControl.ScriptControl
    MSSC.Language = "JavaScript"
    '/**/
    
    
    '/*�ĥ�java script��eval����ơA��C�ӦW�����ӬO�b�}�C�X�s��java script���ܼƤ�*/
    FreeFilePort = FreeFile
    Open TemplateFilePath For Input As #FreeFilePort
        Do Until EOF(1)
            Line Input #1, InputValue
            
            FieldValue = Split(InputValue, "=")
            Call MSSC.Eval("var " & FieldValue(0) & "=" & FieldValue(1))
        Loop
    Close #FreeFilePort
    '/**/


    '/*��q��r�ɤ����ƭȡA��java script��J���ө��ݪ���줤�C��J�e�n����psv��edv�Ӥ����k*/
    FreeFilePort = FreeFile
    Open ReportFilePath For Input As #FreeFilePort
        Do Until EOF(1)
            Line Input #1, InputValue
            
            FieldValue = Split(InputValue, "=")
            FieldWord = Split(FieldValue(0), "_")
                        
            If LCase(FieldWord(2)) = "psv" Then
                ReportValue_PSV(MSSC.Eval(FieldWord(0) & "_" & FieldWord(1))) = FieldValue(1)
            ElseIf LCase(FieldWord(2)) = "edv" Then
                ReportValue_EDV(MSSC.Eval(FieldWord(0) & "_" & FieldWord(1))) = FieldValue(1)
            End If
        Loop
    Close #FreeFilePort
    '/**/
    
    
    '/*�N���Ӥ��j�����Υk���ɮ���s���r�ɤ��A�öǦ^�H�K�Q��*/
    For i = 0 To MAX_FIELD - 1
        ReportValue(i) = ReportValue_PSV(i) & "/" & ReportValue_EDV(i)
    Next
    NI_Report_Load_From_File = ReportValue
    '/**/
    
    
    If False Then
errout:
        Select Case err.Number
        Case 9, 5009
            Resume Next
        Case Else
            Call PrintLog("Error In NI_Report_Load_From_File")
            ReDim ReportValue(0)
            ReportValue(0) = "Error"
            NI_Report_Load_From_File = ReportValue
        End Select
    End If
End Function
'/*20100120*/




'/*�ΥHŪ����r�ɤ���template�λ����ɡA�ö��Spread���W���Ƶ{��*/
Public Function NI_Spread_Report_Load_From_File(ByVal ReportFilePath As String, ByVal TemplateFilePath As String, ByRef fps As fpSpread, ByVal SheetID As Integer) As Boolean
    On Error GoTo errout:

    '/*���O�O�A�Ψө�mŪ�X�Ӫ�spread��m���ܼơAŪ�ɮ׮ɪ��ȥ��ܼ�*/
    Dim ReportValuePlace() As String
    Dim TemplateValuePlace As String
    
    Dim InputValue As String
    Dim FieldWord() As String
    '/**/
    
    '/*���w�o���n��s���O���@�isheet*/
    fps.sheet = SheetID
    '/**/
    
    '/*��q��r�ɤ�Ū�X���ƭȡA��J��spread�����󤤡C*/
    FreeFilePort = FreeFile
    Open ReportFilePath For Input As #FreeFilePort
        Do Until EOF(1)
            Line Input #1, InputValue
            
            FieldWord = Split(InputValue, "=")
            If UBound(FieldWord) > 0 Then
            
                TemplateValuePlace = InputINI("Sheet_" & SheetID, FieldWord(0), TemplateFilePath)
                If TemplateValuePlace <> "" Then
                
                    ReportValuePlace = Split(TemplateValuePlace, ",")
                    If UBound(ReportValuePlace) > 0 Then
                        Call fps.SetText(CharEN2Numeric(ReportValuePlace(0)), ReportValuePlace(1), FieldWord(1))
                    End If
                End If
            End If
        Loop
    Close #FreeFilePort
    '/**/


    NI_Spread_Report_Load_From_File = True
    
    
    If False Then
errout:
        Select Case err.Number
        Case 9
            Resume Next
        Case Else
            Call PrintLog("Error In NI_Spread_Report_Load_From_File")
            NI_Spread_Report_Load_From_File = False
        End Select
    End If
End Function
'/*20100122*/


