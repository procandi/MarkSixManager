Attribute VB_Name = "Module1"
'�t�ΥD��Ʈw��m
Global dbConnection$ '= "Server=127.0.0.1;DRIVER={SQL Server};UID=sa;PWD=sameway;DATABASE=CRIS;"

'�y�t
Global NLS_LANG$

'������ե�
'Global Const dbConnection$ = "SERVER=CRS;DRIVER={Microsoft ODBC for Oracle};UID=mpacs;PWD=mpacs;"
Global sDB2Conn As String

'�P�s�� CASViewer�����Ѽ� -------------------------------------------------

'��Ͻs��Ҧ�--------------------------------------------------------------
Global Draft_Simp$
Global Draft_Comp$
'--------------------------------------------------------------------------

Global xLastUpdateDate$, xLastUpdateTime$
Global sLastUpdateDate$, sLastUpdateTime$
'--------------------------------------------------------------------------

'Public Const RunType$ = "TEST"
'Public Const RunType$ = "PRODUCTION"

'�ϥΪ�
Global UserID$
Global UserType$
Global UserDivision$
Global UserName$
Global Password$


'�ֳt��w�q
Global Const HotKeyTrigger = 39
Global current_Dr_info$

Global originalStud_No As String
Global currForm As Object
Global currControl As Object
Global currSelectTopic As String
Global curr_YYMM$
Global curr_Dept$
Global Font_Size%

'Capture �����t���ܼ�
Global portStepBoard%               '��O�ϥ� Port #  CapSVR.ini
Global portRefresh%                 '��O�F�ӫ�
Global ImgSVR_HostName$             '�u�@���W�� CapSVR.ini / ExamSVR.ini
Global Instrument$                  '�ˬd�����s��
Global RoomName$                    '�ˬd��ǦW��
Global Dr_from$                     '�ӷ��O
Global global_Site$
Global global_SiteEng$
Global IMP_SCP$
Global Report_Name$                 '�ˬd�����έp����W��
Global Report_Name1$                '�ƯZ����W��
Global Report_Name2$                'structed report�έp����W��
Global CheckValue$                  '<>NO�ɡA�N�|�P�_�߶W�ƭȶW�L�зǭȮ��ܬ����r
Global Pre_PickImage$               'Pre_PickImage=YES�ɡA�~�|�۰ʦb�ӵ����������ˬd���A�U�A�B�������v���ɡA�۰ʥ���Ҧ��v��
Global IsPopSmart$                  '�����R341005����conclusion��쬰�ŮɡA�O�_�۰ʼu�X���z���PŪ�e���AIsPopSmart$<>"NO"�ɼu�X
Global Q_SortOrder$                 '����d�ߵe�����ƧǥѤ���ܧ��ˬd�渹�AQ_SortOrder$="UNI_KEY"���ܧ󬰨��ˬd�渹�Ƨ�
Global Need_Dr_On$                  '�O�_�b�L�ޮv��ƮɡA�۰ʥH���i��v��J�ANeed_Dr_On=YES�ɦ۰ʶ�J

'���W/ARFI����һݪ��}�C
Global AContent(20) As String       '�Ω�s��U���ت���
Global Aindex(20, 3) As String      '�Ω�s��U���ت��W��/���W��/�w�]��/ARFI���س]�w

'���������һݪ��}�C
Public Const Bronchoscopy_Number = 300
Global BContent(Bronchoscopy_Number, 1) As String   '�Ω�s��USER���ﶵ����ȻP�ܧ��
Global Bindex(Bronchoscopy_Number, 5) As Integer    '�Ω�s��U���ت����Y�A�̧Ǭ�(�ʽ�/���ؼh��/�W�@�h���p/�U�@�h���p/�e�����p/�ᶵ���p)
Global xList(3, 50, 1) As String      '�x�sListbox�r��
Global yList(3) As Integer          '�x�s�ﶵ�ʽ�
Global zList(3, 50) As Integer      '�x�s�e���W�������w��`�I�ȡA����̬�-1

Global ExamType$
Global trimLeft%, trimTop%, trimWidth%, trimHeight%
Global trimDefault%
Global Offline_Path$
Global Current_ImagePath$
Global Destination_ImagePath$
Global Grayscale$

Global Msg As String, Title As String
Global style As Integer, Response As Integer

Global LastPageNo As Integer
Global LastDraftNo As Integer

Global curr_ChartNo$, curr_ExamDate$, curr_ExamType$

Global global_currOption As String

'Global currForm As Variant

Type typeExam_online

     System As String
     
     uni_key As String
     chartno As String
     Date As String
     time As String
     
     Type As String
     Room As String
     Age As String
     Item1 As String
     Item2 As String
     Item3 As String
     Item4 As String
     Item5 As String
     Item6 As String
     Item7 As String
     Item8 As String
     Item9 As String
     Item10 As String
     Others As String
     UploadCode As String
     
     Dr_from As String
     Dr_on As String
     Dr_order As String
     Dr_report As String
     
     '�x�x�z�G����s�W�H�f��v
     Dr_follow As String
     Division_on As String
     Division_from As String
     
     Status As String
     Class As String
     ImgPicked As String
     
     Name As String
     Sex As String
     BirthDay As String
     Phone As String
     Address As String
     CitizenID As String
     
     Modality As String
     
     Reg_Date As String
     ExamDetail As String
     
     OrderDate As String
     OrderTime As String
     ReportDate As String
     ReportTime As String
     LastUpdateDate As String
     LastUpdateTime As String
     
     Division_Seq As String
     ClinicalImp As String
     
     TemplateName As String
     TemplateFile As String
     
     ChargeBy As String
     
     HIS_ReqNo As String
     
End Type

Global curr_Record As typeExam_online
Global save_Record As typeExam_online

Type typeImage
     ImgOwner As String
     Class As String
     Type As String
     GroupName As String
     GroupIndex As String
     CreateDate As String
     FilePath As String
     FileName As String
     CreateTime As String
     CreateBy As String
     Backup As String
     Memo As String
     MutliFrame As Boolean
End Type
Global Img_Array() As typeImage
Global Draft_Array() As typeImage

Type BasicInfo
     chartno As String
     Name As String
     CitizenID As String
     Sex As String
     BirthDay As String
     Nationality As String
     Address As String
     Phone As String
     Date_1st As String
     Recorder As String
     Status As String
End Type
Global thisOne As BasicInfo

Type schema_Diction
     Exam_Type As String
     Class As String
     Code As String
     Content As String
End Type
'Global xDiction1() As schema_Diction
Global xDiction() As Variant
Global xDictionByType() As Variant
Global xDictionMaxRows%

Global xPhrase As Variant

Type dictionary
     System As String
     Type As String
     Class As String
     Code As String
     Division As String
     UserID As String
     Content As String
End Type
Global xDictionarySpread() As dictionary

'���i�d��
Type typeReportTemplate
     DivisionID As String
     DivisionName As String
     ExamID As String
     ExamName As String
     TemplateFileName As String
     TemplateFileSource As String
     UserID As String
     DefaultUse As String
     ExamDescription As String
End Type
Global xReportTemplate() As typeReportTemplate

Global path_System$, path_Images$, path_Define$
Global path_Target$
Global db_Name$

Global ImgShellApp$ '�t�s�s�ɮɩҩI�s�����ε{��

Global dbDictionFile As String
Global initPath As String

Global StrTrans$
Global currObj As Object
Global CodeSet(36)  As String

'Print �����ܼ�
Global ImgSVRPrinter As String
Global prnForm_Stream_NoImage As String
Global prnForm_Stream1 As String
Global prnForm_Stream2 As String
Global prnForm_Stream3 As String
Global prnForm_Stream4 As String
Global prnForm_Stream5 As String
Global prnForm_Stream6 As String
Global prnForm_NoImage As String

Global prnFormString$
Global prnFormImageCols%

Global history_Form As String
Global xPrintForm As String

Global ReportName As String
Global ReportPath As String
Global BAKPath As String
Global LableStatus As String
Global xSpread2$
Global xTimeSet$
Global xDivision_On$
Global xDisplay_UnikeyName$
Global xSee_OCR$
Global xall_up_button$
Global Need_Dr_Confirm$
Global Open_PDF$
Global Update_Backup$
Global Enable_Report$
Global GetReportEdit$
Global IS_Hync$
Global No_Report_Image$
Global Is_GetPatientFromWeb$
Global URLAddress$

'�Ω󷧥ܹϽs�誺�Ѽ�
Global PLocationNote As String
Global PDiagnosisNote As String
Global Psplit As String

Sub INI_Read()
    Dim rtn As Long
    Dim tmpA As String * 260
    Dim SVR_Queue_INI$, SVR_Capture_INI$
    
    'SVR_Queue �D�n�Ѽ�----------------------------------------------------------------------------
    SVR_Queue_INI$ = App.Path & "\EXAMSVR.INI"
    
    '20131002�A�p�حn�D�A�s�W���ˮɡA�Y�L���f�w��ƮɡA�O�_�qWEB�����f�w��ơA�w�]��NO�AYSE�ɤ~����
    Is_GetPatientFromWeb$ = InputINI("ImgSVR Host", "Is_GetPatientFromWeb", SVR_Queue_INI$)
    If Is_GetPatientFromWeb$ = "" Then
        Is_GetPatientFromWeb$ = "NO"
    End If
    
    '20131003�A����Is_GetPatientFromWeb���t�M�]�w�A�Ω�]�w���������}x�A�{�����|�b�����}��[�W�y�f�����z�h����
    URLAddress$ = InputINI("ImgSVR Host", "URLAddress", SVR_Queue_INI$)
    If Right(URLAddress$, 1) <> "/" And Right(URLAddress$, 1) <> "\" Then
        URLAddress$ = URLAddress$ & "/"
    End If
    
    '20130809�A�p�حn�D�A�z�G��]�����q�����C�A�ҥH�W�[�@�ӳ]�w�A�������i��v���A�w�]No_Report_Image��NO�A��YES�ɤ������i��v��
    No_Report_Image$ = InputINI("ImgSVR Host", "No_Report_Image", SVR_Queue_INI$)
    If No_Report_Image$ = "" Then
        No_Report_Image$ = "NO"
    End If
    
    '20130809�A�p�حn�D�A�]�wINI�A���Juni_key�ɤ�����A�w�]��YES(����)
    IS_Hync$ = InputINI("ImgSVR Host", "IS_Hync", SVR_Queue_INI$)
    If IS_Hync$ = "" Then
        IS_Hync$ = "YES"
    End If
    
    '���s��x�x�z�G���쪺���W����w�]�O���g�s����i�����i�J�w���C�L�A���Y�n�g�s����i�ɡA�h�ݳ]�w��YES
    GetReportEdit$ = InputINI("ImgSVR Host", "GetReportEdit", SVR_Queue_INI$)
    If GetReportEdit$ = "" Then
        GetReportEdit$ = "NO"
    End If
    
    '���s��s�W�w�A�o��v�~�i�H�T�{���i�A�]�ȥH��i��|�����СA�ҥH�אּINI�]�w�A�w�]�Ȭ�Need_Dr_Confirm=YES�A����YES�ɴN�����w
    Need_Dr_Confirm$ = InputINI("ImgSVR Host", "Need_Dr_Confirm", SVR_Queue_INI$)
    If Need_Dr_Confirm$ = "" Then
        Need_Dr_Confirm$ = "YES"
    End If
    
    '�O�_�i�H�����i�A�w�]���i�AEnable_Report=NO�ɤ��i�ܧ���i���e
    Enable_Report$ = InputINI("ImgSVR Host", "Enable_Report", SVR_Queue_INI$)
    If Enable_Report$ = "" Then
        Enable_Report$ = "YES"
    End If
    
    '���s��x�x�z�G���쪺�v���������n�W�ǡA�s�W����sbackup���]�w�H����USER���p���I������A�]�^���v���ɹw�]backup=0�A�w�]��Update_Backup=YES�A��Update_Backup=NO�ɤ���s
    Update_Backup$ = InputINI("ImgSVR Host", "Update_Backup", SVR_Queue_INI$)
    If Update_Backup$ = "" Then
        Update_Backup$ = "YES"
    End If
    
    '�O�_�n�P�t�ήɶ���ɡA�w�]��YES�A��TimeSet=NO�ɬ�����
    xTimeSet$ = InputINI("ImgSVR Host", "TimeSet", SVR_Queue_INI$)
    If xTimeSet$ <> "NO" Then
        xTimeSet$ = "YES"
    End If
    
    '���s��ߤ��n�D�s�W�}��PDF�ؿ������s�\��A�w�]���}�ҡA�u��Open_PDF=NO�~�|����
    Open_PDF$ = InputINI("ImgSVR Host", "Open_PDF", SVR_Queue_INI$)
    If Open_PDF$ = "" Then
        Open_PDF$ = "YES"
    End If
    
    '�x�x�z�G����n�D�A�w���C�L���W�[�@�ӦC�L�ݽT�{���i�����s�Aall_up_button=YES�ɬ���ܸӫ��s�A�w�]��all_up_button<>YES�ɤ����
    xall_up_button$ = InputINI("ImgSVR Host", "all_up_button", SVR_Queue_INI$)
    If xall_up_button$ = "" Then
        xall_up_button$ = "NO"
    End If
    
    '�O�_�}��OCR�v���ɥi�b���i�t�Τ��ݨ�A�ﶵ���e���ˬd���O����r�r��A�����H�y�F�z�j�}�A�ҦC���ˬd���O�N�O�n���OCR�v�����ˬd���O
    xSee_OCR$ = InputINI("ImgSVR Host", "See_OCR", SVR_Queue_INI$)
    
    '�d�ߵe���W�A�ӽг渹����r�ܧ�A�w�]�Ȭ�Display_UnikeyName=�ӽг渹�A�x�x�z�G����n�D��ܡy�Ǹ��z
    xDisplay_UnikeyName$ = InputINI("ImgSVR Host", "Display_UnikeyName", SVR_Queue_INI$)
    If Len(xDisplay_UnikeyName$) < 1 Then
        xDisplay_UnikeyName$ = "�ӽг渹"
    End If
    
    '���i�t�Τ�����ɪ���ǧO�P�w�A�p��Ŧ����/���ѩʤ�Ŧ�f��/�x�x�z�G���쵥�A�w�]�Ȭ���Ŧ����
    xDivision_On$ = InputINI("ImgSVR Host", "Division_On", SVR_Queue_INI$)
    If Len(xDivision_On$) < 1 Then
        xDivision_On$ = "��Ŧ����"
    End If
    
    xSpread2$ = InputINI("ImgSVR Host", "Spread2", SVR_Queue_INI$)
    If Len(xSpread2$) < 1 Then
        xSpread2$ = "�x�x�z�G����"
    End If
    
    '���s��x�x�z�G����A�d�ߵe���Ʊ�ݨ�Ĥ@�Ӽ��ҦW�٬����ˬd�A���ѻP�ߤ��h���w�ˬd�A�w�]�Ȭ��w�ˬd
    LableStatus = InputINI("ImgSVR Host", "Status", SVR_Queue_INI$)
    
'    rtn = ReadINI("Report Exam", "ReportName", "", tmpA, Len(tmpA), SVR_Queue_INI$)
'    If Not rtn = 0 Then ReportName = Replace(Trim(Left(tmpA, rtn)), Chr(0), "")
'    '�]�w�έp����{���W�١A�pReport_Name=ABC.exe
'    ReportName = InputINI("Report Exam", "ReportName", SVR_Queue_INI$)
        
'    rtn = ReadINI("Report Exam", "ReportPath", "", tmpA, Len(tmpA), SVR_Queue_INI$)
'    If Not rtn = 0 Then ReportPath = Replace(Trim(Left(tmpA, rtn)), Chr(0), "")
    ReportPath = InputINI("Report Exam", "ReportPath", SVR_Queue_INI$)
    
'    rtn = ReadINI("Report Exam", "ReportPath", "", tmpA, Len(tmpA), SVR_Queue_INI$)
'    If Not rtn = 0 Then BAKPath = Replace(Trim(Left(tmpA, rtn)), Chr(0), "")
    BAKPath = InputINI("Report Exam", "ReportPath", SVR_Queue_INI$)
    
'    rtn = ReadINI("Image Shell", "Application", "", tmpA, Len(tmpA), SVR_Queue_INI$)
'    If Not rtn = 0 Then ImgShellApp$ = Left(tmpA, rtn)
    ImgShellApp$ = InputINI("Image Shell", "Application", SVR_Queue_INI$)
      
'    rtn = ReadINI("ImgSVR Host", "HostName", "", tmpA, Len(tmpA), SVR_Queue_INI$)
'    If Not rtn = 0 Then ImgSVR_HostName$ = Left(tmpA, rtn)
    ImgSVR_HostName$ = InputINI("ImgSVR Host", "HostName", SVR_Queue_INI$)
    
'    rtn = ReadINI("ImgSVR Host", "Pre_PickImage", "", tmpA, Len(tmpA), SVR_Queue_INI$)
'    If Not rtn = 0 Then Pre_PickImage$ = Left(tmpA, rtn)
    Pre_PickImage$ = InputINI("ImgSVR Host", "Pre_PickImage", SVR_Queue_INI$)
    
'    rtn = ReadINI("ImgSVR Host", "CheckValue", "", tmpA, Len(tmpA), SVR_Queue_INI$)
'    If Not rtn = 0 Then CheckValue$ = Left(tmpA, rtn)
    CheckValue$ = InputINI("ImgSVR Host", "CheckValue", SVR_Queue_INI$)
    
'    rtn = ReadINI("ImgSVR Host", "RoomName", "", tmpA, Len(tmpA), SVR_Queue_INI$)
'    If Not rtn = 0 Then RoomName$ = Left(tmpA, rtn)
    RoomName$ = InputINI("ImgSVR Host", "RoomName", SVR_Queue_INI$)
    
'    rtn = ReadINI("ImgSVR Host", "Site", "", tmpA, Len(tmpA), SVR_Queue_INI$)
'    If Not rtn = 0 Then tmp$ = Trim(Left(tmpA, rtn))
    global_Site$ = InputINI("ImgSVR Host", "Site", SVR_Queue_INI$)
'    global_Site$ = ini_Purge(tmp$, rtn)
    
    '�έp����{���W��
'    rtn = ReadINI("ImgSVR Host", "Report_Name", "", tmpA, Len(tmpA), SVR_Queue_INI$)
'    If Not rtn = 0 Then tmp$ = Trim(Left(tmpA, rtn))
'    Report_Name$ = ini_Purge(tmp$, rtn)
    Report_Name$ = InputINI("ImgSVR Host", "Report_Name", SVR_Queue_INI$)
    
    '�ƯZ����{���W��
'    rtn = ReadINI("ImgSVR Host", "Report_Name1", "", tmpA, Len(tmpA), SVR_Queue_INI$)
'    If Not rtn = 0 Then tmp$ = Trim(Left(tmpA, rtn))
'    Report_Name1$ = ini_Purge(tmp$, rtn)
    Report_Name1$ = InputINI("ImgSVR Host", "Report_Name1", SVR_Queue_INI$)
    
    'SR�έp����{���W��
'    rtn = ReadINI("ImgSVR Host", "Report_Name2", "", tmpA, Len(tmpA), SVR_Queue_INI$)
'    If Not rtn = 0 Then tmp$ = Trim(Left(tmpA, rtn))
'    Report_Name2$ = ini_Purge(tmp$, rtn)
    Report_Name2$ = InputINI("ImgSVR Host", "Report_Name2", SVR_Queue_INI$)
    
    '�����R341005����conclusion��쬰�ŮɡA�O�_�۰ʼu�X���z���PŪ�e���AIsPopSmart$<>"NO"�ɼu�X
'    rtn = ReadINI("ImgSVR Host", "IsPopSmart", "", tmpA, Len(tmpA), SVR_Queue_INI$)
'    If Not rtn = 0 Then tmp$ = Trim(Left(tmpA, rtn))
'    IsPopSmart$ = UCase(ini_Purge(tmp$, rtn))
    IsPopSmart$ = InputINI("ImgSVR Host", "IsPopSmart", SVR_Queue_INI$)
    
    '����d�ߵe�����ƧǥѤ���ܧ��ˬd�渹�AQ_SortOrder$="UNI_KEY"���ܧ󬰨��ˬd�渹�Ƨ�
'    rtn = ReadINI("ImgSVR Host", "Q_SortOrder", "", tmpA, Len(tmpA), SVR_Queue_INI$)
'    If Not rtn = 0 Then tmp$ = Trim(Left(tmpA, rtn))
'    Q_SortOrder$ = UCase(ini_Purge(tmp$, rtn))
    Q_SortOrder$ = InputINI("ImgSVR Host", "Q_SortOrder", SVR_Queue_INI$)
    
'    rtn = ReadINI("ImgSVR Host", "SiteEng", "", tmpA, Len(tmpA), SVR_Queue_INI$)
'    If Not rtn = 0 Then tmp$ = Trim(Left(tmpA, rtn))
'    global_SiteEng$ = ini_Purge(tmp$, rtn)
    global_SiteEng$ = InputINI("ImgSVR Host", "SiteEng", SVR_Queue_INI$)
    
'    rtn = ReadINI("ImgSVR Host", "Dr_from", "", tmpA, Len(tmpA), SVR_Queue_INI$)
'    If Not rtn = 0 Then Dr_from$ = ini_Purge(Left(tmpA, rtn), rtn)
    Dr_from$ = InputINI("ImgSVR Host", "Dr_from", SVR_Queue_INI$)
    
'    rtn = ReadINI("Image Server", "Destination", "", tmpA, Len(tmpA), SVR_Queue_INI$)
'    If Not rtn = 0 Then path_Images$ = Left(tmpA, rtn)
    path_Images$ = InputINI("Image Server", "Destination", SVR_Queue_INI$)
    If Right(path_Images$, 1) <> "\" Then
        path_Images$ = path_Images$ & "\"
    End If
'    rtn = ReadINI("ImgSVR Host", "Instrument", "", tmpA, Len(tmpA), SVR_Queue_INI$)
'    If Not rtn = 0 Then Instrument$ = Left(tmpA, rtn)
    Instrument$ = InputINI("ImgSVR Host", "Instrument", SVR_Queue_INI$)
    
'    rtn = ReadINI("Database", "Connection", "", tmpA, Len(tmpA), SVR_Queue_INI$)
'    If Not rtn = 0 Then dbConnection$ = Left(tmpA, rtn)
    dbConnection$ = InputINI("Database", "Connection", SVR_Queue_INI$)
    
'    rtn = ReadINI("Database", "Connection2", "", tmpA, Len(tmpA), SVR_Queue_INI$)
'    If Not rtn = 0 Then Target_Connection_String = Left(tmpA, rtn)
    Target_Connection_String = InputINI("Database", "Connection2", SVR_Queue_INI$)
    
'    rtn = ReadINI("Database", "NLS_LANG", "", tmpA, Len(tmpA), SVR_Queue_INI$)
'    If Not rtn = 0 Then NLS_LANG$ = Left(tmpA, rtn)
    NLS_LANG$ = InputINI("Database", "NLS_LANG", SVR_Queue_INI$)
    
'    rtn = ReadINI("Default Printer", "Device", "", tmpA, Len(tmpA), SVR_Queue_INI$)
'    If Not rtn = 0 Then ImgSVRPrinter = Left(tmpA, rtn)
    ImgSVRPrinter = InputINI("Default Printer", "Device", SVR_Queue_INI$)
    
'    rtn = ReadINI("Default Font", "Size", "", tmpA, Len(tmpA), SVR_Queue_INI$)
'    If Not rtn = 0 Then Font_Size% = Val(Left(tmpA, rtn))
    Font_Size% = Val(InputINI("Default Font", "Size", SVR_Queue_INI$))

'    rtn = ReadINI("Image Shell", "IMP_SCP", "", tmpA, Len(tmpA), SVR_Queue_INI$)
'    If Not rtn = 0 Then IMP_SCP$ = Left(tmpA, rtn)
    IMP_SCP$ = InputINI("Image Shell", "IMP_SCP", SVR_Queue_INI$)
    
    '�O�_�b�L�ޮv��ƮɡA�۰ʥH���i��v��J�ANeed_Dr_On=YES�ɦ۰ʶ�J
    Need_Dr_On$ = InputINI("ImgSVR Host", "Need_Dr_On", SVR_Queue_INI$)

End Sub


Function rtn_SubString(subHead$, originalStr$, stopChar$) As String
    Dim posStart As Integer, posEnd As Integer
    Dim tmp$
    
    rtn_SubString = ""
    
    If subHead$ = "*" Then
       posStop = InStr(originalStr$, stopChar$)
       If posStop > 0 Then
          rtn_SubString = Left(originalStr$, posStop - 1)
       Else
          rtn_SubString = originalStr$
       End If
    End If
    
End Function

Sub get_RecCount(formObj As Object)
       
    With formObj
    
         If Not .adoExam_Online.Recordset.EOF Then .adoExam_Online.Recordset.MoveLast 'If Not .datSource.Recordset.EOF Then .datSource.Recordset.MoveLast
         
         .lblTotal.Caption = "�O���`���� = " & str(.adoExam_Online.Recordset.RecordCount) & "  "
         If Not .adoExam_Online.Recordset.BOF Then .adoExam_Online.Recordset.MoveFirst
         DoEvents
'         Else
'            .lblTotal.Caption = "�O���`���� = 0"
'         End If

    End With
    
End Sub


Function NoNull(cString)
         If IsNull(cString) Then
            NoNull = ""
         Else
            NoNull = Trim(cString)
         End If
End Function


Sub Basic_get(tblName$, filter$)
    Dim adoDB As adoDB.Connection
    Dim adoTable As adoDB.Recordset
    Dim conn$, SQL$
    
    Set adoDB = New adoDB.Connection
    Set adoTable = New adoDB.Recordset
    adoDB.Open dbConnection$
    
    '/**/
    'rec$ = "SELECT ALL * FROM " & tblName$ & " WHERE " & filter$ 'DISTINCT Code FROM Predefines WHERE " & SQL$ & " Order by Code"
    '/**/
    rec$ = "SELECT ALL * FROM " & tblName$ & " WHERE " & filter$
    '/**/
    adoTable.Open rec$, adoDB, adOpenForwardOnly, adLockReadOnly 'tblName$
    
    If Not adoTable.EOF Then
       thisOne.chartno = NoNull(adoTable!chartno)
       thisOne.Name = NoNull(adoTable!Name)
       thisOne.CitizenID = NoNull(adoTable!CitizenID)
       thisOne.BirthDay = NoNull(adoTable!BirthDay)
       thisOne.Nationality = NoNull(adoTable!Nationality)
       thisOne.Sex = NoNull(adoTable!Sex)
       thisOne.Address = NoNull(adoTable!Address)
       thisOne.Phone = NoNull(adoTable!Phone)
       thisOne.Date_1st = NoNull(adoTable!Date_1st)
       thisOne.Recorder = NoNull(adoTable!Recorder)
       thisOne.Status = NoNull(adoTable!Status)
    Else
       thisOne.chartno = ""
       thisOne.Name = ""
       thisOne.CitizenID = ""
       thisOne.BirthDay = ""
       thisOne.Nationality = ""
       thisOne.Sex = ""
       thisOne.Address = ""
       thisOne.Phone = ""
       thisOne.Date_1st = ""
       thisOne.Recorder = ""
       thisOne.Status = ""
    End If
    
    adoTable.Close
    adoDB.Close
    
    Set adoTable = Nothing
    Set adoDB = Nothing
    
End Sub
Sub cmb_Table_Initial(dbsName$, tblName$, fieldName$, filter$, cmbControl As Object)
    Dim adoDB As New adoDB.Connection
    Dim adoTable As New adoDB.Recordset
    Dim conn$, SQL$
    
    
    adoDB.Open dbConnection$
    
    
    If Len(filter$) > 0 Then
       '/**/
       'rec$ = "SELECT " & fieldName$ & " FROM " & tblName$ & " WHERE " & filter$ 'DISTINCT Code FROM Predefines WHERE " & SQL$ & " Order by Code"
       '/**/
       'rec$ = "SELECT " & fieldName$ & " FROM " & tblName$ & " with(nolock) WHERE " & filter$ 'DISTINCT Code FROM Predefines WHERE " & SQL$ & " Order by Code"
       rec$ = "SELECT " & fieldName$ & " FROM " & tblName$ & " WHERE " & filter$ 'DISTINCT Code FROM Predefines WHERE " & SQL$ & " Order by Code"
       '/**/
    Else
       '/**/
       'rec$ = "SELECT " & fieldName$ & " FROM " & tblName$
       '/**/
       'rec$ = "SELECT " & fieldName$ & " FROM " & tblName$ & " with(nolock) "
       rec$ = "SELECT " & fieldName$ & " FROM " & tblName$
       '/**/
    End If
    rec$ = rec$
    adoTable.Open rec$, adoDB, adOpenForwardOnly, adLockReadOnly 'tblName$
    
    cmbControl.Clear
    Do While Not adoTable.EOF
       cmbControl.AddItem Trim(NoNull(adoTable(Replace(fieldName$, "DISTINCT ", ""))))
'       cmbControl.AddItem Trim(NoNull(adoTable(fieldName$)))
       adoTable.MoveNext
    Loop
    
    adoTable.Close
    adoDB.Close
    
    Set adoTable = Nothing
    Set adoDB = Nothing
End Sub

Function isRecordExist(tblName$, filter$) As Boolean
    Dim adoDB As New adoDB.Connection
    Dim adoTable As New adoDB.Recordset
    Dim conn$, SQL$
    
    '/**/
    'SQL$ = "SELECT ALL * FROM " & tblName$ & " " & filter$
    '/**/
    SQL$ = "SELECT ALL * FROM " & tblName$ & " " & filter$
    '/**/
    
    adoDB.Open dbConnection$
    adoTable.Open SQL$, adoDB, adOpenForwardOnly, adLockReadOnly
    
    isRecordExist = IIf(adoTable.EOF, False, True)
    
    adoTable.Close
    adoDB.Close
    
    Set adoTable = Nothing
    Set adoDB = Nothing
    
End Function
Sub array_Phrase_Initial()
    Dim SQL_Diction$, recCount&, i&
    Dim adoDB As New adoDB.Connection
    Dim adoTable As New adoDB.Recordset
    Dim conn$, SQL$
    Dim tmpClass$, j%
    
    '/**/
    'SQL_Diction$ = "SELECT Type, Class, Item, Code, Content FROM CRIS_Phrases " & _
                   "ORDER BY Type, Class, Item, Code"
    '/**/
    'SQL_Diction$ = "SELECT Type, Class, Item, Code, Content FROM CRIS_Phrases with(nolock) " & _
                   "ORDER BY Type, Class, Item, Code"
    SQL_Diction$ = "SELECT Type, Class, Item, Code, Content FROM CRIS_Phrases ORDER BY Type, Class, Item, Code"
    '/**/
    
    adoDB.Open dbConnection$
    adoTable.Open SQL_Diction$, adoDB, adOpenStatic, adLockReadOnly
    
    rcount = 0
    While Not adoTable.EOF
        adoTable.MoveNext
        rcount = rcount + 1
    Wend
    ReDim xPhrase(5, rcount)
    
    If rcount > 0 Then
        adoTable.MoveFirst
'    End If
'    If adoTable.RecordCount > 0 Then
'        ReDim xPhrase(5, adoTable.RecordCount)
        
        i& = 0
        Do While Not adoTable.EOF
           xPhrase(0, i&) = adoTable!Type
           xPhrase(1, i&) = adoTable!Class
           xPhrase(2, i&) = adoTable!item
           xPhrase(3, i&) = adoTable!Code
           xPhrase(4, i&) = NoNull(adoTable!Content)
           i& = i& + 1
           adoTable.MoveNext
        Loop
    End If
    
    adoTable.Close
    adoDB.Close
    
End Sub
Sub array_Diction_Initial()
    Dim SQL_Diction$, recCount&, i&
    Dim adoDB As New adoDB.Connection
    Dim adoTable As New adoDB.Recordset
    Dim conn$, SQL$
    Dim tmpClass$, j%
    
    ReDim xDiction(7, 1000, 6) 'recCount&)

    
    adoDB.Open dbConnection$
    
        tmpClass$ = "Chief Complain": j% = 0
        GoSub Import_Section
        
        tmpClass$ = "Finding": j% = 1
        GoSub Import_Section
        
        tmpClass$ = "Diagnosis": j% = 2
        GoSub Import_Section
        
        tmpClass$ = "Therapy": j% = 3
        GoSub Import_Section
        
        tmpClass$ = "Pathology": j% = 4
        GoSub Import_Section
        
        tmpClass$ = "Others": j% = 5
        GoSub Import_Section
    
    adoDB.Close
    
    Set adoTable = Nothing
    Set adoDB = Nothing
    
    Exit Sub
    
    
Import_Section:
    '/**/
    'SQL_Diction$ = "SELECT System, Type, Class, Code, Division, UserID, Content FROM CRIS_DICTIONARY WHERE Class='" & _
                   tmpClass$ & "' ORDER BY Type, Class, Code"
    '/**/
    'SQL_Diction$ = "SELECT System, Type, Class, Code, Division, UserID, Content FROM CRIS_DICTIONARY with(nolock) WHERE Class='" & _
                   tmpClass$ & "' ORDER BY Type, Class, Code"
    SQL_Diction$ = "SELECT System, Type, Class, Code, Division, UserID, Content FROM CRIS_DICTIONARY  WHERE Class='" & _
                   tmpClass$ & "' ORDER BY Type, Class, Code"
    '/**/
    
    adoTable.Open SQL_Diction$, adoDB, adOpenKeyset, adLockReadOnly
    i& = 0
    Do While Not adoTable.EOF
       xDiction(0, i&, j%) = adoTable!System
       xDiction(1, i&, j%) = adoTable!Type
       xDiction(2, i&, j%) = adoTable!Class
       xDiction(3, i&, j%) = adoTable!Code
       xDiction(4, i&, j%) = NoNull(adoTable!Division)
       xDiction(5, i&, j%) = NoNull(adoTable!UserID)
       xDiction(6, i&, j%) = NoNull(adoTable!Content)
       i& = i& + 1
       If i& >= 1000 Then Exit Do
       adoTable.MoveNext
    Loop
    
    If i& > xDictionMaxRows% Then xDictionMaxRows% = i&
    
    adoTable.Close
    Return
    
End Sub

Sub array_DictionbySpread_Initial(xUserID$)
    Dim SQL_Diction$, recCount&, i&
    Dim adoDB As New adoDB.Connection
    Dim adoTable As New adoDB.Recordset
    Dim conn$, SQL$
    Dim tmpClass$, j%
    
    
    adoDB.Open dbConnection$
    
    '/**/
    'SQL_Diction$ = "SELECT System, Type, Class, Code, Division, UserID, Content " & _
                   "FROM CRIS_DICTIONARY " & _
                   "WHERE System='" & "PublicPhrase" & "' AND " & _
                   "(UserID='" & xUserID$ & "' OR UserID='00000') " & _
                   "ORDER BY Type, Class, Code, UserID"
    '/**/
    SQL_Diction$ = "SELECT System, Type, Class, Code, Division, UserID, Content " & _
                   "FROM CRIS_DICTIONARY " & _
                   "WHERE System='" & "PublicPhrase" & "' AND " & _
                   "(UserID='" & xUserID$ & "' OR UserID='00000') " & _
                   "ORDER BY Type, Class, Code, UserID"
    '/**/
    
    adoTable.Open SQL_Diction$, adoDB, adOpenKeyset, adLockReadOnly
    
    rcount = 0
    While Not adoTable.EOF
        adoTable.MoveNext
        rcount = rcount + 1
    Wend
    ReDim xDictionarySpread(rcount)
    
    If rcount > 0 Then
        adoTable.MoveFirst
    End If
'    ReDim xDictionarySpread(adoTable.RecordCount)

    i& = 0
    Do While Not adoTable.EOF
       
       xDictionarySpread(i&).System = adoTable!System
       xDictionarySpread(i&).Type = adoTable!Type
       xDictionarySpread(i&).Class = adoTable!Class
       xDictionarySpread(i&).Code = adoTable!Code
       xDictionarySpread(i&).Division = NoNull(adoTable!Division)
       xDictionarySpread(i&).UserID = NoNull(adoTable!UserID)
       xDictionarySpread(i&).Content = NoNull(adoTable!Content)
       
       i& = i& + 1
       adoTable.MoveNext
    Loop
    
    adoTable.Close
    adoDB.Close
    
    Set adoTable = Nothing
    Set adoDB = Nothing
    
End Sub

Sub array_DictionByType_Initial(xType$)
    Dim i%, j%, k%
    
    ReDim xDictionByType(6, xDictionMaxRows%, 6)

    For j% = 0 To 6
        k% = 0
        For i% = 0 To 1000
            If xDiction(0, i%, j%) = "" Then
               i% = 1001
            Else
            If xDiction(1, i%, j%) = xType$ Or xDiction(1, i%, j%) = "�@��" Then
                xDictionByType(0, k%, j%) = xDiction(0, i%, j%) 'Exam_Type
                xDictionByType(1, k%, j%) = xDiction(1, i%, j%) 'Class
                xDictionByType(2, k%, j%) = xDiction(2, i%, j%) 'Code
                xDictionByType(3, k%, j%) = xDiction(3, i%, j%) 'Content)
                xDictionByType(4, k%, j%) = xDiction(4, i%, j%) 'Content)
                xDictionByType(5, k%, j%) = xDiction(5, i%, j%) 'Content)
                xDictionByType(6, k%, j%) = xDiction(6, i%, j%) 'Content)
                k% = k% + 1
            End If
            End If
        Next
    Next
    
    
End Sub

Function get_NewFileName() As String
    Dim NameYear As String * 1
    Dim NameMonth, NameDay, NameHour As String * 1
    Dim NameMin As String, NameSec As String

    NameYear = Right(Format(Date, "yyyy"), 1)
    NameMonth = CodeSet(Val(Format(Date, "mm")))
    NameDay = CodeSet(Val(Format(Date, "dd")))
    NameHour = CodeSet(Val(Format(Now, "hh")))
    
    NameMin = Trim(Minute(time))
    If Len(NameMin) = 1 Then NameMin = "0" & NameMin
    
    NameSec = Trim(Second(time))
    If Len(NameSec) = 1 Then NameSec = "0" & NameSec

    get_NewFileName = NameYear & NameMonth & NameDay & NameHour & NameMin & NameSec

End Function
Sub CodeSet_Define()
    
    CodeSet(0) = "0": CodeSet(1) = "1": CodeSet(2) = "2"
    CodeSet(3) = "3": CodeSet(4) = "4": CodeSet(5) = "5"
    CodeSet(6) = "6": CodeSet(7) = "7": CodeSet(8) = "8"
    CodeSet(9) = "9"
    CodeSet(10) = "A": CodeSet(11) = "B": CodeSet(12) = "C"
    CodeSet(13) = "D": CodeSet(14) = "E": CodeSet(15) = "F"
    CodeSet(16) = "G": CodeSet(17) = "H": CodeSet(18) = "I"
    CodeSet(19) = "J"
    CodeSet(20) = "K": CodeSet(21) = "L": CodeSet(22) = "M"
    CodeSet(23) = "N": CodeSet(24) = "O": CodeSet(25) = "P"
    CodeSet(26) = "Q": CodeSet(27) = "R": CodeSet(28) = "S"
    CodeSet(29) = "T"
    CodeSet(30) = "U": CodeSet(31) = "V": CodeSet(32) = "W"
    CodeSet(33) = "X": CodeSet(34) = "Y": CodeSet(35) = "Z"

End Sub
Sub WaitForEventsToFinish(NbrTimes As Integer)
   Dim i As Integer
   
   For i = 1 To NbrTimes
      Dummy% = DoEvents()
   Next i
   
End Sub
Function HotShot(sType$, sClass$, Scode$) As String
    Dim CutLen, KeyLen, i As Integer
    Dim adoDB As New adoDB.Connection
    Dim adoTable As New adoDB.Recordset
    Dim conn$ ', SQL$
    
    '/**/
    'SQL$ = "SELECT ALL * FROM CRIS_Dictionary WHERE (Type='" & sType$ & "' OR Type='�@��') AND Class='" & sClass$ & "' AND Code='" & Scode$ & "'"
    '/**/
    SQL$ = "SELECT ALL * FROM CRIS_Dictionary WHERE (Type='" & sType$ & "' OR Type='�@��') AND Class='" & sClass$ & "' AND Code='" & Scode$ & "'"
    '/**/
    
    adoDB.Open dbConnection$
    adoTable.Open SQL$, adoDB, adOpenForwardOnly, adLockReadOnly
    
    DoEvents
    If Not adoTable.EOF Then
       HotShot = Trim(adoTable!Content)
    Else
       HotShot = ""
    End If
    
    adoTable.Close
    adoDB.Close
    
    Set adoTable = Nothing
    Set adoDB = Nothing

End Function

Function Check_CitizenID(ByRef pId As String) As Boolean                           '�ˬd�����Ҧr��
  Dim tmpString As String
  Dim chkSum As Integer
  Dim II As Integer
  
  Check_CitizenID = False
  pId = UCase(pId)                                                      '�N�����Ҧr���ର�j�g
  If Len(pId) <> 10 Then Exit Function                                  '���׬��Q�X
  If Mid(pId, 1, 1) < "A" And Mid(pId, 1, 1) > "Z" Then Exit Function   '�Ĥ@�X���^��
  If Mid(pId, 2, 1) <> "1" And Mid(pId, 2, 1) <> "2" Then Exit Function '�ĤG�X�� 1 �� 2
  For II = 3 To Len(pId)
    If Mid(pId, II, 1) < "0" Or Mid(pId, II, 1) > "9" Then Exit Function    '�@�w���Ʀr�X
  Next
  tmpString = Trim(str((InStr("ABCDEFGHJKLMNPQRSTUVWXYZIO", Mid(pId, 1, 1)) + 9))) & Mid(pId, 2, 9)
  chkSum = 0
  For II = 1 To Len(tmpString)
    Select Case II
      Case 1
        chkSum = chkSum + Val(Mid(tmpString, II, 1)) * 1
      Case 11
        chkSum = chkSum + Val(Mid(tmpString, II, 1)) * 1
      Case Else
        chkSum = chkSum + Val(Mid(tmpString, II, 1)) * (11 - II)
    End Select
  Next
  If chkSum Mod 10 <> 0 Then Exit Function
  Check_CitizenID = True
End Function


Sub xDraft_Get(ImgOwner$, ImgType$, CreateDate$, xArray() As typeImage)
    Dim adoDB As New adoDB.Connection
    Dim adoTable As New adoDB.Recordset
    Dim conn$, i% ', SQL$
    
    '/**/
    'SQL$ = "SELECT ALL * FROM CRIS_Images_Reference WHERE ImgOwner='" & ImgOwner$ & "' AND Class='DRAFT' AND Type='" & ImgType$ & "' AND CreateDate='" & CreateDate$ & "'"
    '/**/
    SQL$ = "SELECT ALL * FROM CRIS_Images_Reference WHERE ImgOwner='" & ImgOwner$ & "' AND Class='DRAFT' AND Type='" & ImgType$ & "' AND CreateDate='" & CreateDate$ & "'"
    '/**/
    adoDB.Open dbConnection$
    adoTable.Open SQL$, adoDB, adOpenStatic, adLockReadOnly
    
    DoEvents
    
    rcount = 0
    While Not adoTable.EOF
        adoTable.MoveNext
        rcount = rcount + 1
    Wend
    ReDim xArray(rcount)
    
    If rcount > 0 Then
        adoTable.MoveFirst
    End If
'    If Not adoTable.EOF Then
'       adoTable.MoveLast
'       ReDim xArray(adoTable.RecordCount)
'       adoTable.MoveFirst
'    Else
'       ReDim xArray(0)
'    End If
    
    i% = 0
    Do While Not adoTable.EOF
       xArray(i%).ImgOwner = adoTable("ImgOwner")
       xArray(i%).Class = NoNull(adoTable("Class"))
       xArray(i%).Type = NoNull(adoTable("Type"))
       xArray(i%).GroupName = NoNull(adoTable("GroupName"))
       xArray(i%).GroupIndex = NoNull(adoTable("GroupIndex"))
       xArray(i%).CreateDate = NoNull(adoTable("CreateDate"))
       xArray(i%).CreateTime = NoNull(adoTable("CreateTime"))
       xArray(i%).FilePath = NoNull(adoTable("FilePath"))
       xArray(i%).FileName = NoNull(adoTable("FileName"))
       xArray(i%).CreateBy = NoNull(adoTable("CreateBy"))
       xArray(i%).Memo = NoNull(adoTable("Memo"))
       
       i% = i% + 1
       adoTable.MoveNext
    Loop
    
    LastDraftNo = i%
    
    adoTable.Close
    adoDB.Close
    Set adoTable = Nothing
    Set adoDB = Nothing
    
   On Error GoTo 0
    DoEvents
    Exit Sub
    
ImgLoadError:
    If err = 53 Then
       imgDraft.Picture = LoadPicture(path_Define$ & adoTable!Type & ".bmp")
       SavePicture imgDraft.Picture, xArray(0).FilePath & xArray(0).FileName
    End If
    Resume
    

End Sub



Sub Draft_Get(ImgOwner$, xUni_key$)
    Dim adoDB As New adoDB.Connection
    Dim adoTable As New adoDB.Recordset
    Dim conn$ ', SQL$
    
    On Error GoTo ImgLoadError
    
    '/**/
    'SQL$ = "SELECT ALL * FROM CRIS_Images_Reference WHERE ImgOwner='" & ImgOwner$ & "' AND Class='DRAFT' AND Uni_key='" & xUni_key$ & "'"
    '/**/
    SQL$ = "SELECT ALL * FROM CRIS_Images_Reference WHERE ImgOwner='" & ImgOwner$ & "' AND Class='DRAFT' AND Uni_key='" & xUni_key$ & "'"
    '/**/
    
    adoDB.Open dbConnection$
    adoTable.Open SQL$, adoDB, adOpenStatic, adLockReadOnly
    
    DoEvents
    
    rcount = 0
    While Not adoTable.EOF
        adoTable.MoveNext
        rcount = rcount + 1
    Wend
    ReDim Draft_Array(rcount)
    
    If rcount > 0 Then
        adoTable.MoveFirst
    End If
'    If Not adoTable.EOF Then
'       adoTable.MoveLast
'       ReDim Draft_Array(adoTable.RecordCount)
'       adoTable.MoveFirst
'    Else
'       ReDim Draft_Array(0)
'    End If
    
    i% = 0
    Do While Not adoTable.EOF
       Draft_Array(i%).ImgOwner = adoTable("ImgOwner")
       Draft_Array(i%).Class = NoNull(adoTable("Class"))
       Draft_Array(i%).Type = NoNull(adoTable("Type"))
       Draft_Array(i%).GroupName = NoNull(adoTable("GroupName"))
       Draft_Array(i%).GroupIndex = NoNull(adoTable("GroupIndex"))
       Draft_Array(i%).CreateDate = NoNull(adoTable("CreateDate"))
       Draft_Array(i%).CreateTime = NoNull(adoTable("CreateTime"))
       Draft_Array(i%).FilePath = NoNull(adoTable("FilePath"))
       Draft_Array(i%).FileName = NoNull(adoTable("FileName"))
       Draft_Array(i%).CreateBy = NoNull(adoTable("CreateBy"))
       Draft_Array(i%).Memo = NoNull(adoTable("Memo"))
       
       i% = i% + 1
       adoTable.MoveNext
    Loop
    
    LastDraftNo = i%
    
    adoTable.Close
    adoDB.Close
    Set adoTable = Nothing
    Set adoDB = Nothing
    
    On Error GoTo 0
    DoEvents
    Exit Sub
    
ImgLoadError:

    If err = 53 Then
       imgDraft.Picture = LoadPicture(path_Define$ & curr_Record.Type & ".bmp")
       SavePicture imgDraft.Picture, Draft_Array(0).FilePath & Draft_Array(0).FileName
    'Else
    '   MsgBox Error(Err)
    End If
    Resume Next
    
End Sub

Public Function GetScopyTime(chartno, uni_key, scopyserial) As String
    Dim adoDB As New adoDB.Connection
    Dim adoTable As New adoDB.Recordset
    Dim conn$, t$, ts$, te$
    Dim intHrs As Integer, intMins As Integer, intSecs As Integer
    
    t$ = "00:00:00"
    SQL$ = "select * from cris_scopy_online "
    SQL$ = SQL$ & " where chartno = '" & chartno & "' and uni_key = '" & uni_key & "' "
    SQL$ = SQL$ & " and scopyserial = '" & scopyserial & "' "
    SQL$ = SQL$ & " order by scopyorder "
    
    adoDB.Open dbConnection$
    adoTable.Open SQL$, adoDB, adOpenStatic, adLockReadOnly
    intHrs = 0
    intMins = 0
    intSecs = 0
    While Not adoTable.EOF
        '�Y�v���^���ɤ����`�����A�i��scopyendtime�|���ťաA�h�ӵ����p��
        '�Yscopystarttime�]���ťաA�h�ӵ��]���p��
        ts$ = NoNull(adoTable("scopystarttime"))
        te$ = NoNull(adoTable("scopyendtime"))
        If ts$ <> "" And te$ <> "" Then
            intHrs = intHrs + DateDiff("h", TimeValue(ts$), TimeValue(te$))
            intMins = intMins + DateDiff("n", TimeValue(ts$), TimeValue(te$)) Mod 60
            intSecs = intSecs + DateDiff("s", TimeValue(ts$), TimeValue(te$)) Mod 60
        End If
        adoTable.MoveNext
    Wend
    intMins = intMins + intSecs \ 60
    intSecs = intSecs Mod 60
    intHrs = intHrs + intMins \ 60
    intMins = intMins Mod 60
    t$ = Format(intHrs, "00") & ":" & Format(intMins, "00") & ":" & Format(intSecs, "00")
    adoTable.Close
    adoDB.Close
    Set adoTable = Nothing
    Set adoDB = Nothing
    GetScopyTime = t$
End Function

Public Function GetRangeTime(chartno, uni_key) As String
    Dim adoDB As New adoDB.Connection
    Dim adoTable As New adoDB.Recordset
    Dim conn$, examtime$
    
'    t$ = ""
    examtime$ = ""
    SQL$ = "select scopeorder, sumtime from cris_capture_range "
    SQL$ = SQL$ & " where chartno = '" & chartno & "' and uni_key = '" & uni_key & "' "
    SQL$ = SQL$ & " order by scopeorder "
    adoDB.Open dbConnection$
    adoTable.Open SQL$, adoDB, adOpenStatic, adLockReadOnly
    
    While Not adoTable.EOF
        If examtime$ <> "" Then
'            t$ = t$ & ", "
            examtime$ = examtime$ & ", "
        End If
'        t$ = t$ & NoNull(adoTable("scopeorder"))
        examtime$ = examtime$ & NoNull(adoTable("sumtime"))
'        examtime = examtime & GetScopyTime(chartno, uni_key, NoNull(adoTable("scopyserial")))
        adoTable.MoveNext
    Wend
    adoTable.Close
    adoDB.Close
    Set adoTable = Nothing
    Set adoDB = Nothing
    GetRangeTime = examtime$
End Function

Public Sub Image_Get(ImgOwner$, ImgUni_key$)
    Dim adoDB As New adoDB.Connection
    Dim adoTable As New adoDB.Recordset
    Dim conn$, SQL$
    Dim tmp() As String
    Dim i As Integer
    Dim tflag As Boolean
    
'    SQL$ = "SELECT ALL * FROM CRIS_Images_Reference " & _
           "WHERE ImgOwner='" & ImgOwner$ & "' AND Class<>'DRAFT' AND Uni_key='" & ImgUni_key$ & "' " & _
           "ORDER BY CreateTime"
    '/**/
    'SQL$ = "SELECT ALL * FROM CRIS_Images_Reference " & _
           "WHERE ImgOwner='" & ImgOwner$ & "' AND Class<>'DRAFT' AND Uni_key='" & ImgUni_key$ & "' " & _
           "ORDER BY FileName"
    '/**/
    'SQL$ = "SELECT ALL * FROM CRIS_Images_Reference with(nolock) " & _
           "WHERE ImgOwner='" & ImgOwner$ & "' AND Class<>'DRAFT' AND Uni_key='" & ImgUni_key$ & "' " & _
           "ORDER BY FileName"
    '/**/
    SQL$ = "SELECT ALL * FROM CRIS_Images_Reference "
    SQL$ = SQL$ & " WHERE ImgOwner='" & ImgOwner$ & "' AND Class<>'DRAFT' "
    
    If Len(xSee_OCR$) < 1 Then
        SQL$ = SQL$ & " and Class <> 'OCR' "
    Else
        tflag = True
        tmp = Split(xSee_OCR$, ";")
        For i = 0 To UBound(tmp)
            If Trim(tmp(i)) = curr_Record.Type Then
                tflag = False
            End If
        Next
        If tflag Then
            SQL$ = SQL$ & " and Class <> 'OCR' "
        End If
    End If
    
    SQL$ = SQL$ & " AND Uni_key='" & ImgUni_key$ & "' "
    SQL$ = SQL$ & " and (backup <= '2' or backup is null) ORDER BY createdate, createtime, groupindex "
    '/**/
    
    
    adoDB.Open dbConnection$
    adoTable.Open SQL$, adoDB, adOpenStatic, adLockReadOnly
    
    DoEvents
    
    rcount = 0
    While Not adoTable.EOF
        adoTable.MoveNext
        rcount = rcount + 1
    Wend
'    If InStr(Trim(curr_Record.ExamDetail), "18019") > 0 Then
'        rcount = 0
'    End If
    
    ReDim Img_Array(rcount)
    
    If rcount > 0 Then
        adoTable.MoveFirst
    End If
'    If Not adoTable.EOF Then
'       adoTable.MoveLast
'       ReDim Img_Array(adoTable.RecordCount)
'       adoTable.MoveFirst
'    Else
'       ReDim xArray(0)
'    End If
    
    i% = 0
    Do While Not adoTable.EOF
       Img_Array(i%).ImgOwner = adoTable("ImgOwner")
       Img_Array(i%).Class = NoNull(adoTable("Class"))
       Img_Array(i%).Type = NoNull(adoTable("Type"))
       Img_Array(i%).GroupName = NoNull(adoTable("GroupName"))
       Img_Array(i%).GroupIndex = NoNull(adoTable("GroupIndex"))
       Img_Array(i%).CreateDate = NoNull(adoTable("CreateDate"))
       Img_Array(i%).CreateTime = NoNull(adoTable("CreateTime"))
       Img_Array(i%).FilePath = NoNull(adoTable("FilePath"))
       Img_Array(i%).FileName = NoNull(adoTable("FileName"))
       Img_Array(i%).CreateBy = NoNull(adoTable("CreateBy"))
       Img_Array(i%).Backup = NoNull(adoTable("Backup"))
       If NoNull(adoTable("Multiframe")) <> "1" Then
            Img_Array(i%).MutliFrame = False
       Else
            Img_Array(i%).MutliFrame = True
       End If
       i% = i% + 1
       adoTable.MoveNext
    Loop
    
    LastPageNo = i%
    
    adoTable.Close
    adoDB.Close
    Set adoTable = Nothing
    Set adoDB = Nothing
    
    DoEvents
    
End Sub

Sub xImage_Get(ImgOwner$, ImgType$, CreateDate$, xArray() As typeImage)
    Dim adoDB As New adoDB.Connection
    Dim adoTable As New adoDB.Recordset
    Dim conn$ ', SQL$
    
    '/**/
    'SQL$ = "SELECT ALL * FROM CRIS_Images_Reference WHERE ImgOwner='" & ImgOwner$ & "' AND Class<>'DRAFT' AND Type='" & ImgType$ & "' AND CreateDate='" & CreateDate$ & "' ORDER BY CreateTime"
    '/**/
    SQL$ = "SELECT ALL * FROM CRIS_Images_Reference WHERE ImgOwner='" & ImgOwner$ & "' AND Class<>'DRAFT' AND Type='" & ImgType$ & "' AND CreateDate='" & CreateDate$ & "' ORDER BY CreateTime"
    '/**/
    adoDB.Open dbConnection$
    adoTable.Open SQL$, adoDB, adOpenStatic, adLockReadOnly
    
    DoEvents
    
    rcount = 0
    While Not adoTable.EOF
        adoTable.MoveNext
        rcount = rcount + 1
    Wend
    ReDim xArray(rcount)
    
    If rcount > 0 Then
        adoTable.MoveFirst
    End If
'    If Not adoTable.EOF Then
'       adoTable.MoveLast
'       ReDim xArray(adoTable.RecordCount)
'       adoTable.MoveFirst
'    Else
'       ReDim xArray(0)
'    End If
    
    i% = 0
    Do While Not adoTable.EOF
       xArray(i%).ImgOwner = adoTable("ImgOwner")
       xArray(i%).Class = NoNull(adoTable("Class"))
       xArray(i%).Type = NoNull(adoTable("Type"))
       xArray(i%).GroupName = NoNull(adoTable("GroupName"))
       xArray(i%).GroupIndex = NoNull(adoTable("GroupIndex"))
       xArray(i%).CreateDate = NoNull(adoTable("CreateDate"))
       xArray(i%).CreateTime = NoNull(adoTable("CreateTime"))
       xArray(i%).FilePath = NoNull(adoTable("FilePath"))
       xArray(i%).FileName = NoNull(adoTable("FileName"))
       xArray(i%).CreateBy = NoNull(adoTable("CreateBy"))
       
       i% = i% + 1
       adoTable.MoveNext
    Loop
    
    LastPageNo = i%
    
    
    adoTable.Close
    adoDB.Close
    Set adoTable = Nothing
    Set adoDB = Nothing
    
    DoEvents
    
End Sub

Sub xReportTemplate_Get()
    Dim adoDB As New adoDB.Connection
    Dim adoTable As New adoDB.Recordset
    Dim conn$ ', SQL$
    Dim rcount As Integer
    
    '/**/
    'SQL$ = "SELECT ALL * FROM CRIS_ReportTemplate ORDER BY UserID, DivisionID, ExamID"
    '/**/
    'SQL$ = "SELECT ALL * FROM CRIS_ReportTemplate with(nolock) ORDER BY UserID, DivisionID, ExamID"
    SQL$ = "SELECT all * FROM CRIS_ReportTemplate ORDER BY USERID, DivisionID, ExamID"
    '/**/
    adoDB.Open dbConnection$
    adoTable.Open SQL$, adoDB, adOpenStatic, adLockReadOnly
    
    DoEvents
    
    rcount = 0
    While Not adoTable.EOF
        adoTable.MoveNext
        rcount = rcount + 1
    Wend
    
    If rcount > 0 Then
        ReDim xReportTemplate(rcount - 1)
        adoTable.MoveFirst
    End If
'    If Not adoTable.EOF Then
'       adoTable.MoveLast
'       ReDim xReportTemplate(adoTable.RecordCount)
'       adoTable.MoveFirst
'    Else
'       ReDim xReportTemplate(0)
'    End If
    
    i% = 0
    Do While Not adoTable.EOF
       xReportTemplate(i%).DivisionID = adoTable("DivisionID")
       xReportTemplate(i%).DivisionName = NoNull(adoTable("DivisionName"))
       xReportTemplate(i%).ExamID = NoNull(adoTable("ExamID"))
       xReportTemplate(i%).ExamName = NoNull(adoTable("ExamName"))
       xReportTemplate(i%).TemplateFileName = NoNull(adoTable("TemplateFileName"))
       xReportTemplate(i%).TemplateFileSource = NoNull(adoTable("TemplateFileSource"))
       xReportTemplate(i%).UserID = NoNull(adoTable("UserID"))
       xReportTemplate(i%).DefaultUse = NoNull(adoTable("DefaultUse"))
       xReportTemplate(i%).ExamDescription = NoNull(adoTable("ExamDescription"))
       
       i% = i% + 1
       adoTable.MoveNext
    Loop
    
    LastPageNo = i%
    
    
    adoTable.Close
    adoDB.Close
    Set adoTable = Nothing
    Set adoDB = Nothing
    
    DoEvents
    
End Sub

Function isFileExist(File$, attr%) As Integer
    
    On Error GoTo existError
    
    Result = Dir(File$, attr%)
    If Len(Result) > 0 Then
       isFileExist = True
    Else
       isFileExist = False
    End If
    Exit Function

existError:
    isFileExist = False
    
End Function

Function Get_Uni_key(sType$) As String
    Dim adoDB As New adoDB.Connection
    Dim adoTable As New adoDB.Recordset
    Dim SQL$, i%, sCounter&, sAbbreviation$
    
    '/**/
    'SQL$ = "SELECT ALL * FROM CRIS_ExamType WHERE Type='" & sType$ & "'"
    '/**/
    SQL$ = "SELECT ALL * FROM CRIS_ExamType WHERE Type='" & sType$ & "'"
    '/**/
    adoDB.Open dbConnection$
    adoTable.Open SQL$, adoDB, adOpenDynamic, adLockPessimistic
    
    DoEvents
    
    'adoTable.Find "Type='" & sType$ & "'"
    If Not adoTable.EOF Then
        sAbbreviation$ = adoTable!abbreviation
        sCounter& = adoTable!Counter
        
       'adoTable.Edit
       adoTable!Counter = adoTable!Counter + 1
       adoTable.Update
        Get_Uni_key = sAbbreviation$ & Trim(str(sCounter&))
    Else
       Get_Uni_key = ""
    End If
    
    adoTable.Close
    adoDB.Close
    Set adoTable = Nothing
    Set adoDB = Nothing
    
End Function

Function Check_DBServer() As Boolean
    Dim adoDB As New adoDB.Connection
    Dim adoTable As New adoDB.Recordset
    Dim conn$, SQL$
    
    On Error GoTo db_error
    adoDB.Open dbConnection$
     'rec$ = "SELECT ALL * FROM Td_def_detail"
    'adoTable.Open rec$, adoDB
    
    'adoTable.Close
    adoDB.Close
    
    'Set adoTable = Nothing
    Set adoDB = Nothing
    Check_DBServer = True
    Exit Function
    
db_error:
    MsgBox Error(err)
    Check_DBServer = False
    
End Function
Function Check_ChartNo(sChartNo$) As Boolean
    Dim psw_chk&, tmpNo%(8), pwdLast$
    
    For i% = 1 To 8
       tmpNo%(i% - 1) = Val(Mid(sChartNo$, i%, 1))
    Next
    
    For i% = 1 To 7 Step 2
        psw_chk& = psw_chk& + tmpNo%(i% - 1)
    Next
    
    psw_chk& = psw_chk& * 3 + tmpNo%(1) + tmpNo%(3) + tmpNo%(5)
    pwdLast$ = Trim(str(psw_chk&))
    pwdLast$ = Right(pwdLast$, 1)
    
    If tmpNo%(7) = Val(Right(str(10 - Val(pwdLast$)), 1)) Then
        Check_ChartNo = True
    Else
        Check_ChartNo = False
    End If
    
End Function
Sub cmb_DR_Initial(dbsName$, tblName$, fieldName$, filter$, cmbControl As Object)
    Dim adoDB As New adoDB.Connection
    Dim adoTable As New adoDB.Recordset
    Dim conn$, SQL$
    
    adoDB.Open dbConnection$
    
    If Len(filter$) > 0 Then
       '/**/
       'rec$ = "SELECT ALL * FROM " & tblName$ & " WHERE " & filter$
       '/**/
       'rec$ = "SELECT ALL * FROM " & tblName$ & " with(nolock) WHERE " & filter$
       rec$ = "SELECT ALL * FROM " & tblName$ & " WHERE " & filter$ & " order by UserID "
       '/**/
    Else
       '/**/
       'rec$ = "SELECT ALL * FROM " & tblName$
       '/**/
       'rec$ = "SELECT ALL * FROM " & tblName$ & " with(nolock) "
       rec$ = "SELECT ALL * FROM " & tblName$ & " order by UserID "
       '/**/
    End If
    rec$ = rec$
    adoTable.Open rec$, adoDB, adOpenForwardOnly, adLockReadOnly 'tblName$
    
    cmbControl.Clear
    Do While Not adoTable.EOF
       cmbControl.AddItem Trim(NoNull(adoTable!UserID)) & Trim(NoNull(adoTable(fieldName$)))
       'cmbControl.AddItem Trim(NoNull(adoTable(fieldName$)))
       adoTable.MoveNext
    Loop
    
    adoTable.Close
    adoDB.Close
    
    Set adoTable = Nothing
    Set adoDB = Nothing
End Sub

Sub cmb_Division_Initial(dbsName$, tblName$, fieldName$, filter$, cmbControl As Object)
    Dim adoDB As New adoDB.Connection
    Dim adoTable As New adoDB.Recordset
    Dim conn$, SQL$
    
    adoDB.Open dbConnection$
    
    If Len(filter$) > 0 Then
       '/**/
       'rec$ = "SELECT ALL * FROM " & tblName$ & " WHERE " & filter$
       '/**/
       rec$ = "SELECT ALL * FROM " & tblName$ & " WHERE " & filter$
       '/**/
    Else
       '/**/
       'rec$ = "SELECT ALL * FROM " & tblName$
       '/**/
       rec$ = "SELECT ALL * FROM " & tblName$ & " "
       '/**/
    End If
    adoTable.Open rec$, adoDB, adOpenForwardOnly, adLockReadOnly 'tblName$
    
    cmbControl.Clear
    Do While Not adoTable.EOF
'       cmbControl.AddItem Trim(NoNull(adoTable!Code)) & Trim(NoNull(adoTable!Remark))
       cmbControl.AddItem Trim(NoNull(adoTable!Remark))
       adoTable.MoveNext
    Loop
    
    adoTable.Close
    adoDB.Close
    
    Set adoTable = Nothing
    Set adoDB = Nothing
End Sub

Sub cmb_ExamDetail_Initial(dbsName$, tblName$, fieldName$, filter$, cmbControl As Object)
    Dim adoDB As New adoDB.Connection
    Dim adoTable As New adoDB.Recordset
    Dim conn$, SQL$, tmp$
    
    adoDB.Open dbConnection$
    
    If Len(filter$) > 0 Then
       '/**/
       'rec$ = "SELECT ALL * FROM " & tblName$ & " WHERE " & filter$
       '/**/
       'rec$ = "SELECT ALL * FROM " & tblName$ & " with(nolock) WHERE " & filter$
       rec$ = "SELECT ALL * FROM " & tblName$ & " WHERE " & filter$
       '/**/
    Else
       '/**/
       'rec$ = "SELECT ALL * FROM " & tblName$
       '/**/
       'rec$ = "SELECT ALL * FROM " & tblName$ & " with(nolock) "
       rec$ = "SELECT ALL * FROM " & tblName$
       '/**/
    End If
    'rec$ = rec$ & " ORDER BY " & fieldName$
    rec$ = rec$ & " ORDER BY " & fieldName$
    adoTable.Open rec$, adoDB, adOpenForwardOnly, adLockReadOnly 'tblName$
    
    cmbControl.Clear
    Do While Not adoTable.EOF
        tmp$ = Trim(NoNull(adoTable!Code))
        If Len(tmp$) < 10 Then
           tmp$ = tmp$ & String(10 - Len(tmp$), " ")
        End If
        'tmp$ = Format(tmp$, "          ")

        cmbControl.AddItem UCase(tmp$) & ": " & Trim(NoNull(adoTable(fieldName$)))
        'cmbControl.AddItem Trim(NoNull(adoTable(fieldName$)))
       
        adoTable.MoveNext
    Loop
    
    adoTable.Close
    adoDB.Close
    
    Set adoTable = Nothing
    Set adoDB = Nothing
End Sub

Function Report_Adjust(ctl As TextBox) As String

    Dim LineCnt As Long, pos As Long, Length As Integer, i%
    Dim s As String
    
    Report_Adjust = ""
    LineCnt = TextBoxLineCnt(ctl)
    For i% = 0 To LineCnt - 1
        pos = SendMessage(ctl.hWnd, EM_LINEINDEX, i%, ByVal 0&)
        Length = SendMessage(ctl.hWnd, EM_LINELENGTH, pos, ByVal 0&)
        s = String(Length, Chr(0))
        
        CopyMemory ByVal s, Length, 2
        SendMessage ctl.hWnd, EM_GETLINE, i%, ByVal s
        Report_Adjust = Report_Adjust & s & Chr(13) & Chr(10)
    Next
    

End Function

Public Function TextBoxLineCnt(ctl As TextBox) As Long
    TextBoxLineCnt = SendMessage(ctl.hWnd, EM_GETLINECOUNT, 0, 0)
End Function

Function Field_get(dbsName$, tblName$, fieldName$, filter$) As Variant
    Dim adoDB As adoDB.Connection
    Dim adoTable As adoDB.Recordset
    Dim conn$, SQL$
    
    Set adoDB = New adoDB.Connection
    Set adoTable = New adoDB.Recordset
    adoDB.Open dbConnection$
    
    If Len(filter$) > 0 Then
       '/**/
       'rec$ = "SELECT " & fieldName$ & " FROM " & tblName$ & " WHERE " & filter$ 'DISTINCT Code FROM Predefines WHERE " & SQL$ & " Order by Code"
       '/**/
       rec$ = "SELECT " & fieldName$ & " FROM " & tblName$ & " WHERE " & filter$
       '/**/
    Else
       '/**/
       'rec$ = "SELECT " & fieldName$ & " FROM " & tblName$
       '/**/
       rec$ = "SELECT " & fieldName$ & " FROM " & tblName$ & " "
       '/**/
    End If
    adoTable.Open rec$, adoDB, adOpenForwardOnly, adLockReadOnly 'tblName$
    
    If Not adoTable.EOF Then
       Field_get = NoNull(adoTable(fieldName$))
    Else
       Field_get = ""
    End If
    
    adoTable.Close
    adoDB.Close
    
    Set adoTable = Nothing
    Set adoDB = Nothing
    
End Function
Function getDoctorFromUser(drid$) As String
    Dim adoDB As adoDB.Connection
    Dim adoTable As adoDB.Recordset
    Dim conn$, SQL$
    
    Set adoDB = New adoDB.Connection
    Set adoTable = New adoDB.Recordset
    
    adoDB.Open dbConnection$
    
'    SQL$ = "SELECT Code,Remark FROM CRIS_Reference WHERE Class='Doctor' AND Code='" & drid$ & "'"
    '/**/
    'SQL$ = "SELECT UserID, Name FROM CRIS_User WHERE UserID='" & drid$ & "'"
    '/**/
    SQL$ = "SELECT UserID, Name FROM CRIS_User WHERE UserID='" & drid$ & "'"
    '/**/
    adoTable.Open SQL$, adoDB, adOpenForwardOnly, adLockReadOnly
    
    If Not adoTable.EOF Then
       getDoctorFromUser = UCase(drid$) & NoNull(adoTable!Name)
    Else
       getDoctorFromUser = UCase(drid$)
    End If
    
    adoTable.Close
    Set adoTable = Nothing
    
    adoDB.Close
    Set adoDB = Nothing
    
End Function

Sub setPrnFormNoTitle()

    prnFormString$ = "<title>ImageSVR9 $Type$</title>" & _
                     "<meta http-equiv=""Content-Type"" content=""text/html; charset=big5"" />" & _
                     "<meta http-equiv=""Content-Language"" content=""big5"" />"
    prnFormString$ = prnFormString$ & "<BODY LINK=""#0000ff"" VLINK=""#800080"">"
    
'    prnFormString$ = prnFormString$ & _
            "<TABLE CELLSPACING=0 BORDER=0 CELLPADDING=1 WIDTH=""100%"" ALIGN=""CENTER""><TR>" & _
            "<TD WIDTH=""23%"" VALIGN=""TOP"" ROWSPAN=3><FONT FACE=""�з���"" SIZE=1></FONT></TD>" & _
            "<TD WIDTH=""43%"" ALIGN=""CENTER"">" & _
            "<p><font  FACE=""�з���"" size=""4""><b>�����X��|�@�x�_�`�|</b></font> <p>" & _
            "<FONT FACE=""�з���"" SIZE=5><B>$Type$</B></FONT></TD>"
    
    prnFormString$ = prnFormString$ & _
            "<TD WIDTH=""34%""  ALIGN=""RIGHT"" >" & _
            "<table cellspacing=0 border=1 cellpadding=1 width=""100%"" align=""RIGHT""><tr>" & _
            "<td width=""80%"" align=""left"" colspan=2><font face=""�з���"" size=3 color=""#000080"">�f�����G</font><font face=""�з���"" size=3>$ChartNo$</font></td>" & _
            "<td width=""20%"" align=""CENTER"" rowspan=3><font face=""�з���"" size=3>$Dr_from$<br>$Bed_No$</font></td>" & _
            "</tr><tr>" & _
            "<td width=80%"" align=""left"" colspan=2><font face=""�з���"" size=3 color=""#000080"">�m�@�W�G</font><font face=""�з���"" size=3>$Name$</font></td>" & _
            "</tr><tr>" & _
            "<td width=""60%"" align=""left""><font face=""�з���"" size=3 color=""#000080"">�͡@��G</font><font face=""�з���"" size=2>$BirthDay$</font></td>" & _
            "<td width=""20%"" align=""CENTER""><font face=""�з���"" size=3>$Sex$��</font></td>" & _
            "</tr></table></TD>" & _
            "</TR></TABLE>"
    
    prnFormString$ = prnFormString$ & _
            "<hr align=""CENTER"" width=""100%"" size=1>" & _
            "<table cellspacing=0 border=0 cellpadding=1 width=630 align=""CENTER""><font size=3><tr>" & _
            "<td align=""left"" width=""12%""><font face=""�з���"" color=""#000080""><b>�f����</b></font></td>" & _
            "<td align=""left"" width=""20%""><font face=""�з���"" >$ChartNo$</font></td>" & _
            "<td align=""left"" width=""12%""><font face=""�з���"" color=""#000080""><b>�m�@�W</b></font></td>" & _
            "<td align=""left"" width=""20%""><font face=""�з���"" size=3>$Name$</font></td>" & _
            "<td align=""left"" width=""12%""><font face=""�з���""  color=""#000080""><b>�͡@��</b></font></td>" & _
            "<td align=""left"" width=""24%""><font face=""�з���"" >$BirthDay$�@$Sex$��</font></td></tr>" & _
            "<tr>" & _
            "<td align=""left"" width=""12%""><font face=""�з���"" color=""#000080""><b>�ˬd��</b></font></td>" & _
            "<td align=""left"" width=""20%""><font face=""�з���"" >$Dr_on$</font></td>" & _
            "<td align=""left"" width=""12%""><font face=""�з���"" color=""#000080""><b>�ˬd��O</b></font></td>" & _
            "<td align=""left"" width=""20%""><font face=""�з���"" size=3>$Division_on$</font></td>" & _
            "<td align=""left"" width=""12%""><font face=""�з���""  color=""#000080""><b>�ˬd���</b></font></td>" & _
            "<td align=""left"" width=""24%""><font face=""�з���"" >$Date$-$Time$</font></td></tr>" & _
            "<tr>"
            
    prnFormString$ = prnFormString$ & _
            "<td align=""left"" width=""12%""><font face=""�з���"" color=""#000080""><b>�}����v</b></font></td>" & _
            "<td align=""left"" width=""20%""><font face=""�з���"" >$Dr_order$</font></td>" & _
            "<td align=""left"" width=""12%""><font face=""�з���"" color=""#000080""><b>Modality</b></font></td>" & _
            "<td align=""left"" width=""56%"" COLSPAN=3><font face=""�з���"">$Modality$</font></td></tr>" & _
            "<tr>" & _
            "<td align=""left""><font face=""�з���"" size=3 color=""#000080""><b>�ˬd�Ӷ�</b></font></td>" & _
            "<td align=""left"" colspan=""5""><font face=""�з���"" >$ExamDetail$</font><font face=""�з���""  color=""#000080""></font></td>" & _
            "<tr>" & _
            "<td align=""left""><font face=""�з���"" size=2 color=""#000080""><B>C.Imp</B></font></td>" & _
            "<td align=""left"" colspan=""5""><font face=""�з���"" size=""3"">$ClinicImpression$</font></td></font>" & _
            "</table>"
    
    prnFormString$ = prnFormString$ & "<hr align=""CENTER"" width=""100%"" size=1>$ReportZone$ "
        
    signZone$ = "<TABLE CELLSPACING=0 BORDER=0 ALIGN='CENTER'><TR>" & _
                "<TD WIDTH=20% ALIGN=left><FONT FACE=""�з���"" SIZE=3 COLOR=""#000080"">���i��v�G</font></TD>" & _
                "<TD WIDTH=80% ALIGN=left><FONT FACE=""�з���"" SIZE=3 COLOR=""#000080"">$DrSignIn$</font></TD></TR>" & _
                "</TABLE><HR ALIGN=""CENTER"" WIDTH=""100%"" SIZE=1>"
    prnFormString$ = prnFormString$ & signZone$
    '$DrSignIn$_________________________ $LicenseNo$

    prnFormString$ = prnFormString$ & "$ImageArea$"

End Sub
Sub setPrnForm()

    'Dim MyString, MyNumber
    
    'prnFormString$ = ""
    'Open "reportLayout.txt" For Input As #1   ' �}�ҿ�J�ɮסC
    'Do While Not EOF(1)   ' ����j�骽���ɧ�����C
    '   Input #1, MyString ', MyNumber   ' �N���Ū�J����ܼƤ��C
       'Debug.Print MyString, MyNumber   ' �N��Ʀb�u�ߧY�v��������ܡC
    '   prnFormString$ = prnFormString$ & MyString
    'Loop
    'Close #1   ' �����ɮסC
    'Exit Sub
    
    '2007 01 05 �ק�, �ݳ]�w IE �C�L�榡�������� "&b &w page:&p/&P &b"  -----------------------------
    prnFormString$ = "<title>$ChartNo$ $Name$ ( $Type$ $Date$ ) </title>" & _
                     "<meta http-equiv=""Content-Type"" content=""text/html; charset=big5"" />" & _
                     "<meta http-equiv=""Content-Language"" content=""big5"" />"
    '------------------------------------------------------------------------------------------------
    
    prnFormString$ = prnFormString$ & "<BODY LINK=""#0000ff"" VLINK=""#800080"">"
    
'    prnFormString$ = prnFormString$ & _
            "<TABLE CELLSPACING=0 BORDER=0 CELLPADDING=1 WIDTH=""100%"" ALIGN=""CENTER""><TR>" & _
            "<TD WIDTH=""23%"" VALIGN=""TOP"" ROWSPAN=3><FONT FACE=""�з���"" SIZE=1>$Title$</FONT></TD>" & _
            "<TD WIDTH=""43%"" ALIGN=""CENTER"">" & _
            "<p><font  FACE=""�з���"" size=""4""><b>�����X��|�@������|</b></font> <p>" & _
            "<FONT FACE=""�з���"" SIZE=5><B>$Type$�ˬd���i</B></FONT></TD>"
    
    prnFormString$ = prnFormString$ & _
            "<TABLE CELLSPACING=0 BORDER=0 CELLPADDING=1 WIDTH=""100%"" ALIGN=""CENTER""><TR>" & _
            "<TD WIDTH=""23%"" VALIGN=""TOP"" ROWSPAN=3><FONT FACE=""�з���"" SIZE=1></FONT></TD>" & _
            "<TD WIDTH=""43%"" ALIGN=""CENTER"">" & _
            "<p><font  FACE=""�з���"" size=""4""><b>$Site$</b></font> <p>" & _
            "<FONT FACE=""�з���"" SIZE=5><B>$Type$</B></FONT></TD>"
    
    prnFormString$ = prnFormString$ & _
            "<TD WIDTH=""34%""  ALIGN=""RIGHT"" >" & _
            "<table cellspacing=0 border=1 cellpadding=1 width=""100%"" align=""RIGHT""><tr>" & _
            "<td width=""80%"" align=""left"" colspan=2><font face=""�з���"" size=3 color=""#000080"">�f�����G</font><font face=""�з���"" size=3>$ChartNo$</font></td>" & _
            "<td width=""20%"" align=""CENTER"" rowspan=3><font face=""�з���"" size=3>$Dr_from$<br>$Bed_No$</font></td>" & _
            "</tr><tr>" & _
            "<td width=80%"" align=""left"" colspan=2><font face=""�з���"" size=3 color=""#000080"">�m�@�W�G</font><font face=""�з���"" size=3>$Name$</font></td>" & _
            "</tr><tr>" & _
            "<td width=""60%"" align=""left""><font face=""�з���"" size=3 color=""#000080"">�͡@��G</font><font face=""�з���"" size=2>$BirthDay$</font></td>" & _
            "<td width=""20%"" align=""CENTER""><font face=""�з���"" size=3>$Sex$��</font></td>" & _
            "</tr></table></TD>" & _
            "</TR></TABLE>"
    
    prnFormString$ = prnFormString$ & _
            "<hr align=""CENTER"" width=""100%"" size=1>" & _
            "<table cellspacing=0 border=0 cellpadding=1 width=630 align=""CENTER""><font size=3><tr>" & _
            "<td align=""left"" width=""12%""><font face=""�з���"" color=""#000080""><b>�ˬd��</b></font></td>" & _
            "<td align=""left"" width=""20%""><font face=""�з���"" >$Dr_on$</font></td>" & _
            "<td align=""left"" width=""12%""><font face=""�з���"" color=""#000080""><b>�ˬd��O</b></font></td>" & _
            "<td align=""left"" width=""20%""><font face=""�з���"" size=3>$Division_on$</font></td>" & _
            "<td align=""left"" width=""12%""><font face=""�з���""  color=""#000080""><b>�ˬd���</b></font></td>" & _
            "<td align=""left"" width=""24%""><font face=""�з���"" >$Date$-$Time$</font></td></tr>" & _
            "<tr>" & _
            "<td align=""left"" width=""12%""><font face=""�з���"" color=""#000080""><b>�}����v</b></font></td>" & _
            "<td align=""left"" width=""20%""><font face=""�з���"" >$Dr_order$</font></td>" & _
            "<td align=""left"" width=""12%""><font face=""�з���"" color=""#000080""><b>Modality</b></font></td>" & _
            "<td align=""left"" width=""56%"" COLSPAN=3><font face=""�з���"">$Modality$</font></td></tr>" & _
            "<tr>" & _
            "<td align=""left""><font face=""�з���"" size=3 color=""#000080""><b>�ˬd�Ӷ�</b></font></td>" & _
            "<td align=""left"" colspan=""5""><font face=""�з���"" >$ExamDetail$</font><font face=""�з���""  color=""#000080""></font></td>" & _
            "</table>"

'2006/11/15 ����P�`�|���˭n�D���C�L c.c. IMP ��� ---------------------------------------
'            "<tr>" & _
            "<td align=""left""><font face=""�з���"" size=2 color=""#000080""><B>C.Imp</B></font></td>" & _
            "<td align=""left"" colspan=""5""><font face=""�з���"" size=""3"">$ClinicImpression$</font></td></font>" & _
            "</table>"
'-----------------------------------------------------------------------------------------

    prnFormString$ = prnFormString$ & "<hr align=""CENTER"" width=""100%"" size=1>$ReportZone$ "
        
'    signZone$ = "<TABLE CELLSPACING=0 BORDER=0 ALIGN='CENTER'>" & _
                "<TR><TD WIDTH=100% ALIGN=CENTER><FONT FACE=""�з���"" SIZE=3 COLOR=""#000080"">���i��v�G$DrSignIn$_________________________<BR>$LicenseNo$</font></TD></TR>" & _
                "</TABLE><HR ALIGN=""CENTER"" WIDTH=""100%"" SIZE=1>"
'    prnFormString$ = prnFormString$ & signZone$
    signZone$ = "<TABLE CELLSPACING=0 BORDER=0 ALIGN='CENTER'><TR>" & _
                "<TD WIDTH=12% ALIGN=left VALIGN=TOP><FONT FACE=""�з���"" SIZE=3 color=""#0000180""><B>���i��v</B></font></TD>" & _
                "<TD WIDTH=88% ALIGN=left><FONT FACE=""�з���"" SIZE=3 COLOR=""#000080"">$DrSignIn$</font></TD></TR>" & _
                "</TABLE><HR ALIGN=""CENTER"" WIDTH=""100%"" SIZE=1>"
    prnFormString$ = prnFormString$ & signZone$
    '$DrSignIn$_________________________ $LicenseNo$
    
    prnFormString$ = prnFormString$ & "$ImageArea$"
    'Open App.Path & "\PrintLayout.txt" For Output As #6
    'Print #6, prnFormString$
    'Close #6
    
End Sub
Sub setPrnFormfromDB(xType$)
    Dim adoDB As New adoDB.Connection
    Dim adoTable As New adoDB.Recordset
    Dim SQL$, i%, sCounter&, sAbbreviation$
    
    '/**/
    'SQL$ = "SELECT ALL * FROM CRIS_PrintLayout WHERE ReportName='" & xType$ & "'"
    '/**/
    SQL$ = "SELECT ALL * FROM CRIS_PrintLayout  WHERE ReportName='" & xType$ & "'"
    '/**/
    adoDB.Open dbConnection$
    adoTable.Open SQL$, adoDB, adOpenDynamic, adLockPessimistic
    
    DoEvents
    
    If Not adoTable.EOF Then
       prnFormString$ = adoTable!Layout
       prnFormImageCols% = adoTable!ImageCols
    Else
       prnFormString$ = ""
       prnFormImageCols% = 0
    End If
    adoTable.Close
    
    If prnFormString$ = "" Then
        '/**/
        'SQL$ = "SELECT ALL * FROM CRIS_PrintLayout WHERE ReportName='�w�]'"
        '/**/
        SQL$ = "SELECT ALL * FROM CRIS_PrintLayout WHERE ReportName='�w�]'"
        '/**/
        adoTable.Open SQL$, adoDB, adOpenDynamic, adLockPessimistic
        If Not adoTable.EOF Then
           prnFormString$ = adoTable!Layout
           prnFormImageCols% = adoTable!ImageCols
        Else
           prnFormString$ = ""
           prnFormImageCols% = 0
        End If
        adoTable.Close
    
    End If
    
    adoDB.Close
    Set adoTable = Nothing
    Set adoDB = Nothing
    
End Sub

Function getPageHeadfromDB(xType$) As String
    Dim adoDB As New adoDB.Connection
    Dim adoTable As New adoDB.Recordset
    Dim SQL$, i%, sCounter&, sAbbreviation$
    
    '/**/
    'SQL$ = "SELECT ALL * FROM CRIS_PrintLayout WHERE ReportName='" & xType$ & "'"
    '/**/
    SQL$ = "SELECT ALL * FROM CRIS_PrintLayout WHERE ReportName='" & xType$ & "'"
    '/**/
    adoDB.Open dbConnection$
    adoTable.Open SQL$, adoDB, adOpenForwardOnly, adLockReadOnly
    If Not adoTable.EOF Then
       getPageHeadfromDB = NoNull(adoTable!Layout)
    Else
       getPageHeadfromDB = ""
    End If
    
    adoTable.Close
    adoDB.Close
    Set adoTable = Nothing
    Set adoDB = Nothing
    
End Function
Function getDeptDoctor(drid$) As String

    Dim adoDB As adoDB.Connection
    Dim adoTable As adoDB.Recordset
    Dim conn$, SQL$
    
    Set adoDB = New adoDB.Connection
    Set adoTable = New adoDB.Recordset
    
    adoDB.Open dbConnection$
    
    '/**/
    'SQL$ = "SELECT Class,Type,Code,Remark FROM CRIS_Reference WHERE Class='Doctor' AND Code='" & drid$ & "'"
    '/**/
    SQL$ = "SELECT Class,Type,Code,Remark FROM CRIS_Reference WHERE Class='Doctor' AND Code='" & drid$ & "'"
    '/**/
    adoTable.Open SQL$, adoDB, adOpenForwardOnly, adLockReadOnly 'tblName$
    
    If Not adoTable.EOF Then
       'Field_get = NoNull(adoTable(fieldName$))
       getDeptDoctor = UCase(drid$) & NoNull(adoTable!Remark)
    Else
       getDeptDoctor = drid$
    End If
    
    adoTable.Close
    adoDB.Close
    
    Set adoTable = Nothing
    Set adoDB = Nothing

End Function

Function getLastClass(xChartNo$, xUni_key$) As String
    Dim SQL$, i%
    Dim adoDB As New adoDB.Connection
    Dim adoTable As New adoDB.Recordset
    Dim conn$ ', SQL$
    
    '/**/
    'SQL$ = "SELECT Class FROM CRIS_Exam_Online WHERE ChartNo='" & xChartNo$ & "' AND Uni_key='" & xUni_key$ & "'"
    '/**/
    SQL$ = "SELECT Class FROM CRIS_Exam_Online WHERE status<>'�w�R��' and ChartNo='" & xChartNo$ & "' AND Uni_key='" & xUni_key$ & "'"
    '/**/
    
    adoDB.Open dbConnection$
    adoTable.Open SQL$, adoDB
    If Not adoTable.EOF Then
       getLastClass = NoNull(adoTable!Class)
    Else
       getLastClass = ""
    End If
    
    adoTable.Close
    adoDB.Close
    Set adoTable = Nothing
    Set adoDB = Nothing
    
End Function

Function ini_Purge(inString$, rtnNumber&) As String
    Dim i%
    
    If rtnNumber& = 0 Or Len(inString$) < 1 Or Len(inString$) > rtnNumber& Then
        ini_Purge = ""
    Else
        For i% = 1 To Len(inString$)
            If Asc(Mid(inString$, i%, 1)) = 0 Then Exit For
        Next
        i% = i% - 1
        ini_Purge = Left(inString$, i%)
    End If
    
End Function
Public Function ChKChartno(ss As String)
        Dim aa As Integer
        Dim cno(1 To 10) As Integer
        
        cno(1) = 4
        cno(2) = 3
        cno(3) = 2
        cno(4) = 9
        cno(5) = 8
        cno(6) = 7
        cno(7) = 4
        cno(8) = 3
        cno(9) = 2
        ss = Format(ss, "0000000000")
        aa = 0
        For i = 1 To 9
            aa = aa + Val(Mid(ss, i, 1)) * cno(i)
        Next i
        BB = aa Mod 10
        If (9 - BB) = Val(Mid(ss, 10, 1)) Then
            ChKChartno = True
        Else
            ChKChartno = False
        End If
        
End Function

Sub getLastUpdate()
    Dim SQL$, i%
    Dim adoDB As New adoDB.Connection
    Dim adoTable As New adoDB.Recordset
    Dim conn$ ', SQL$
    
    '/**/
    'SQL$ = "SELECT lastUpdateDate, lastUpdateTime " & _
           "FROM CRIS_Exam_Online " & _
           "WHERE ChartNo='" & curr_Record.ChartNo & "' AND " & _
           "Uni_key='" & curr_Record.Uni_key & "' "
    '/**/
    SQL$ = "SELECT lastUpdateDate, lastUpdateTime " & _
           "FROM CRIS_Exam_Online " & _
           "WHERE status<>'�w�R��' and ChartNo='" & curr_Record.chartno & "' AND " & _
           "Uni_key='" & curr_Record.uni_key & "' "
    '/**/
    adoDB.Open dbConnection$
    adoTable.Open SQL$, adoDB
    DoEvents
    
    If Not adoTable.EOF Then
        xLastUpdateDate$ = NoNull(adoTable!LastUpdateDate)
        xLastUpdateTime$ = NoNull(adoTable!LastUpdateTime)
    End If
        
    adoTable.Close
    adoDB.Close
    Set adoTable = Nothing
    Set adoDB = Nothing

End Sub

Public Function prepStringForSQL(ByVal sValue As String) As String
    Dim sAns As String
    
    sAns = Replace(sValue, Chr(39), "''")
    sAns = "'" & sAns & "'"
    prepStringForSQL = sAns

End Function
Sub chkExamDetail(lstExamDetail As ListBox, xExamDetail$)
    Dim tmp$, xtmp$, i%
    'Dim xExamCode(lstExamDetail.ListCount) As String
    
    tmp$ = xExamDetail$
    Do While Len(tmp$) > 0
       If InStr(tmp$, ",") > 0 Then
          xtmp$ = Trim(Left(tmp$, InStr(tmp$, ",") - 1))
          tmp$ = Right(tmp$, Len(tmp$) - InStr(tmp$, ","))
       Else
          xtmp$ = Trim(tmp$)
          tmp$ = ""
       End If
       
       For i% = 0 To lstExamDetail.ListCount - 1
'           If xTmp$ = Trim(left(lstExamDetail.List(i%), 10)) Then
           If xtmp$ = Trim(Left(lstExamDetail.List(i%), InStr(lstExamDetail.List(i%), ":") - 1)) Then
              lstExamDetail.Selected(i%) = True
           End If
       Next
          
    Loop

End Sub
Sub setOPLog(xUserID$, xOPType$, xOPLog$)
        Dim adoDB As New adoDB.Connection
        Dim conn$, SQL$, xUploadCode$, tmpHISup$, timStamp$
        
        On Error GoTo insertOPLogErr
        
        timStamp$ = str(Timer)
        
        adoDB.Open dbConnection$
        SQL$ = "INSERT INTO CRIS_OperationLog (UserID, OPType, OPLog, LogDate, LogTime) " & _
               "VALUES ('" & _
               xUserID$ & "','" & xOPType$ & "', " & prepStringForSQL(xOPLog$) & ", '" & Format(Date, "yyMMdd") & "', '" & _
               timStamp$ & "')"
        adoDB.Execute SQL$
        
        adoDB.Close
        Set adoDB = Nothing
        On Error GoTo 0
        
        Exit Sub
        
insertOPLogErr:
        timStamp$ = str(Val(timStamp$) + 1)
        SQL$ = "INSERT INTO CRIS_OperationLog (UserID, OPType, OPLog, LogDate, LogTime) " & _
               "VALUES ('" & _
               xUserID$ & "','" & xOPType$ & "', " & prepStringForSQL(xOPLog$) & ", '" & Format(Date, "yyMMdd") & "', '" & _
               timStamp$ & "')"
        Resume

End Sub
'Function ReadINI_DB2(ByVal iType As Integer) As String
'    Dim rtn As Long
'    Dim rtnString As String * 260, strTemp As String
'    Dim sResult As String
'    Dim sAppFile As String
'
'    sResult = ""
'    sAppFile = App.Path & "\ExamSVR.ini"
'
'    Select Case iType
'    Case 0 '�O�_���DB2�����f�w���, 0-NO, 1-YES
'         rtn = ReadINI("DB2", "DB2", "0", rtnString, Len(rtnString), sAppFile)
'         If Not rtn = 0 Then
'           sResult = Left(rtnString, rtn)
'         Else
'           sResult = "0"
'         End If
'    Case 1 '���oDB2���s���r��
'         sResult = "PROVIDER=MSDASQL;dsn=CGHHQDB;uid=minipacs;pwd=minipacs;"
'         rtn = ReadINI("DB2", "CONNECTIONSTRING", sResult, rtnString, Len(rtnString), sAppFile)
'         If Not rtn = 0 Then sResult = Left(rtnString, rtn)
'    End Select
'    ReadINI_DB2 = sResult
'
'End Function
