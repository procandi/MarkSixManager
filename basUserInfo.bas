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
Public Const COMMANDFROM_INI As String = "./CommandFrom.ini" '�O���@��CommandFrom�����O�����
Public Const CHECKPRINTERSTATUS_INI As String = "./CheckPrinterStatus.ini" '�O���@��CheckPrinterStatus�����A�O�������
Public Const SYSTEMSYNC_INI As String = "./SystemSync.ini" '�O���@��SystemSync�һݪ��P�B��Ƨ������|�����
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


'/*�O����^���S���q�q�l�f���{���ɷ|�Ψ쪺����m�`��*/
Public Const MWL_MWL_KEY As Integer = 0
Public Const MWL_TRIGGER_DTTM As Integer = 1
Public Const MWL_REPLICA_DTTM As Integer = 2
Public Const MWL_CHARACTER_SET As Integer = 3
Public Const MWL_SCHEDULED_AETITLE As Integer = 4
Public Const MWL_SCHEDULED_DTTM As Integer = 5
Public Const MWL_SCHEDULED_MODALITY As Integer = 6
Public Const MWL_SCHEDULED_STATION As Integer = 7
Public Const MWL_SCHEDULED_LOCATION As Integer = 8
Public Const MWL_SCHEDULED_PROC_ID As Integer = 9
Public Const MWL_SCHEDULED_PROC_DESC As Integer = 10
Public Const MWL_SCHEDULED_ACTION_CODES As Integer = 11
Public Const MWL_SCHEDULED_PROC_STATUS As Integer = 12
Public Const MWL_PREMEDICATION As Integer = 13
Public Const MWL_CONTRAST_AGENT As Integer = 14
Public Const MWL_REQUESTED_PROC_ID As Integer = 15
Public Const MWL_REQUESTED_PROC_DESC As Integer = 16
Public Const MWL_REQUESTED_PROC_CODES As Integer = 17
Public Const MWL_REQUESTED_PROC_PRIORITY As Integer = 18
Public Const MWL_REQUESTED_PROC_REASON As Integer = 19
Public Const MWL_REQUESTED_PROC_COMMENTS As Integer = 20
Public Const MWL_STUDY_INSTANCE_UID As Integer = 21
Public Const MWL_PROC_PLACER_ORDER_NO As Integer = 22
Public Const MWL_PROC_FILLER_ORDER_NO As Integer = 23
Public Const MWL_ACCESSION_NO As Integer = 24
Public Const MWL_ATTEND_DOCTOR As Integer = 25
Public Const MWL_PERFORM_DOCTOR As Integer = 26
Public Const MWL_CONSULT_DOCTOR As Integer = 27
Public Const MWL_REQUEST_DOCTOR As Integer = 28
Public Const MWL_REFER_DOCTOR As Integer = 29
Public Const MWL_REQUEST_DEPARTMENT As Integer = 30
Public Const MWL_IMAGING_REQUEST_REASON As Integer = 31
Public Const MWL_IMAGING_REQUEST_COMMENTS As Integer = 32
Public Const MWL_IMAGING_REQUEST_DTTM As Integer = 33
Public Const MWL_ISR_PLACER_ORDER_NO As Integer = 34
Public Const MWL_ISR_FILLER_ORDER_NO As Integer = 35
Public Const MWL_ADMISSION_ID As Integer = 36
Public Const MWL_PATIENT_TRANSPORT As Integer = 37
Public Const MWL_PATIENT_LOCATION As Integer = 38
Public Const MWL_PATIENT_RESIDENCY As Integer = 39
Public Const MWL_PATIENT_NAME As Integer = 40
Public Const MWL_PATIENT_ID As Integer = 41
Public Const MWL_OTHER_PATIENT_NAME As Integer = 42
Public Const MWL_OTHER_PATIENT_ID As Integer = 43
Public Const MWL_PATIENT_BIRTH_DATE As Integer = 44
Public Const MWL_PATIENT_SEX As Integer = 45
Public Const MWL_PATIENT_WEIGHT As Integer = 46
Public Const MWL_PATIENT_SIZE As Integer = 47
Public Const MWL_PATIENT_STATE As Integer = 48
Public Const MWL_CONFIDENTIALITY As Integer = 49
Public Const MWL_PREGNANCY_STATUS As Integer = 50
Public Const MWL_MEDICAL_ALERTS As Integer = 51
Public Const MWL_CONTRAST_ALLERGIES As Integer = 52
Public Const MWL_SPECIAL_NEEDS As Integer = 53
Public Const MWL_SPECIALTY As Integer = 54
Public Const MWL_DIAGNOSIS As Integer = 55
Public Const MWL_ADMIT_DTTM As Integer = 56
Public Const MWL_REGISTER_DTTM As Integer = 57
Public Const MWL_FIELD_COUNT As Integer = 58

Public Const MWLWL_MWL_KEY As Integer = 0
Public Const MWLWL_TRIGGER_DTTM As Integer = 1
Public Const MWLWL_REPLICA_DTTM As Integer = 2
Public Const MWLWL_EVENT_TYPE As Integer = 3
Public Const MWLWL_CHARACTER_SET As Integer = 4
Public Const MWLWL_SCHEDULED_AETITLE As Integer = 5
Public Const MWLWL_SCHEDULED_DTTM As Integer = 6
Public Const MWLWL_SCHEDULED_MODALITY As Integer = 7
Public Const MWLWL_SCHEDULED_STATION As Integer = 8
Public Const MWLWL_SCHEDULED_LOCATION As Integer = 9
Public Const MWLWL_SCHEDULED_ACTION_CODES As Integer = 10
Public Const MWLWL_SCHEDULED_PROC_STATUS As Integer = 11
Public Const MWLWL_PREMEDICATION As Integer = 12
Public Const MWLWL_CONTRAST_AGENT As Integer = 13
Public Const MWLWL_REQUESTED_PROC_PRIORITY As Integer = 14
Public Const MWLWL_REQUESTED_PROC_REASON As Integer = 15
Public Const MWLWL_REQUESTED_PROC_COMMENTS As Integer = 16
Public Const MWLWL_STUDY_INSTANCE_UID As Integer = 17
Public Const MWLWL_PROC_PLACER_ORDER_NO As Integer = 18
Public Const MWLWL_PROC_FILLER_ORDER_NO As Integer = 19
Public Const MWLWL_ACCESSION_NO As Integer = 20
Public Const MWLWL_ATTEND_DOCTOR As Integer = 21
Public Const MWLWL_PERFORM_DOCTOR As Integer = 22
Public Const MWLWL_CONSULT_DOCTOR As Integer = 23
Public Const MWLWL_REQUEST_DOCTOR As Integer = 24
Public Const MWLWL_REFER_DOCTOR As Integer = 25
Public Const MWLWL_REQUEST_DEPARTMENT As Integer = 26
Public Const MWLWL_IMAGING_REQUEST_REASON As Integer = 27
Public Const MWLWL_IMAGING_REQUEST_COMMENTS As Integer = 28
Public Const MWLWL_IMAGING_REQUEST_DTTM As Integer = 29
Public Const MWLWL_ISR_PLACER_ORDER_NO As Integer = 30
Public Const MWLWL_ISR_FILLER_ORDER_NO As Integer = 31
Public Const MWLWL_ADMISSION_ID As Integer = 32
Public Const MWLWL_PATIENT_TRANSPORT As Integer = 33
Public Const MWLWL_PATIENT_LOCATION As Integer = 34
Public Const MWLWL_PATIENT_RESIDENCY As Integer = 35
Public Const MWLWL_PATIENT_NAME As Integer = 36
Public Const MWLWL_PATIENT_ID As Integer = 37
Public Const MWLWL_OTHER_PATIENT_NAME As Integer = 38
Public Const MWLWL_OTHER_PATIENT_ID As Integer = 39
Public Const MWLWL_PATIENT_BIRTH_DATE As Integer = 40
Public Const MWLWL_PATIENT_SEX As Integer = 41
Public Const MWLWL_PATIENT_WEIGHT As Integer = 42
Public Const MWLWL_PATIENT_SIZE As Integer = 43
Public Const MWLWL_PATIENT_STATE As Integer = 44
Public Const MWLWL_CONFIDENTIALITY As Integer = 45
Public Const MWLWL_PREGNANCY_STATUS As Integer = 46
Public Const MWLWL_MEDICAL_ALERTS As Integer = 47
Public Const MWLWL_CONTRAST_ALLERGIES As Integer = 48
Public Const MWLWL_SPECIAL_NEEDS As Integer = 49
Public Const MWLWL_SPECIALTY As Integer = 50
Public Const MWLWL_DIAGNOSIS As Integer = 51
Public Const MWLWL_ADMIT_DTTM As Integer = 52
Public Const MWLWL_REGISTER_DTTM As Integer = 53
Public Const MWLWL_PATIENT_ID_ISSUER As Integer = 54
Public Const MWLWL_OTHER_PATIENT_ID_ISSUER As Integer = 55
Public Const MWLWL_VALIDATE_DTTM As Integer = 56
Public Const MWLWL_ORDER_COUNT As Integer = 57
Public Const MWLWL_SCHEDULED_PROC_ID1 As Integer = 58
Public Const MWLWL_SCHEDULED_PROC_DESC1 As Integer = 59
Public Const MWLWL_SCHEDULED_PROC_ID2 As Integer = 60
Public Const MWLWL_SCHEDULED_PROC_DESC2 As Integer = 61
Public Const MWLWL_SCHEDULED_PROC_ID3 As Integer = 62
Public Const MWLWL_SCHEDULED_PROC_DESC3 As Integer = 63
Public Const MWLWL_SCHEDULED_PROC_ID4 As Integer = 64
Public Const MWLWL_SCHEDULED_PROC_DESC4 As Integer = 65
Public Const MWLWL_SCHEDULED_PROC_ID5 As Integer = 66
Public Const MWLWL_SCHEDULED_PROC_DESC5 As Integer = 67
Public Const MWLWL_SCHEDULED_PROC_ID6 As Integer = 68
Public Const MWLWL_SCHEDULED_PROC_DESC6 As Integer = 69
Public Const MWLWL_SCHEDULED_PROC_ID7 As Integer = 70
Public Const MWLWL_SCHEDULED_PROC_DESC7 As Integer = 71
Public Const MWLWL_SCHEDULED_PROC_ID8 As Integer = 72
Public Const MWLWL_SCHEDULED_PROC_DESC8 As Integer = 73
Public Const MWLWL_SCHEDULED_PROC_ID9 As Integer = 74
Public Const MWLWL_SCHEDULED_PROC_DESC9 As Integer = 75
Public Const MWLWL_SCHEDULED_PROC_ID10 As Integer = 76
Public Const MWLWL_SCHEDULED_PROC_DESC10 As Integer = 77
Public Const MWLWL_SCHEDULED_PROC_ID11 As Integer = 78
Public Const MWLWL_SCHEDULED_PROC_DESC11 As Integer = 79
Public Const MWLWL_SCHEDULED_PROC_ID12 As Integer = 80
Public Const MWLWL_SCHEDULED_PROC_DESC12 As Integer = 81
Public Const MWLWL_SCHEDULED_PROC_ID13 As Integer = 82
Public Const MWLWL_SCHEDULED_PROC_DESC13 As Integer = 83
Public Const MWLWL_SCHEDULED_PROC_ID14 As Integer = 84
Public Const MWLWL_SCHEDULED_PROC_DESC14 As Integer = 85
Public Const MWLWL_SCHEDULED_PROC_ID15 As Integer = 86
Public Const MWLWL_SCHEDULED_PROC_DESC15 As Integer = 87
Public Const MWLWL_SCHEDULED_PROC_ID16 As Integer = 88
Public Const MWLWL_SCHEDULED_PROC_DESC16 As Integer = 89
Public Const MWLWL_SCHEDULED_PROC_ID17 As Integer = 90
Public Const MWLWL_SCHEDULED_PROC_DESC17 As Integer = 91
Public Const MWLWL_SCHEDULED_PROC_ID18 As Integer = 92
Public Const MWLWL_SCHEDULED_PROC_DESC18 As Integer = 93
Public Const MWLWL_SCHEDULED_PROC_ID19 As Integer = 94
Public Const MWLWL_SCHEDULED_PROC_DESC19 As Integer = 95
Public Const MWLWL_SCHEDULED_PROC_ID20 As Integer = 96
Public Const MWLWL_SCHEDULED_PROC_DESC20 As Integer = 97
Public Const MWLWL_FIELD_COUNT As Integer = 98



Public Const REPORTWL_REPORTWL_KEY As Integer = 0
Public Const REPORTWL_TRIGGER_DTTM As Integer = 1
Public Const REPORTWL_REPLICA_DTTM As Integer = 2
Public Const REPORTWL_EXAM_ID As Integer = 3
Public Const REPORTWL_PATIENT_ID As Integer = 4
Public Const REPORTWL_REPORT_STAT As Integer = 5
Public Const REPORTWL_CREATOR_ID As Integer = 6
Public Const REPORTWL_CREATOR_NAME As Integer = 7
Public Const REPORTWL_CREATE_DTTM As Integer = 8
Public Const REPORTWL_DICTATOR_ID As Integer = 9
Public Const REPORTWL_DICTATOR_NAME As Integer = 10
Public Const REPORTWL_DICTATE_DTTM As Integer = 11
Public Const REPORTWL_TRANSCRIBER_ID As Integer = 12
Public Const REPORTWL_TRANSCRIBER_NAME As Integer = 13
Public Const REPORTWL_TRANSCRIBE_DTTM As Integer = 14
Public Const REPORTWL_APPROVER_ID As Integer = 15
Public Const REPORTWL_APPROVER_NAME As Integer = 16
Public Const REPORTWL_APPROVE_DTTM As Integer = 17
Public Const REPORTWL_REVISER_ID As Integer = 18
Public Const REPORTWL_REVISER_NAME As Integer = 19
Public Const REPORTWL_REVISE_DTTM As Integer = 20
Public Const REPORTWL_REPORT_TYPE As Integer = 21
Public Const REPORTWL_REPORT_TEXT As Integer = 22
Public Const REPORTWL_CONCLUSION As Integer = 23
Public Const REPORTWL_FIELD_COUNT As Integer = 24



Public Const EXAM_ONLINE_SYSTEM As Integer = 0
Public Const EXAM_ONLINE_UNI_KEY As Integer = 1
Public Const EXAM_ONLINE_CHARTNO As Integer = 2
Public Const EXAM_ONLINE_EXAMDATE As Integer = 3
Public Const EXAM_ONLINE_EXAMTIME As Integer = 4
Public Const EXAM_ONLINE_TYPE As Integer = 5
Public Const EXAM_ONLINE_ROOM As Integer = 6
Public Const EXAM_ONLINE_AGE As Integer = 7
Public Const EXAM_ONLINE_ITEM1 As Integer = 8
Public Const EXAM_ONLINE_ITEM2 As Integer = 9
Public Const EXAM_ONLINE_ITEM3 As Integer = 10
Public Const EXAM_ONLINE_ITEM4 As Integer = 11
Public Const EXAM_ONLINE_ITEM5 As Integer = 12
Public Const EXAM_ONLINE_ITEM6 As Integer = 13
Public Const EXAM_ONLINE_OTHERS As Integer = 14
Public Const EXAM_ONLINE_DR_ON As Integer = 15
Public Const EXAM_ONLINE_STATUS As Integer = 16
Public Const EXAM_ONLINE_CLASS As Integer = 17
Public Const EXAM_ONLINE_IMGPICKED As Integer = 18
Public Const EXAM_ONLINE_MODALITY As Integer = 19
Public Const EXAM_ONLINE_REG_DATE As Integer = 20
Public Const EXAM_ONLINE_DR_FROM As Integer = 21
Public Const EXAM_ONLINE_EXAMDETAIL As Integer = 22
Public Const EXAM_ONLINE_ORDERDATE As Integer = 23
Public Const EXAM_ONLINE_ORDERTIME As Integer = 24
Public Const EXAM_ONLINE_REPORTDATE As Integer = 25
Public Const EXAM_ONLINE_REPORTTIME As Integer = 26
Public Const EXAM_ONLINE_ACCESSIONNUMBER As Integer = 27
Public Const EXAM_ONLINE_UPLOADCODE As Integer = 28
Public Const EXAM_ONLINE_LASTUPDATEDATE As Integer = 29
Public Const EXAM_ONLINE_LASTUPDATETIME As Integer = 30
Public Const EXAM_ONLINE_DIVISION_FROM As Integer = 31
Public Const EXAM_ONLINE_DIVISION_ON As Integer = 32
Public Const EXAM_ONLINE_DR_ORDER As Integer = 33
Public Const EXAM_ONLINE_DR_REPORT As Integer = 34
Public Const EXAM_ONLINE_DIVISION_SEQ As Integer = 35
Public Const EXAM_ONLINE_CLINICALIMP As Integer = 36
Public Const EXAM_ONLINE_HISUP As Integer = 37
Public Const EXAM_ONLINE_HIS_REQNO As Integer = 38
Public Const EXAM_ONLINE_HIS_ACCNO As Integer = 39
Public Const EXAM_ONLINE_TEMPLATENAME As Integer = 40
Public Const EXAM_ONLINE_CHARGEBY As Integer = 41
Public Const EXAM_ONLINE_NOCHECKINCODE As Integer = 42
Public Const EXAM_ONLINE_NOCHECKINTEXT As Integer = 43
Public Const EXAM_ONLINE_TRACK As Integer = 44
Public Const EXAM_ONLINE_MEMO As Integer = 45
Public Const EXAM_ONLINE_TEMPLATEFILE As Integer = 46
Public Const EXAM_ONLINE_CHECKINDATE As Integer = 47
Public Const EXAM_ONLINE_CHECKINTIME As Integer = 48
Public Const EXAM_ONLINE_SPREADID As Integer = 49
Public Const EXAM_ONLINE_DIAGNOSIS As Integer = 50
Public Const EXAM_ONLINE_CONTENT As Integer = 51
Public Const EXAM_ONLINE_FIELD_COUNT As Integer = 52


Public Const PATIENT_INFO_CHARTNO As Integer = 52
Public Const PATIENT_INFO_NAME As Integer = 53
Public Const PATIENT_INFO_CITIZENID As Integer = 54
Public Const PATIENT_INFO_BIRTHDAY As Integer = 55
Public Const PATIENT_INFO_SEX As Integer = 56
Public Const PATIENT_INFO_FIELD_COUNT As Integer = 57


Public Const EREP_ID As Integer = 0
Public Const EREP_REFERNO As Integer = 1
Public Const EREP_PATNAME As Integer = 2
Public Const EREP_PATIDTYPE As Integer = 3
Public Const EREP_PATIDNUMBER As Integer = 4
Public Const EREP_PATBIRTH As Integer = 5
Public Const EREP_PATAGE As Integer = 6
Public Const EREP_PATNO As Integer = 7
Public Const EREP_SOURCE As Integer = 8
Public Const EREP_ROOM As Integer = 9
Public Const EREP_MEMO As Integer = 10
Public Const EREP_ORDERDESC As Integer = 11
Public Const EREP_ORDERDATE As Integer = 12
Public Const EREP_EXAMDATE As Integer = 13
Public Const EREP_REPORTDATE As Integer = 14
Public Const EREP_DEP As Integer = 15
Public Const EREP_ICD As Integer = 16
Public Const EREP_EXAMRESULT As Integer = 17
Public Const EREP_DESCRIPT As Integer = 18
Public Const EREP_PICTURE As Integer = 19
Public Const EREP_HOSPNAME As Integer = 20
Public Const EREP_DOCTORUSERID As Integer = 21
Public Const EREP_OPUSERID As Integer = 22
Public Const EREP_LOCATION As Integer = 23
Public Const EREP_INFSOURCE As Integer = 24
Public Const EREP_ORDNO As Integer = 25
Public Const EREP_SAMPLING As Integer = 26
Public Const EREP_EXAMRESULT2 As Integer = 27
Public Const EREP_CREATEDATE As Integer = 28
Public Const EREP_SEX As Integer = 29
Public Const EREP_EQM As Integer = 30
Public Const EREP_TEL As Integer = 31
Public Const EREP_PACSREFERNO As Integer = 32
Public Const EREP_EXAMDEP As Integer = 33
Public Const EREP_DOCTORUSERDEP As Integer = 34
Public Const EREP_MAJORDOCTORUSERID As Integer = 35
Public Const EREP_MEDICALHISTORY As Integer = 36
Public Const EREP_IMAGECOUNT As Integer = 37
Public Const EREP_WORKLISTUID As Integer = 38
Public Const EREP_EXAMINEDATE As Integer = 39
Public Const EREP_ICD9 As Integer = 40
Public Const EREP_CHARGECODE As Integer = 41
Public Const EREP_SECRETCODE As Integer = 42
Public Const EREP_IMPRESSION As Integer = 43
Public Const EREP_FIELD_COUNT As Integer = 44
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
    Dim I As Long
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
    I = 0
    For Each obj In frm.Controls
        If VarType(obj) <> VARTYPE_COMMONDIALOG And VarType(obj) <> VARTYPE_COMMAND And VarType(obj) <> VARTYPE_PICTUREBOX Then
            SortObj(I).ID = I
            SortObj(I).Name = obj.Name
            Select Case UCase(Left(obj.Name, 3))
            Case "TXT", "TEX"
                SortObj(I).Text = obj.Text
            Case "LBL", "LAB"
                SortObj(I).Text = obj.Caption
            Case "OPT", "CHE"
                If obj.Value Then
                    SortObj(I).Text = "��" & obj.Caption
                Else
                    SortObj(I).Text = "��" & obj.Caption
                End If
            End Select
            


           SortObj(I).Left = obj.Left
           SortObj(I).Top = obj.Top
           If UCase(Left(obj.Name, 3)) = "TXT" Or UCase(Left(obj.Name, 3)) = "LAB" Then
                If obj.LinkItem <> "" Then
                    temp = Split(obj.LinkItem, "_")
                    
                    If temp(0) = "1" Then
                        SortObj(I).TrueLeft = obj.Left + Pic1_Left(temp(1))
                        SortObj(I).TrueTop = obj.Top + Pic1_Top(temp(1))
                    ElseIf temp(0) = "3" Then
                        SortObj(I).TrueLeft = obj.Left + Pic3_Left(temp(1))
                        SortObj(I).TrueTop = obj.Top + Pic3_Top(temp(1))
                    End If
                Else
                    SortObj(I).TrueLeft = obj.Left
                    SortObj(I).TrueTop = obj.Top
                End If
            Else
                SortObj(I).TrueLeft = obj.Left
                SortObj(I).TrueTop = obj.Top
            End If
            
            
            I = I + 1
        End If
    Next
    '/**/
    
    '/*�w��o�Ǫ��󪺸�ơA���쥻�b���W����m���ӱƧ�*/
    For I = 0 To All_Count - 1
        For j = I + 1 To All_Count - 1
            If SortObj(I).TrueTop > SortObj(j).TrueTop Or (SortObj(I).TrueLeft > SortObj(j).TrueLeft And SortObj(I).TrueTop >= SortObj(j).TrueTop) Then
                Call swap(SortObj(I).Name, SortObj(j).Name)
                Call swap(SortObj(I).Left, SortObj(j).Left)
                Call swap(SortObj(I).Top, SortObj(j).Top)
                Call swap(SortObj(I).Text, SortObj(j).Text)
                Call swap(SortObj(I).ID, SortObj(j).ID)
                Call swap(SortObj(I).TrueLeft, SortObj(j).TrueLeft)
                Call swap(SortObj(I).TrueTop, SortObj(j).TrueTop)
            End If
        Next
    Next
    '/**/
    
    '/*�̤W�U���k�����ǡA�N��g���@���r�ꪫ�󤤡A���ݭn���ܤ]�i�H�g�i�����(�C�@�Ӫ���b���W���רC�Z��500���A�b��r�ɤ��Y�_�@����)*/
    Login_LastOpenReportBody = ""
    'Open "test.txt" For Output As #1
        For I = 0 To All_Count - 2
            Call Replace(SortObj(I).Text, " ", "")
            
            'Print #1, SortObj(i).Text, ;
            If SortObj(I).Text <> "" Then
                Login_LastOpenReportBody = Login_LastOpenReportBody & SortObj(I).Text
            End If
            
            If SortObj(I + 1).TrueTop > SortObj(I).TrueTop Then
                diff = SortObj(I + 1).TrueTop - SortObj(I).TrueTop
                
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
        Login_LastOpenReportBody = Login_LastOpenReportBody & SortObj(I).Text & vbCrLf
    'Close #1
    'Debug.Print Len(Login_LastOpenReportBody)
    '/**/
    
    
    '/*�g�J�ڭ̪���Ʈw�A�H��MPPS�W�Ǩ�HIS����Ʈw��*/
    If Connection.State Then
        I = InStr(1, Login_LastOpenReportBody, "SPIROMETRIC INTERPRETATION")
        Login_LastOpenReportBody = Right(Login_LastOpenReportBody, Len(Login_LastOpenReportBody) - I + 1)
        
        Connection.Execute "update cris_exam_online set item1='" & Login_LastOpenReportBody & "' where uni_key='" & Login_LastOpenReportUnikey & "' "
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
    
    Dim I As Integer
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
    For I = 0 To MAX_FIELD - 1
        ReportValue(I) = ReportValue_PSV(I) & "/" & ReportValue_EDV(I)
    Next
    NI_Report_Load_From_File = ReportValue
    '/**/
    
    
    If False Then
errout:
        Select Case Err.Number
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
    fps.Sheet = SheetID
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
        Select Case Err.Number
        Case 9
            Resume Next
        Case Else
            Call PrintLog("Error In NI_Spread_Report_Load_From_File")
            NI_Spread_Report_Load_From_File = False
        End Select
    End If
End Function
'/*20100122*/


