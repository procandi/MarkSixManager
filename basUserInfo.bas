Attribute VB_Name = "basUserInfo"
'/******************************************************************/
'/*說明：　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*放置跟使用者個人設定、資料有關的地方。                          */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*模組：　　　　　　　　　　　　　　　    　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*參考：    　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*msscript.dll。                                                  */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*元件：    　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*FPSPR70.OCX。                                                   */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*注意事項：　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*                                           Edit By Edward Hsieh */
'/*                                      Last Edit Date 2009/02/26 */
'/******************************************************************/
Option Explicit



'/**************************公用的登入使用者常數資料***********************************/

'/**/
Public Const CHARTNO_LENGTH As Integer = 10 '本系統的病歷號長度
Public Const EMERGENCY_ID As String = "55555" '心內簡訊會用到的，緊急寄簡訊的代碼
'/**/


'/*因為ni_svr_queue已經不太可能使用到了，故這部份的都先關掉*/
'Public Const NI_REPORT_TEMPLATE_PATH_A As String = "./JS_Template_A.txt" '頸超用到A報表的模版路徑
'Public Const NI_REPORT_TEMPLATE_PATH_B As String = "./JS_Template_B.txt" '頸超用到B報表的模版路徑
'Public Const NI_REPORT_TEMPLATE_PATH_C As String = "./JS_Template_C.txt" '頸超用到C報表的模版路徑
'Public Const NI_REPORT_TEMPLATE_PATH_D As String = "./JS_Template_D.txt" '頸超用到D報表的模版路徑
'Public Const NI_REPORT_TEMPLATE_PATH_E As String = "./JS_Template_E.txt" '頸超用到E報表的模版路徑
'Public Const NI_REPORT_TEMPLATE_PATH_F As String = "./JS_Template_F.txt" '頸超用到F報表的模版路徑
'Public Const MAX_NI_FIELD_A As Integer = 50 '頸超用到A報表的欄位數量
'Public Const MAX_NI_FIELD_B As Integer = 48 '頸超用到B報表的欄位數量
'Public Const MID_NI_FIELD_C_AND_D As Integer = 28 '頸超用到C、D報表的欄位數量
'Public Const MAX_NI_FIELD_E As Integer = 16 '頸超用到E報表的欄位數量
'Public Const MAX_NI_FIELD_F As Integer = 28 '頸超用到F報表的欄位數量
'/**/


'/**/
Public Const NI_REPORT_TEMPLATE_PATH_NECK_TCI As String = "./NI_Template_NECK_TCI.ini" '頸超用到NECK_TCI報表的模版路徑
Public Const NI_REPORT_TEMPLATE_PATH_NECK As String = "./NI_Template_NECK.ini" '頸超用到NECK報表的模版路徑
Public Const NI_REPORT_TEMPLATE_PATH_TCI As String = "./NI_Template_TCI.ini" '頸超用到TCI報表的模版路徑
Public Const NI_REPORT_TEMPLATE_PATH_LimpUpper As String = "./NI_Template_LimpUpper.ini" '頸超用到LimpUpper報表的模版路徑
Public Const NI_REPORT_TEMPLATE_PATH_LimpLower As String = "./NI_Template_LimpLower.ini" '頸超用到LimpLower報表的模版路徑
'/**/


'/*跟ini處理有關的常數*/
Public Const EXAMSVR_INI As String = "./ExamSVR.ini" '記錄有用於連線到HIS抓資料的連結字串等資料的ini檔路徑(國泰的)
Public Const EXAMHCR_INI As String = "./ExamHCR.ini" '記錄有用於連線到HIS抓資料的連結字串等資料的ini檔路徑(國泰高階的)
Public Const SCHEDULE_PRINTER_INI As String = "./Schedule_Printer.ini" '記錄Schdule該程式，所有有關印表機的設定資料的ini檔路徑(不分版本)
Public Const TIF_MDI_2_JPG_INI As String = "./TIF_MDI_2_JPG.ini" '記錄TIF_MDI_2_JPG該程式，所有預設的檔案存取路徑，以及要抓取的dialog視窗的名稱在此設定(不分版本)
Public Const TIF_MDI_2_DCM_INI As String = "./TIF_MDI_2_DCM.ini" '記錄TIF_MDI_2_DCM該程式，所有預設的檔案存取路徑，以及要抓取的dialog視窗的名稱在此設定(不分版本)
Public Const PDF_2_JPG_INI As String = "./PDF_2_JPG.ini" '記錄PDF2JPG該程式，所有預設的檔案存取路徑，以及要抓取的dialog視窗的名稱在此設定(不分版本)
Public Const SVR_HC_PDF_2_JPG_INI As String = "./SVR_HC_PDF_2_JPG.ini" '記錄SVR_HC_PDF2JPG該程式，所有預設的檔案存取路徑，以及要抓取的dialog視窗的名稱在此設定(新竹版本)
Public Const REFERENCE_INI As String = "./Reference.ini" '記錄一些跟擷取有關的資料(不分版本)
Public Const HISSYNC_INI As String = "./HISSync.ini" '記錄一些跟HISSync要下載哪些資料有關的資料(不分版本)
Public Const SWCONFIG_TAIAN_INI As String = "./SWConfig_TAIAN.ini" '記錄一些跟HISSync要下載哪些資料有關的資料(不分版本)
Public Const CATH_XML_INI As String = "./Cath_XML.ini" '記錄一些XML轉換系統會用到的XML預設存放位置
Public Const PACS_UPLOAD_INI As String = "./Pacs_Upload.ini" '記錄一些Dicom_Upload該程式，預設存資料庫的位置等
Public Const OFFLINE_REPORT_INI As String = "./OfflineReport.ini" 'OfflineReport那支程式會用到的設定資料
Public Const DMMHE_INI As String = "./DMMHE.ini" '記錄一些Delete_MMH_Exam該程式，要刪多少日期以上的資料夾及資料庫資料等
Public Const ERASE_LOG_INI As String = "./Erase_Log.ini" '記錄一些Erase_Log該程式，要移開多少日期以上的特定副檔名資料等
Public Const DRAFT_INI As String = "./Draft.ini" '記錄一些Draft所需的存檔位置等資料
Public Const COMMANDFROM_INI As String = "./CommandFrom.ini" '記錄一些CommandFrom的指令等資料
Public Const CHECKPRINTERSTATUS_INI As String = "./CheckPrinterStatus.ini" '記錄一些CheckPrinterStatus的狀態記錄等資料
Public Const SYSTEMSYNC_INI As String = "./SystemSync.ini" '記錄一些SystemSync所需的同步資料夾的路徑等資料
'/*小華修改的(2009/02/04)*/

'/**************************小華修改的(2009/04/14)***********************************/


'/**************************公用的登入使用者變數資料***********************************/

'/*記錄登入者的資訊*/
Public Login_ID As String '使用者帳號，對應資料庫cris_user的logonuser
Public Login_PW As String '使用者密碼，對應資料庫cris_user的password
Public Login_Name As String '使用者名稱，對應資料庫cris_user的name
Public Login_Position As String '使用者職稱，對應資料庫cris_user的type
Public Login_No As String '使用者編號，對應資料庫cris_user的userid
Public Login_Power As Integer '使用者權限，對應資料庫cris_user的authorid
Public Login_Phone As String '使用者電話，對應資料庫cris_user的phone
Public Login_HostName As String '使用者主機名稱，對應ini裡的hostname
Public Login_System As String '使用者科別，對應資料庫cris_user的system
'/*小華修改的(2009/03/18)*/


'/*記錄開啟的報表的資訊*/
Public Login_LastOpen As String '已無法得知
Public Login_LastOpenReportList As String ''使用者最後開的報表內容代號(高階的主畫面用)
Public Login_LastOpenReportBody As String '使用者最後開的報表內容，對應資料庫cris_exam_online的item6
Public Login_LastOpenReportType As String '使用者最後開的報表的檢查類別，對應資料庫cris_exam_online的type
Public Login_LastOpenReportStatus As String '使用者最後開的報表的報告狀態，對應資料庫cris_exam_online的status
Public Login_LastOpenReportUnikey As String '使用者最後開的報表的報告unikey，對應資料庫cris_exam_online的uni_key
Public Login_LastOpenReportChartNO As String '使用者最後開的報表的報告的病歷號，對應資料庫cris_exam_online的chartno
Public Login_LastOpenReportChartName As String '使用者最後開的報表的報告病患姓名
Public Login_LastOpenReportDrFrom As String '使用者最後開的報表的報告的診斷來源，對應資料庫cris_exam_online的dr_from
Public Login_LastOpenReportModifyDate As String '使用者最後開的報表的確認日期，對應資料庫cris_exam_online的examdate
Public Login_LastOpenReportModifyTime As String '使用者最後開的報表的確認時間，對應資料庫cris_exam_online的examtime
Public Login_LastOpenReportReturnDate As String '使用者最後開的報表的報告日期，對應資料庫cris_exam_online的reportdate
Public Login_LastOpenReportReturnTime As String '使用者最後開的報表的報告時間，對應資料庫cris_exam_online的reporttime
Public Login_LastOpenReportCreateDRName As String '使用者最後開的報表的開單醫生名字，對應資料庫cris_exam_online的dr_on
Public Login_LastOpenReportCreateDRPhone As String '使用者最後開的報表的開單醫生電話
Public Login_LastOpenReportFieldList As String '使用者最後開的報表的選擇的欄位
Public Login_LastOpenReportSystem As String '使用者最後開的報表的科別，對應資料庫cris_exam_online的system
'/*小華修改的(2009/03/18)*/



'/*用於記錄frmQueue的資料，以傳到frmReport該頁使用*/
Public Report_DrReport As String '用於記錄frmQueue的報告醫師
Public Report_DrOn As String '用於記錄frmQueue的檢查醫師
'/*小華修改的(2009/03/18)*/


'/*用於記錄frmSpread的資料，以傳到frmSpread該頁使用*/
Public Spread_ID As String
Public Spread_Name As String
'/*小華修改的(2009/09/14)*/


'/*記錄使用者在使用frmImageSelect的設定*/
Public Img_Sel_Cols As Long '儲存在frmImageSelect的使用者的橫排預設該有幾張的設定
Public Img_Sel_Rows As Long '儲存在frmImageSelect的使用者的直排預設該有幾張的設定
'/*小華修改的(2009/03/18)*/


'/*記錄使用者在使用frmRepPreview的設定*/
Public Img_Sel_Opt As Long '儲存在frmRepPreview的列印格式預設要設幾乘幾
'/*小華修改的(2009/03/18)*/


'/*記錄使用者在使用frmReport的設定*/
Public defaultImgZoomLeft As Long '記錄在frmReport的圖片大小改變前的預設X軸位置
Public defaultImgZoomTop As Long '記錄在frmReport的圖片大小改變前的預設Y軸位置
Public defaultImgZoomWidth As Long '記錄在frmReport的圖片大小改變前的預設寬度
Public defaultImgZoomHeight As Long '記錄在frmReport的圖片大小改變前的預設高度
'/*小華修改的(2009/04/15)*/


'/*記錄使用者在使用frmRepPreview的設定*/
Public ImgSvr_Path As String '遠端要登入進去的Image Server的路徑
'/*小華修改的(2009/04/16)*/


'/*用來決定主任是否能看這份報告*/
Public MasterCanNotSee As Boolean
'/*小華修改的(2009/06/26)*/


'/*記錄跟英飛特溝通電子病歷程式時會用到的欄位位置常數*/
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

'/**************************小華修改的(2009/02/03)***********************************/



'/*************************公用的登入使用者結構資料**********************************/

'/*用於上傳HIS那個副程式，把畫面上所有的物件特有的屬性都COPY下來用的TYPE*/
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

'/**************************小華修改的(2009/02/03)***********************************/



'/*                  把指定的表單及其上的物件，全轉換成文字檔，並上傳報告到HIS                                     */
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
    
    '/*找出非視窗盒、命令按鈕及圖片的物件共有幾個*/
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
    
    '/*定義一下要抓的物件的大小*/
    ReDim SortObj(All_Count)
    ReDim Pic1_Left(Pic1_Count)
    ReDim Pic1_Top(Pic1_Count)
    ReDim Pic3_Left(Pic3_Count)
    ReDim Pic3_Top(Pic3_Count)
    '/**/
    
    '/*把所有的picturebox的資料抓出來*/
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
    
    '/*把所有非視窗盒、命令按鈕及圖片的資料抓出來*/
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
                    SortObj(I).Text = "■" & obj.Caption
                Else
                    SortObj(I).Text = "□" & obj.Caption
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
    
    '/*針對這些物件的資料，把其原本在表單上的位置做個排序*/
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
    
    '/*依上下左右的順序，將其寫成一份字串物件中，有需要的話也可以寫進文件檔(每一個物件在表單上高度每距離500單位，在文字檔中即斷一次行)*/
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
    
    
    '/*寫入我們的資料庫，以讓MPPS上傳到HIS的資料庫中*/
    If Connection.State Then
        I = InStr(1, Login_LastOpenReportBody, "SPIROMETRIC INTERPRETATION")
        Login_LastOpenReportBody = Right(Login_LastOpenReportBody, Len(Login_LastOpenReportBody) - I + 1)
        
        Connection.Execute "update cris_exam_online set item1='" & Login_LastOpenReportBody & "' where uni_key='" & Login_LastOpenReportUnikey & "' "
    End If
    '/**/
End Sub
'/**/
'/**************************小華修改的(2009/07/03)***********************************/




'/*用以讀取文字檔中的template及儀器檔，並填到表單上的副程式*/
Public Function NI_Report_Load_From_File(ByVal ReportFilePath As String, ByVal TemplateFilePath As String, ByVal MAX_FIELD As Integer) As String()
    On Error GoTo errout:

    '/*分別是，整理傳回檔用的report系列變數，讀檔案時的暫用變數，公用變數*/
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
    
    '/*script物件，語言採用java script*/
    Dim MSSC As New MSScriptControl.ScriptControl
    MSSC.Language = "JavaScript"
    '/**/
    
    
    '/*採用java script的eval的函數，把每個名稱應該是在陣列幾存到java script的變數中*/
    FreeFilePort = FreeFile
    Open TemplateFilePath For Input As #FreeFilePort
        Do Until EOF(1)
            Line Input #1, InputValue
            
            FieldValue = Split(InputValue, "=")
            Call MSSC.Eval("var " & FieldValue(0) & "=" & FieldValue(1))
        Loop
    Close #FreeFilePort
    '/**/


    '/*把從文字檔中的數值，用java script填入應該所屬的欄位中。填入前要先依psv跟edv來分左右*/
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
    
    
    '/*將本來分隔為左及右的檔案轉存到文字檔中，並傳回以便利用*/
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




'/*用以讀取文字檔中的template及儀器檔，並填到Spread表單上的副程式*/
Public Function NI_Spread_Report_Load_From_File(ByVal ReportFilePath As String, ByVal TemplateFilePath As String, ByRef fps As fpSpread, ByVal SheetID As Integer) As Boolean
    On Error GoTo errout:

    '/*分別是，用來放置讀出來的spread位置的變數，讀檔案時的暫用變數*/
    Dim ReportValuePlace() As String
    Dim TemplateValuePlace As String
    
    Dim InputValue As String
    Dim FieldWord() As String
    '/**/
    
    '/*指定這次要更新的是哪一張sheet*/
    fps.Sheet = SheetID
    '/**/
    
    '/*把從文字檔中讀出的數值，填入到spread報表物件中。*/
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


