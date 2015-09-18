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
Public Const COMMANDFROM_INI As String = "./CommandFrom.ini" '記錄一些CommandFrom所需的存檔位置等資料
Public Const CHECKPRINTERSTATUS_INI As String = "./CheckPrinterStatus.ini" '記錄一些CheckPrinterStatus所需的存檔位置等資料
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


'/*用於跟frmViewDcmList溝通的變數*/
Public imgList As Integer
Public imgFilePath() As String
Public imgFileName() As String
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
                    SortObj(i).Text = "■" & obj.Caption
                Else
                    SortObj(i).Text = "□" & obj.Caption
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
    
    '/*針對這些物件的資料，把其原本在表單上的位置做個排序*/
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
    
    '/*依上下左右的順序，將其寫成一份字串物件中，有需要的話也可以寫進文件檔(每一個物件在表單上高度每距離500單位，在文字檔中即斷一次行)*/
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
    
    
    '/*寫入我們的資料庫，以讓MPPS上傳到HIS的資料庫中*/
    If Connection.State Then
        i = InStr(1, Login_LastOpenReportBody, "SPIROMETRIC INTERPRETATION")
        Login_LastOpenReportBody = Right(Login_LastOpenReportBody, Len(Login_LastOpenReportBody) - i + 1)
        Call DBRecordLog("update", "update cris_exam_online set item1='" & Login_LastOpenReportBody & "' where status<>'已刪除' and uni_key='" & Login_LastOpenReportUnikey & "' ", "讓MPPS上傳到HIS的資料庫cris_exam_online")
        Connection.Execute "update cris_exam_online set item1='" & Login_LastOpenReportBody & "' where status<>'已刪除' and uni_key='" & Login_LastOpenReportUnikey & "' "
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
    
    Dim i As Integer
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
    fps.sheet = SheetID
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


