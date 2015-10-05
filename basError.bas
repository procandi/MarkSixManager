Attribute VB_Name = "basError"
'/******************************************************************/
'/*說明：　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*放置處理發生例外狀況時用到的函式、變數、常數等的地方。          */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*模組：　　　　　　　　　　　　　　　    　　　　　　　　　　　　*/
'/*basFile.bas及basVariable.bas。                                  */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*參考：　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*scrrun.dll。　　　　　　　　　　　　　　　　　　　　　　　　　  */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*元件：　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*無。　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*注意事項：　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*                                           Edit By Edward Hsieh */
'/*                                      Last Edit Date 2009/03/13 */
'/******************************************************************/
Option Explicit



'/**************************印出Log檔至指定的路徑，若路徑不存在則會自動製作***************************/
Public Function PrintLog(ByVal Log_String As String, Optional ByVal FileName As String = "") As String
    Dim FSO_FileExist As New FileSystemObject
    Dim SavePath As String
    Dim SaveFile As String
    Dim SaveDate As String
    Dim SaveTime As String
    
    
    SavePath = App.Path & "\log\"
    SaveFile = FileName & Format(DateTime.Date, "yyyyMMdd") & ".log"
    SaveDate = DateTime.Date
    SaveTime = DateTime.Time
    
    
    If Not FSO_FileExist.FolderExists(SavePath) Then
        Call CreatePath(SavePath)
    End If
    

    FreeFilePort = FreeFile
    Open SavePath & SaveFile For Append As #FreeFilePort
        Print #FreeFilePort, SaveDate, SaveTime, "說明-" & Log_String
        Print #FreeFilePort, SaveDate, SaveTime, "代碼-" & err.Number
        Print #FreeFilePort, SaveDate, SaveTime, "訊息-" & err.Description
    Close #FreeFilePort
    
    PrintLog = SavePath & SaveFile
End Function
'/**************************小華修改的(2009/03/13)***********************************/



'/**************************顯示錯誤訊息，並記錄到Log檔中，以便排除錯誤***********************************/
Public Function ErrorOut(ByVal Log_String As String, Optional ByVal FileName As String = "") As Boolean
    Dim Log_File As String
    
    Log_File = PrintLog(Log_String, FileName)
    MsgBox "發生系統錯誤：" & vbCrLf & "說明-" & Log_String & vbCrLf & "代碼-" & err.Number & vbCrLf & "訊息-" & err.Description & vbCrLf & vbCrLf & "請將錯誤記錄檔" & Log_File & "通知並寄予軒崴系統工程師", vbCritical
    
    ErrorOut = True
End Function
'/**************************小華修改的(2009/03/13)***********************************/
