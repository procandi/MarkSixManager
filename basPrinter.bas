Attribute VB_Name = "basPrinter"
'/******************************************************************/
'/*說明：　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*放置跟印表機操作、設定等有關的地方。                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*模組：　　　　　　　　　　　　　　　    　　　　　　　　　　　　*/
'/*basString.bas。                                                 */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*參考：　　　　　　　　　　　　　　　    　　　　　　　　　　　　*/
'/*MDIVWCTL.DLL。                                                  */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*元件：    　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*注意事項：　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*                                           Edit By Edward Hsieh */
'/*                                      Last Edit Date 2009/04/09 */
'/******************************************************************/
Option Explicit



'/*印表機相關Windows API常數*/
Public Const ERROR_INSUFFICIENT_BUFFER = 122
Public Const PRINTER_STATUS_BUSY = &H200
Public Const PRINTER_STATUS_DOOR_OPEN = &H400000
Public Const PRINTER_STATUS_ERROR = &H2
Public Const PRINTER_STATUS_INITIALIZING = &H8000
Public Const PRINTER_STATUS_IO_ACTIVE = &H100
Public Const PRINTER_STATUS_MANUAL_FEED = &H20
Public Const PRINTER_STATUS_NO_TONER = &H40000
Public Const PRINTER_STATUS_NOT_AVAILABLE = &H1000
Public Const PRINTER_STATUS_OFFLINE = &H80
Public Const PRINTER_STATUS_OUT_OF_MEMORY = &H200000
Public Const PRINTER_STATUS_OUTPUT_BIN_FULL = &H800
Public Const PRINTER_STATUS_PAGE_PUNT = &H80000
Public Const PRINTER_STATUS_PAPER_JAM = &H8
Public Const PRINTER_STATUS_PAPER_OUT = &H10
Public Const PRINTER_STATUS_PAPER_PROBLEM = &H40
Public Const PRINTER_STATUS_PAUSED = &H1
Public Const PRINTER_STATUS_PENDING_DELETION = &H4
Public Const PRINTER_STATUS_PRINTING = &H400
Public Const PRINTER_STATUS_PROCESSING = &H4000
Public Const PRINTER_STATUS_TONER_LOW = &H20000
Public Const PRINTER_STATUS_USER_INTERVENTION = &H100000
Public Const PRINTER_STATUS_WAITING = &H2000
Public Const PRINTER_STATUS_WARMING_UP = &H10000
Public Const JOB_STATUS_PAUSED = &H1
Public Const JOB_STATUS_ERROR = &H2
Public Const JOB_STATUS_DELETING = &H4
Public Const JOB_STATUS_SPOOLING = &H8
Public Const JOB_STATUS_PRINTING = &H10
Public Const JOB_STATUS_OFFLINE = &H20
Public Const JOB_STATUS_PAPEROUT = &H40
Public Const JOB_STATUS_PRINTED = &H80
Public Const JOB_STATUS_DELETED = &H100
Public Const JOB_STATUS_BLOCKED_DEVQ = &H200
Public Const JOB_STATUS_USER_INTERVENTION = &H400
Public Const JOB_STATUS_RESTART = &H800

' constants for PRINTER_DEFAULTS structure
Public Const PRINTER_ACCESS_USE = &H8
Public Const PRINTER_ACCESS_ADMINISTER = &H4

' constants for DEVMODE structure
Public Const CCHDEVICENAME = 32
Public Const CCHFORMNAME = 32
'/**/

'/*印表機相關Windows API結構*/
Public Type PRINTER_DEFAULTS
   pDatatype As String
   pDevMode As Long
   DesiredAccess As Long
End Type
Public Type DEVMODE
   dmDeviceName As String * CCHDEVICENAME
   dmSpecVersion As Integer
   dmDriverVersion As Integer
   dmSize As Integer
   dmDriverExtra As Integer
   dmFields As Long
   dmOrientation As Integer
   dmPaperSize As Integer
   dmPaperLength As Integer
   dmPaperWidth As Integer
   dmScale As Integer
   dmCopies As Integer
   dmDefaultSource As Integer
   dmPrintQuality As Integer
   dmColor As Integer
   dmDuplex As Integer
   dmYResolution As Integer
   dmTTOption As Integer
   dmCollate As Integer
   dmFormName As String * CCHFORMNAME
   dmLogPixels As Integer
   dmBitsPerPel As Long
   dmPelsWidth As Long
   dmPelsHeight As Long
   dmDisplayFlags As Long
   dmDisplayFrequency As Long
End Type
Public Type SYSTEMTIME
   wYear As Integer
   wMonth As Integer
   wDayOfWeek As Integer
   wDay As Integer
   wHour As Integer
   wMinute As Integer
   wSecond As Integer
   wMilliseconds As Integer
End Type
Public Type JOB_INFO_2
   JobId As Long
   pPrinterName As Long
   pMachineName As Long
   pUserName As Long
   pDocument As Long
   pNotifyName As Long
   pDatatype As Long
   pPrintProcessor As Long
   pParameters As Long
   pDriverName As Long
   pDevMode As Long
   pStatus As Long
   pSecurityDescriptor As Long
   Status As Long
   Priority As Long
   Position As Long
   StartTime As Long
   UntilTime As Long
   TotalPages As Long
   Size As Long
   Submitted As SYSTEMTIME
   time As Long
   PagesPrinted As Long
End Type
Public Type PRINTER_INFO_2
   pServerName As Long
   pPrinterName As Long
   pShareName As Long
   pPortName As Long
   pDriverName As Long
   pComment As Long
   pLocation As Long
   pDevMode As Long
   pSepFile As Long
   pPrintProcessor As Long
   pDatatype As Long
   pParameters As Long
   pSecurityDescriptor As Long
   Attributes As Long
   Priority As Long
   DefaultPriority As Long
   StartTime As Long
   UntilTime As Long
   Status As Long
   cJobs As Long
   AveragePPM As Long
End Type
'/**/


'/*印表機相關Windows API函式*/
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long
Public Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Byte, ByVal cbBuf As Long, pcbNeeded As Long) As Long
Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function EnumJobs Lib "winspool.drv" Alias "EnumJobsA" (ByVal hPrinter As Long, ByVal FirstJob As Long, ByVal NoJobs As Long, ByVal Level As Long, pJob As Byte, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
'/**/


'/*******************************跟印表機有關的變數*************************/
Public DefaultPrinter As Printer '記錄使用者預設的印表機
Public DefaultPrinterName As String '記錄使用者預設的印表機名稱
Public PrinterFilePath As String '記錄印表機Microsoft Office Document Image Writer的存檔路徑
Public Printer_A4_X As Integer '記錄使用者預設要印在紙上的起始X軸位置
Public Printer_A4_Y As Integer '記錄使用者預設要印在紙上的起始Y軸位置
Public Printer_A4_Width As Integer '記錄使用者預設要印在紙上的寬度
Public Printer_A4_Height As Integer '記錄使用者預設要印在紙上的高度
'/**************************小華修改的(2009/04/09)***********************************/


'/*******************************跟印表機有關的常數*************************/
Public Const MODIW_PRINTER As String = "Microsoft Office Document Image Writer" '電腦內轉mdi及tif的不印紙本但會輸出檔案的印表機的名字
Public Const PDFC_PRINTER As String = "PDFCreator" '電腦內轉pdf的不印紙本但會輸出檔案的印表機的名字
'/**************************小華修改的(2009/04/09)***********************************/


'/**************************一些跟印表機有關的範例***********************************/
'Sub SomePrinterFunction()
'    '/*印出預設印表機*/
'    Debug.Print Printer.DeviceName
'
'    '/*更改預設印表機*/
'    CreateObject("WScript.Network").SetDefaultPrinter "hp LaserJet 1010 Series Driver" ' 設定預設印表機,傳入印表機名稱
'
'    '/*印出所有印表機*/
'    Dim prn As Printer
'    For Each prn In Printers
'        Debug.Print prn.DeviceName
'    Next
'
'    '/*用預設印表機列印*/
'    Printer.Print "test1"
'    Printer.Print "test2"
'
'    '/*結束預設印表機的傳送，並開始列印*/
'    Printer.EndDoc
'End Sub
'/**************************小華修改的(2009/04/09)***********************************/



'/**************************列印指定的mdi檔，傳回值為>=0則代表正常結束，<0則代表運作失敗，另外傳回傳的數值代表列印的數量***********************************/
Public Function Printer_By_MODIW(ByVal Source_FilePath As String, ByVal Source_FileName As String) As Long
    On Error GoTo errout:


    Dim miDoc As New MODI.Document
    Dim miImg As MODI.Image
    
    
    '/*確保檔案路徑是正確的*/
    If Right(Source_FilePath, 1) <> "\" Then
        Source_FilePath = Source_FilePath & "\"
    End If
    '/**/
    
    
    Call miDoc.Create(Source_FilePath & Source_FileName)
    Call miDoc.PrintOut
    
    Set miImg = Nothing
    Call miDoc.Close(False)
    Set miDoc = Nothing
    
    Printer_By_MODIW = miDoc.Images.Count - 1
    
    If False Then
errout:
        Call PrintLog("Printer_By_MODIW-Not Import A Current Image File!!")
        Printer_By_MODIW = -1
    End If
End Function
'/**************************小華修改的(2009/04/09)***********************************/




'/**************************列印指定的mdi檔，傳回值為>=0則代表正常結束，<0則代表運作失敗，另外傳回傳的數值代表列印的數量***********************************/
Public Function Printer_By_MODIW_EX(ByVal Source_FilePath As String, ByVal Source_FileName As String, ByVal Rotate As Integer) As Long
    On Error GoTo errout:


    Dim miDoc As New MODI.Document
    Dim miImg As MODI.Image
    
    
    '/*確保檔案路徑是正確的*/
    If Right(Source_FilePath, 1) <> "\" Then
        Source_FilePath = Source_FilePath & "\"
    End If
    '/**/
    
    
    Call miDoc.Create(Source_FilePath & Source_FileName)
    
    Dim i As Long
    For i = 0 To miDoc.Images.Count - 1
        Set miImg = miDoc.Images(i)
        Call miImg.Rotate(Rotate)
    Next
    Call miDoc.PrintOut
    
    Set miImg = Nothing
    Call miDoc.Close(False)
    Set miDoc = Nothing
    
    Printer_By_MODIW_EX = i
    
    If False Then
errout:
        Call PrintLog("Printer_By_MODIW_EX-Not Import A Current Image File!!")
        Printer_By_MODIW_EX = -1
    End If
End Function
'/**************************小華修改的(2010/01/11)***********************************/


'/*取得指定名稱的印表機的hwnd*/
Function GetPrinterHWnd(ByVal PrinterName As String) As Long
    Dim hPrinter As Long
    Dim result As Long
    Dim pDefaults As PRINTER_DEFAULTS

    '/*設定印表機的安全性權限為允許存取*/
    pDefaults.DesiredAccess = PRINTER_ACCESS_USE
   
    '/*呼叫WindowsAPI去打開該指定名稱的印表機*/
    result = OpenPrinter(PrinterName, hPrinter, pDefaults)
    
    If result = 0 Then
        Call PrintLog("Cannot open printer " & PrinterName & ", Error: " & err.LastDllError)
        GetPrinterHWnd = -1
    Else
        GetPrinterHWnd = hPrinter
    End If
End Function


'/*取得指定名稱的印表機目前狀態*/
Public Function GetPrinterStatus(ByRef PrinterName As String) As String
    '/*變數*/
    Dim i As Integer
    Dim tempStr As String
    Dim result As Long
    
    Dim hPrinter As Long
    Dim PrinterStr As String
    Dim ByteBuf As Long
    Dim BytesNeeded As Long
    
    Dim PI2 As PRINTER_INFO_2
    Dim PrinterInfo() As Byte
    
    
    '/*取得指定名稱印表機的hwnd*/
    hPrinter = GetPrinterHWnd(PrinterName)
    
    '/*取得指定hwnd的印表機資訊數量*/
    result = GetPrinter(hPrinter, 2, 0&, 0&, BytesNeeded)
    
    '/*要是在取得印表機資訊數量時就往生的話即直接關閉，並離開。其餘的才繼續往下處理*/
    If err.LastDllError <> ERROR_INSUFFICIENT_BUFFER Then
        Call PrintLog(" > GetPrinter Failed on initial call! <")
        Call ClosePrinter(hPrinter)
        GetPrinterStatus = "error"
    Else
        '/*定義一個跟資訊數量一樣大的陣列，以在等一下放印表機的資訊進去*/
        ReDim PrinterInfo(BytesNeeded)
        ByteBuf = BytesNeeded
                
        '/*取得指定hwnd的印表機資訊*/
        result = GetPrinter(hPrinter, 2, PrinterInfo(0), ByteBuf, BytesNeeded)
        
        '/*要是在取得印表機資訊時就往生的話即直接關閉，並離開。其餘的才繼續往下處理*/
        If result = 0 Then
            Call PrintLog("Couldn't get Printer Status!  Error = " & err.LastDllError())
            Call ClosePrinter(hPrinter)
            GetPrinterStatus = "error"
        Else
            '/*從記憶體層級把印表機資訊清單位移到指定的結構中*/
            Call CopyMemory(PI2, PrinterInfo(0), Len(PI2))
            
            
            '/*把印表機資訊一一依代碼判斷再存到清單中*/
            If PI2.Status = 0 Then
                PrinterStr = "Printer Status = Ready" & vbCrLf
            Else
                tempStr = ""
                If (PI2.Status And PRINTER_STATUS_BUSY) Then
                    tempStr = tempStr & "Busy  "
                End If
            
                If (PI2.Status And PRINTER_STATUS_DOOR_OPEN) Then
                    tempStr = tempStr & "Printer Door Open  "
                End If
                
                If (PI2.Status And PRINTER_STATUS_ERROR) Then
                    tempStr = tempStr & "Printer Error  "
                End If
                
                If (PI2.Status And PRINTER_STATUS_INITIALIZING) Then
                    tempStr = tempStr & "Initializing  "
                End If
                
                If (PI2.Status And PRINTER_STATUS_IO_ACTIVE) Then
                    tempStr = tempStr & "I/O Active  "
                End If
                
                If (PI2.Status And PRINTER_STATUS_MANUAL_FEED) Then
                    tempStr = tempStr & "Manual Feed  "
                End If
                
                If (PI2.Status And PRINTER_STATUS_NO_TONER) Then
                    tempStr = tempStr & "No Toner  "
                End If
                
                If (PI2.Status And PRINTER_STATUS_NOT_AVAILABLE) Then
                    tempStr = tempStr & "Not Available  "
                End If
                
                If (PI2.Status And PRINTER_STATUS_OFFLINE) Then
                    tempStr = tempStr & "Off Line  "
                End If
                
                If (PI2.Status And PRINTER_STATUS_OUT_OF_MEMORY) Then
                    tempStr = tempStr & "Out of Memory  "
                End If
                
                If (PI2.Status And PRINTER_STATUS_OUTPUT_BIN_FULL) Then
                    tempStr = tempStr & "Output Bin Full  "
                End If
                
                If (PI2.Status And PRINTER_STATUS_PAGE_PUNT) Then
                    tempStr = tempStr & "Page Punt  "
                End If
                
                If (PI2.Status And PRINTER_STATUS_PAPER_JAM) Then
                    tempStr = tempStr & "Paper Jam  "
                End If
                
                If (PI2.Status And PRINTER_STATUS_PAPER_OUT) Then
                    tempStr = tempStr & "Paper Out  "
                End If
                
                If (PI2.Status And PRINTER_STATUS_OUTPUT_BIN_FULL) Then
                    tempStr = tempStr & "Output Bin Full  "
                End If
                
                If (PI2.Status And PRINTER_STATUS_PAPER_PROBLEM) Then
                    tempStr = tempStr & "Page Problem  "
                End If
                
                If (PI2.Status And PRINTER_STATUS_PAUSED) Then
                    tempStr = tempStr & "Paused  "
                End If
                
                If (PI2.Status And PRINTER_STATUS_PENDING_DELETION) Then
                    tempStr = tempStr & "Pending Deletion  "
                End If
                
                If (PI2.Status And PRINTER_STATUS_PRINTING) Then
                    tempStr = tempStr & "Printing  "
                End If
                
                If (PI2.Status And PRINTER_STATUS_PROCESSING) Then
                    tempStr = tempStr & "Processing  "
                End If
                
                If (PI2.Status And PRINTER_STATUS_TONER_LOW) Then
                    tempStr = tempStr & "Toner Low  "
                End If
                
                If (PI2.Status And PRINTER_STATUS_USER_INTERVENTION) Then
                    tempStr = tempStr & "User Intervention  "
                End If
                
                If (PI2.Status And PRINTER_STATUS_WAITING) Then
                    tempStr = tempStr & "Waiting  "
                End If
                
                If (PI2.Status And PRINTER_STATUS_WARMING_UP) Then
                    tempStr = tempStr & "Warming Up  "
                End If
                
                If Len(tempStr) = 0 Then
                    tempStr = "Unknown Status of " & PI2.Status
                End If
                
                PrinterStr = "Printer Status = " & tempStr & vbCrLf
            End If
                           
                           
            '另外把印表機的名稱、driver、port給存起來
            PrinterStr = PrinterStr & "Printer Name = " & PointerValueToString(PI2.pPrinterName) & vbCrLf
            PrinterStr = PrinterStr & "Printer Driver Name = " & PointerValueToString(PI2.pDriverName) & vbCrLf
            PrinterStr = PrinterStr & "Printer Port Name = " & PointerValueToString(PI2.pPortName) & vbCrLf
                 
            
            '/*列舉完就關掉印表機並把值回傳*/
            Call ClosePrinter(hPrinter)
            GetPrinterStatus = PrinterStr
        End If
    End If
End Function
'/*20100527*/


'/*取得指定名稱的印表機目前工作*/
Public Function GetJobStatus(ByVal PrinterName As String) As String
    '/*變數*/
    Dim i As Integer
    Dim tempStr As String
    Dim result As Long
        
    Dim hPrinter As Long
    Dim JobStr As String
    Dim NumJI2 As Long
    Dim ByteBuf As Long
    Dim BytesNeeded As Long
    
    Dim JI2 As JOB_INFO_2
    Dim JobInfo() As Byte


    '/*取得指定名稱印表機的hwnd*/
    hPrinter = GetPrinterHWnd(PrinterName)
    
    
    '/*列舉出該印表機的所有工作*/
    Call EnumJobs(hPrinter, 0&, &HFFFFFFFF, 2, ByVal 0&, 0&, BytesNeeded, NumJI2)
   
   
    '/*如果沒有工作就show沒工作，並離開。其餘的再繼續處理*/
    If BytesNeeded = 0 Then
        JobStr = "No Print Jobs!"
        Call ClosePrinter(hPrinter)
        GetJobStatus = JobStr
    Else
        '/*定義一個跟工作數量一樣大的陣列，以在等一下放工作的資訊進去*/
        ReDim JobInfo(BytesNeeded)
        
        '/*把工作的資訊抓到儲存工作的陣列中*/
        result = EnumJobs(hPrinter, 0&, &HFFFFFFFF, 2, JobInfo(0), BytesNeeded, ByteBuf, NumJI2)
        
        '/*如果抓失敗就結束，並離開。其餘的再處理*/
        If result = 0 Then
            Call PrintLog(" > EnumJobs Failed on second call! <  Error = " & err.LastDllError)
            Call ClosePrinter(hPrinter)
            GetJobStatus = "error"
        Else
            For i = 0 To NumJI2 - 1
                '/*從記憶體層級把工作資訊清單位移到指定的結構中*/
                Call CopyMemory(JI2, JobInfo(i * Len(JI2)), Len(JI2))
                       
                '/*把工作資訊一一依代碼判斷再存到清單中*/
                JobStr = JobStr & "Job ID = " & JI2.JobId & vbCrLf & "Total Pages = " & JI2.TotalPages & vbCrLf
                tempStr = ""
                If JI2.pStatus = 0& Then
                    If JI2.Status = 0 Then
                        tempStr = tempStr & "Ready!  " & vbCrLf
                    Else
                        If (JI2.Status And JOB_STATUS_SPOOLING) Then
                            tempStr = tempStr & "Spooling  "
                        End If
                        
                        If (JI2.Status And JOB_STATUS_OFFLINE) Then
                            tempStr = tempStr & "Off line  "
                        End If
                        
                        If (JI2.Status And JOB_STATUS_PAUSED) Then
                            tempStr = tempStr & "Paused  "
                        End If
                        
                        If (JI2.Status And JOB_STATUS_ERROR) Then
                            tempStr = tempStr & "Error  "
                        End If
                        
                        If (JI2.Status And JOB_STATUS_PAPEROUT) Then
                            tempStr = tempStr & "Paper Out  "
                        End If
                        
                        If (JI2.Status And JOB_STATUS_PRINTING) Then
                            tempStr = tempStr & "Printing  "
                        End If
                        
                        If (JI2.Status And JOB_STATUS_USER_INTERVENTION) Then
                            tempStr = tempStr & "User Intervention Needed  "
                        End If
                        
                        If Len(tempStr) = 0 Then
                            tempStr = "Unknown Status of " & JI2.Status
                        End If
                    End If
                Else
                    tempStr = PointerCodeToString(JI2.pStatus)
                End If
                
                JobStr = JobStr & tempStr & vbCrLf
            Next
            
            
            '/*列舉完就關掉印表機並把值回傳*/
            Call ClosePrinter(hPrinter)
            GetJobStatus = JobStr
        End If
    End If
End Function
'/*20100527*/

