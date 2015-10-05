Attribute VB_Name = "basConvert"
'/******************************************************************/
'/*說明：　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*放置跟所有格式轉換有關的地方。                                  */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*模組：　　　　　　　　　　　　　　　    　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*參考：　　　　　　　　　　　　　　　    　　　　　　　　　　　　*/
'/*GflAx.dll及MDIVWCTL.DLL及DIjpg.dll。                            */
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


'/*                       一些轉檔會用到的函式庫                       */
Private Declare Function DIWriteJpg Lib "DIjpg.dll" (ByVal DestPath As String, ByVal quality As Long, ByVal progressive As Long) As Long
'/***********************小華修改的(2009/11/10)***************************/


'/*   將TIF檔轉成JPG格式的函式，傳回值為總共轉了幾張JPG檔              */
'/*   Example Input TIF_TO_JPG("C:","doc99.tif","C:","doc*.jpg")       */
Public Function TIF_TO_JPG(ByVal Source_FilePath As String, ByVal Source_FileName As String, ByVal Target_FilePath As String, ByVal Target_FileName As String) As Long
    On Error GoTo errout:
    
    Dim Convert As New GflAx.GflAx
     
    '/*確保檔案路徑是正確的*/
    If Right(Source_FilePath, 1) <> "\" Then
        Source_FilePath = Source_FilePath & "\"
    End If
    If Right(Target_FilePath, 1) <> "\" Then
        Target_FilePath = Target_FilePath & "\"
    End If
    '/**/
    
    
    
    '/*讀取來源檔案*/
    Convert.LoadBitmap Source_FilePath & Source_FileName
    '/**/

    
    '/*寫入檔案到目的地*/
    Dim i As Long
    Dim Temp_FileName As String
    
    For i = 0 To Convert.NumberOfPages - 1
        '/*另存的格式設定，這邊定死設定為JPEG*/
        Convert.SaveFormat = AX_JPEG
        Convert.SaveJPEGProgressive = True
        Convert.SaveJPEGQuality = 70
        '/**/
    
    
        Temp_FileName = Replace(Target_FileName, "*", i)
        Convert.SaveBitmap Target_FilePath & Temp_FileName
        Convert.NextPage
    Next
    '/**/
    
    
    
    '/*將Convert丟入GC(垃圾收集器)*/
    Set Convert = Nothing
    '/**/
    
    TIF_TO_JPG = i
    
    If False Then
errout:
        TIF_TO_JPG = -1
        Call PrintLog("TIF_TO_JPG-Not Import A Current Image File!")
    End If
End Function
'/***********************小華修改的(2009/04/09)***************************/




'/*   將TIF或MDI檔轉成JPG格式的函式，傳回值為總共轉了幾張JPG檔         */
'/*   Example Input TIF_MDI_TO_BMP("C:","doc99.tif","C:","doc*.jpg")   */
Public Function TIF_MDI_TO_BMP(ByVal Source_FilePath As String, ByVal Source_FileName As String, ByVal Target_FilePath As String, ByVal Target_FileName As String) As Long
    On Error GoTo errout:


    Dim miDoc As New MODI.Document
    Dim miImg As MODI.Image
    
    
    '/*確保檔案路徑是正確的*/
    If Right(Source_FilePath, 1) <> "\" Then
        Source_FilePath = Source_FilePath & "\"
    End If
    If Right(Target_FilePath, 1) <> "\" Then
        Target_FilePath = Target_FilePath & "\"
    End If
    '/**/
    

    Call miDoc.Create(Source_FilePath & Source_FileName)
    
    
    Dim i As Long
    Dim Temp_FileName As String
    
    For i = 0 To miDoc.Images.Count - 1
        Set miImg = miDoc.Images(i)
   
        Temp_FileName = Replace(Target_FileName, "*", i)
        Call SavePicture(miImg.Picture, Target_FilePath & Temp_FileName)
    Next
    
    
    Set miImg = Nothing
    Call miDoc.Close(False)
    Set miDoc = Nothing
    
    
    TIF_MDI_TO_BMP = i
    
    If False Then
errout:
        Call PrintLog("TIF_MDI_TO_BMP-Not Import A Current Image File!!")
        TIF_MDI_TO_BMP = -1
    End If
End Function
'/***********************小華修改的(2009/04/29)***************************/




'/*   將TIF或MDI檔轉成JPG格式的函式，傳回值為總共轉了幾張JPG檔               */
'/*   Example Input TIF_MDI_TO_BMP_EX("C:","doc99.tif","C:","doc*.jpg",90)   */
Public Function TIF_MDI_TO_BMP_EX(ByVal Source_FilePath As String, ByVal Source_FileName As String, ByVal Target_FilePath As String, ByVal Target_FileName As String, ByVal Rotate As Integer) As Long
    On Error GoTo errout:


    Dim miDoc As New MODI.Document
    Dim miImg As MODI.Image

    
    '/*確保檔案路徑是正確的*/
    If Right(Source_FilePath, 1) <> "\" Then
        Source_FilePath = Source_FilePath & "\"
    End If
    If Right(Target_FilePath, 1) <> "\" Then
        Target_FilePath = Target_FilePath & "\"
    End If
    '/**/
    

    Call miDoc.Create(Source_FilePath & Source_FileName)
    
    
    Dim i As Long
    Dim Temp_FileName As String
    
    For i = 0 To miDoc.Images.Count - 1
        Set miImg = miDoc.Images(i)
        Call miImg.Rotate(Rotate)
   
        Temp_FileName = Replace(Target_FileName, "*", i)
        Call SavePicture(miImg.Picture, Target_FilePath & Temp_FileName)
    Next
    
    
    Set miImg = Nothing
    Call miDoc.Close(False)
    Set miDoc = Nothing
    
    
    TIF_MDI_TO_BMP_EX = i
    
    If False Then
errout:
        Call PrintLog("TIF_MDI_TO_BMP_EX-Not Import A Current Image File!!")
        TIF_MDI_TO_BMP_EX = -1
    End If
End Function
'/***********************小華修改的(2009/04/30)***************************/




'/*   將BMP檔轉成JPG格式的函式，傳回值為轉好的JPG檔的路徑              */
'/*   Example Input BMP_TO_JPG("C:","doc99.bmp","C:","doc99.jpg")       */
Public Function BMP_TO_JPG(ByVal Source_FilePath As String, ByVal Source_FileName As String, ByVal Target_FilePath As String, ByVal Target_FileName As String) As String
    On Error GoTo errout:
    
    Dim Convert As New GflAx.GflAx
     
    '/*確保檔案路徑是正確的*/
    If Right(Source_FilePath, 1) <> "\" Then
        Source_FilePath = Source_FilePath & "\"
    End If
    If Right(Target_FilePath, 1) <> "\" Then
        Target_FilePath = Target_FilePath & "\"
    End If
    '/**/
    
    
    
    '/*讀取來源檔案*/
    Convert.LoadBitmap Source_FilePath & Source_FileName
    '/**/

    
    '/*寫入檔案到目的地*/
    Dim i As Long
    Dim Temp_FileName As String
    
    '/*另存的格式設定，這邊定死設定為JPEG*/
    Convert.SaveFormat = AX_JPEG
    Convert.SaveJPEGProgressive = True
    Convert.SaveJPEGQuality = 70
    '/**/

    Convert.SaveBitmap Target_FilePath & Target_FileName
    Convert.NextPage
    '/**/
    
    
    
    '/*將Convert丟入GC(垃圾收集器)*/
    Set Convert = Nothing
    '/**/
    
    BMP_TO_JPG = Target_FilePath & Temp_FileName
    
    If False Then
errout:
        BMP_TO_JPG = ""
        Call PrintLog("BMP_TO_JPG-Not Import A Current Image File!")
    End If
End Function
'/***********************小華修改的(2009/05/15)***************************/





'/*   用於Vista及7的電腦，將BMP檔轉成JPG格式的函式，傳回值為成功與否       */
'/*   Example Input DI_BMP_TO_JPG("C:","doc99.bmp","C:","doc99.jpg")       */
Public Function DI_BMP_TO_JPG(ByVal Source_FilePath As String, ByVal Source_FileName As String, ByVal Target_FilePath As String, ByVal Target_FileName As String) As Boolean
    On Error GoTo errout:
    
    Dim FSO As New FileSystemObject
    Dim ResultValue As Boolean
     
    '/*確保檔案路徑是正確的*/
    If Right(Source_FilePath, 1) <> "\" Then
        Source_FilePath = Source_FilePath & "\"
    End If
    If Right(Target_FilePath, 1) <> "\" Then
        Target_FilePath = Target_FilePath & "\"
    End If
    '/**/
    
    
    '/*因為來源檔案必需是在C槽下的tmp.bmp，所以把傳入的檔案移到該位置*/
    Call FSO.CopyFile(Source_FilePath & Source_FileName, "C:\tmp.bmp")
    '/**/
    
    
    '/*讀取來源檔案，並寫入檔案到目的地*/
    ResultValue = DIWriteJpg(Target_FilePath & Target_FileName, 100, 100)
    '/**/
   
    '/*清除暫存的檔案*/
    If (Source_FilePath & Source_FileName) <> "C:\tmp.bmp" Then
        Call Kill("C:\tmp.bmp")
    End If
    '/**/
    
    DI_BMP_TO_JPG = ResultValue
    
    If False Then
errout:
        DI_BMP_TO_JPG = False
        Call PrintLog("DI_BMP_TO_JPG-Not Import A Current Image File!")
    End If
End Function
'/***********************小華修改的(2009/11/10)***************************/




'/*   用於把Bmp檔轉換為標準的Bmp檔用到的函式，傳回值為成功與否              */
'/*   Example Input Convert_BMP(Me,"C:","doc99.bmp","C:","doc99.jpg")       */
Public Function Convert_BMP(ByVal frm As Form, ByVal Source_FilePath As String, ByVal Source_FileName As String, ByVal Target_FilePath As String, ByVal Target_FileName As String) As Boolean
    On Error GoTo errout:
     
     
    '/*確保檔案路徑是正確的*/
    If Right(Source_FilePath, 1) <> "\" Then
        Source_FilePath = Source_FilePath & "\"
    End If
    If Right(Target_FilePath, 1) <> "\" Then
        Target_FilePath = Target_FilePath & "\"
    End If
    '/**/
    
    
    '/*建立一個PictureBox物件，載入要被標準化的Bmp檔，再將其轉存到指定路徑的成為標準的Bmp檔*/
    Dim Convert_BMP_PictureBox As PictureBox
    Set Convert_BMP_PictureBox = frm.Controls.Add("VB.PictureBox", "Convert_BMP_PictureBox")
    Set Convert_BMP_PictureBox.Picture = LoadPicture(Source_FilePath & Source_FileName)
    Call SavePicture(Convert_BMP_PictureBox, Target_FilePath & Target_FileName)
    Call frm.Controls.Remove("Convert_BMP_PictureBox")
    '/**/
    
    Convert_BMP = True
    
    If False Then
errout:
        Convert_BMP = False
        Call PrintLog("Convert_BMP-Not Import A Current Image File!")
    End If
End Function
'/***********************小華修改的(2009/11/11)***************************/


