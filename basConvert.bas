Attribute VB_Name = "basConvert"
'/******************************************************************/
'/*�����G�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*��m��Ҧ��榡�ഫ�������a��C                                  */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�ҲաG�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@    �@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�ѦҡG�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@    �@�@�@�@�@�@�@�@�@�@�@�@*/
'/*GflAx.dll��MDIVWCTL.DLL��DIjpg.dll�C                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*����G    �@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�`�N�ƶ��G�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*                                           Edit By Edward Hsieh */
'/*                                      Last Edit Date 2009/04/09 */
'/******************************************************************/
Option Explicit


'/*                       �@�����ɷ|�Ψ쪺�禡�w                       */
Private Declare Function DIWriteJpg Lib "DIjpg.dll" (ByVal DestPath As String, ByVal quality As Long, ByVal progressive As Long) As Long
'/***********************�p�حק諸(2009/11/10)***************************/


'/*   �NTIF���নJPG�榡���禡�A�Ǧ^�Ȭ��`�@��F�X�iJPG��              */
'/*   Example Input TIF_TO_JPG("C:","doc99.tif","C:","doc*.jpg")       */
Public Function TIF_TO_JPG(ByVal Source_FilePath As String, ByVal Source_FileName As String, ByVal Target_FilePath As String, ByVal Target_FileName As String) As Long
    On Error GoTo errout:
    
    Dim Convert As New GflAx.GflAx
     
    '/*�T�O�ɮ׸��|�O���T��*/
    If Right(Source_FilePath, 1) <> "\" Then
        Source_FilePath = Source_FilePath & "\"
    End If
    If Right(Target_FilePath, 1) <> "\" Then
        Target_FilePath = Target_FilePath & "\"
    End If
    '/**/
    
    
    
    '/*Ū���ӷ��ɮ�*/
    Convert.LoadBitmap Source_FilePath & Source_FileName
    '/**/

    
    '/*�g�J�ɮר�ت��a*/
    Dim i As Long
    Dim Temp_FileName As String
    
    For i = 0 To Convert.NumberOfPages - 1
        '/*�t�s���榡�]�w�A�o��w���]�w��JPEG*/
        Convert.SaveFormat = AX_JPEG
        Convert.SaveJPEGProgressive = True
        Convert.SaveJPEGQuality = 70
        '/**/
    
    
        Temp_FileName = Replace(Target_FileName, "*", i)
        Convert.SaveBitmap Target_FilePath & Temp_FileName
        Convert.NextPage
    Next
    '/**/
    
    
    
    '/*�NConvert��JGC(�U��������)*/
    Set Convert = Nothing
    '/**/
    
    TIF_TO_JPG = i
    
    If False Then
errout:
        TIF_TO_JPG = -1
        Call PrintLog("TIF_TO_JPG-Not Import A Current Image File!")
    End If
End Function
'/***********************�p�حק諸(2009/04/09)***************************/




'/*   �NTIF��MDI���নJPG�榡���禡�A�Ǧ^�Ȭ��`�@��F�X�iJPG��         */
'/*   Example Input TIF_MDI_TO_BMP("C:","doc99.tif","C:","doc*.jpg")   */
Public Function TIF_MDI_TO_BMP(ByVal Source_FilePath As String, ByVal Source_FileName As String, ByVal Target_FilePath As String, ByVal Target_FileName As String) As Long
    On Error GoTo errout:


    Dim miDoc As New MODI.Document
    Dim miImg As MODI.Image
    
    
    '/*�T�O�ɮ׸��|�O���T��*/
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
'/***********************�p�حק諸(2009/04/29)***************************/




'/*   �NTIF��MDI���নJPG�榡���禡�A�Ǧ^�Ȭ��`�@��F�X�iJPG��               */
'/*   Example Input TIF_MDI_TO_BMP_EX("C:","doc99.tif","C:","doc*.jpg",90)   */
Public Function TIF_MDI_TO_BMP_EX(ByVal Source_FilePath As String, ByVal Source_FileName As String, ByVal Target_FilePath As String, ByVal Target_FileName As String, ByVal Rotate As Integer) As Long
    On Error GoTo errout:


    Dim miDoc As New MODI.Document
    Dim miImg As MODI.Image

    
    '/*�T�O�ɮ׸��|�O���T��*/
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
'/***********************�p�حק諸(2009/04/30)***************************/




'/*   �NBMP���নJPG�榡���禡�A�Ǧ^�Ȭ���n��JPG�ɪ����|              */
'/*   Example Input BMP_TO_JPG("C:","doc99.bmp","C:","doc99.jpg")       */
Public Function BMP_TO_JPG(ByVal Source_FilePath As String, ByVal Source_FileName As String, ByVal Target_FilePath As String, ByVal Target_FileName As String) As String
    On Error GoTo errout:
    
    Dim Convert As New GflAx.GflAx
     
    '/*�T�O�ɮ׸��|�O���T��*/
    If Right(Source_FilePath, 1) <> "\" Then
        Source_FilePath = Source_FilePath & "\"
    End If
    If Right(Target_FilePath, 1) <> "\" Then
        Target_FilePath = Target_FilePath & "\"
    End If
    '/**/
    
    
    
    '/*Ū���ӷ��ɮ�*/
    Convert.LoadBitmap Source_FilePath & Source_FileName
    '/**/

    
    '/*�g�J�ɮר�ت��a*/
    Dim i As Long
    Dim Temp_FileName As String
    
    '/*�t�s���榡�]�w�A�o��w���]�w��JPEG*/
    Convert.SaveFormat = AX_JPEG
    Convert.SaveJPEGProgressive = True
    Convert.SaveJPEGQuality = 70
    '/**/

    Convert.SaveBitmap Target_FilePath & Target_FileName
    Convert.NextPage
    '/**/
    
    
    
    '/*�NConvert��JGC(�U��������)*/
    Set Convert = Nothing
    '/**/
    
    BMP_TO_JPG = Target_FilePath & Temp_FileName
    
    If False Then
errout:
        BMP_TO_JPG = ""
        Call PrintLog("BMP_TO_JPG-Not Import A Current Image File!")
    End If
End Function
'/***********************�p�حק諸(2009/05/15)***************************/





'/*   �Ω�Vista��7���q���A�NBMP���নJPG�榡���禡�A�Ǧ^�Ȭ����\�P�_       */
'/*   Example Input DI_BMP_TO_JPG("C:","doc99.bmp","C:","doc99.jpg")       */
Public Function DI_BMP_TO_JPG(ByVal Source_FilePath As String, ByVal Source_FileName As String, ByVal Target_FilePath As String, ByVal Target_FileName As String) As Boolean
    On Error GoTo errout:
    
    Dim FSO As New FileSystemObject
    Dim ResultValue As Boolean
     
    '/*�T�O�ɮ׸��|�O���T��*/
    If Right(Source_FilePath, 1) <> "\" Then
        Source_FilePath = Source_FilePath & "\"
    End If
    If Right(Target_FilePath, 1) <> "\" Then
        Target_FilePath = Target_FilePath & "\"
    End If
    '/**/
    
    
    '/*�]���ӷ��ɮץ��ݬO�bC�ѤU��tmp.bmp�A�ҥH��ǤJ���ɮײ���Ӧ�m*/
    Call FSO.CopyFile(Source_FilePath & Source_FileName, "C:\tmp.bmp")
    '/**/
    
    
    '/*Ū���ӷ��ɮסA�üg�J�ɮר�ت��a*/
    ResultValue = DIWriteJpg(Target_FilePath & Target_FileName, 100, 100)
    '/**/
   
    '/*�M���Ȧs���ɮ�*/
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
'/***********************�p�حק諸(2009/11/10)***************************/




'/*   �Ω��Bmp���ഫ���зǪ�Bmp�ɥΨ쪺�禡�A�Ǧ^�Ȭ����\�P�_              */
'/*   Example Input Convert_BMP(Me,"C:","doc99.bmp","C:","doc99.jpg")       */
Public Function Convert_BMP(ByVal frm As Form, ByVal Source_FilePath As String, ByVal Source_FileName As String, ByVal Target_FilePath As String, ByVal Target_FileName As String) As Boolean
    On Error GoTo errout:
     
     
    '/*�T�O�ɮ׸��|�O���T��*/
    If Right(Source_FilePath, 1) <> "\" Then
        Source_FilePath = Source_FilePath & "\"
    End If
    If Right(Target_FilePath, 1) <> "\" Then
        Target_FilePath = Target_FilePath & "\"
    End If
    '/**/
    
    
    '/*�إߤ@��PictureBox����A���J�n�Q�зǤƪ�Bmp�ɡA�A�N����s����w���|�������зǪ�Bmp��*/
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
'/***********************�p�حק諸(2009/11/11)***************************/


