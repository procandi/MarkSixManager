Attribute VB_Name = "basFile"
'/******************************************************************/
'/*�����G�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*��m���ɮ׳B�z�������禡�C                                      */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�ҲաG�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@    �@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�ѦҡG    �@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*scrrun.dll�C                                                    */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*����G    �@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�`�N�ƶ��G�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*                                           Edit By Edward Hsieh */
'/*                                      Last Edit Date 2009/05/13 */
'/******************************************************************/
Option Explicit



'/**************************���ɮ׳B�z�������ܼ�***********************************/
Public BackupPath As String '�Ψө�m�t�η|�Ψ쪺�ƥ����|��m
'/***************************�p�حק諸(2009/05/13)***************************/


'/**************************���ɮ׳B�z�������`��***********************************/
Public Const MAX_FOLDER_DEPTH As Integer = 255 '�̤j�i�H�������t�θ��|�`�סA�H�קK�b�Q�ΥH�U�o�Ǩ禡�ɡA�إ߹L�`����Ƨ����|
Public Const BMP_CLASS As String = "bmp"
Public Const JPG_CLASS As String = "jpg"
Public Const DOT_BMP_CLASS As String = ".bmp"
Public Const DOT_JPG_CLASS As String = ".jpg"
'/***************************�p�حק諸(2009/05/13)***************************/

'/***************************���ɮ׳B�z�������C�|***************************/
Public Enum FileDateClass
    FileDateCreated = 0
    FileDateLastAccessed = 1
    FileDateLastModified = 2
End Enum
'/***************************�p�حק諸(2009/11/09)***************************/


'/**/
Public FSO As New FileSystemObject
'/**/


'/*******************�����ɮת��ɦW�A�|��ǤJ�Ȥ��Φ��ɦW����ɦW�A�h�l�������A��p���O���|�������|�Q�h���A�Ǧ^�ȯu�N���\�A���N����***************************/
Public Function DivisionFileName(ByVal Input_File_Name As String, ByRef Output_File_Name As String, ByRef Output_File_Class As String) As Boolean
    Dim FSO_FileExist As New FileSystemObject
    
    Output_File_Name = ""
    Output_File_Class = ""


    Dim i As Long
    
    For i = Len(Input_File_Name) To 1 Step -1
        If Mid(Input_File_Name, i, 1) = "." Then
            Output_File_Name = Left(Input_File_Name, i - 1)
            Output_File_Class = Right(Input_File_Name, Len(Input_File_Name) - i)
            Exit For
        End If
    Next
    For i = Len(Output_File_Name) To 1 Step -1
        If Mid(Output_File_Name, i, 1) = "\" Then
            Output_File_Name = Right(Output_File_Name, Len(Output_File_Name) - i)
            Exit For
        End If
    Next
    
    If Output_File_Name = "" Or Output_File_Class = "" Then
        DivisionFileName = False
    Else
        DivisionFileName = True
    End If
End Function
'/***************************�p�حק諸(2009/05/13)***************************/



'/***************************�����ɮת����|�A�|��ǤJ�Ȥ��Φ����|���ɮצW�١A�Ǧ^�Ȭ��u�N���\�A���N����***************************/
Public Function DivisionFilePath(ByVal Input_File_Path As String, ByRef Output_File_Path As String, ByRef Output_File_Name As String) As Boolean
    Dim FSO_FileExist As New FileSystemObject
    
    Output_File_Path = ""
    Output_File_Name = ""
    If FSO_FileExist.FileExists(Input_File_Path) Then
        Dim i As Long
        
        For i = Len(Input_File_Path) To 1 Step -1
            If Mid(Input_File_Path, i, 1) = "\" Then
                Output_File_Path = Left(Input_File_Path, i - 1)
                Output_File_Name = Right(Input_File_Path, Len(Input_File_Path) - i)
                Exit For
            End If
        Next
        
        If Output_File_Path = "" Or Output_File_Name = "" Then
            DivisionFilePath = False
        Else
            DivisionFilePath = True
        End If
    Else
        DivisionFilePath = False
    End If
End Function
'/***************************�p�حק諸(2009/05/13)***************************/



'/***************************�إ߶Ƕi�Ӫ��ɮ׸��|�A�i�H�䴩�إߦh�h����Ƨ��ҳ��٨S�إߪ����p***************************/
Public Function CreatePath(ByVal Path As String) As Boolean
    Dim FSO_CreatePath As New FileSystemObject
    Dim i As Integer
    Dim Count As Integer
    Dim TempPath As String
    Dim AllPath(MAX_FOLDER_DEPTH) As String
    

    
    Count = 0
    Do Until FSO_CreatePath.GetBaseName(Path) = ""
        AllPath(Count) = Path
        Path = FSO_CreatePath.GetParentFolderName(Path)
        Count = Count + 1
    Loop
    
    If FSO_CreatePath.DriveExists(Path) Then
        For i = Count - 1 To 0 Step -1
            If Not FSO_CreatePath.FolderExists(AllPath(i)) Then
                Call FSO_CreatePath.CreateFolder(AllPath(i))
            End If
        Next
        
        CreatePath = True
    Else
        CreatePath = False
    End If
End Function
'/***************************�p�حק諸(2009/05/13)***************************/




'/***************************���o�ɮת��إ߮ɶ��B�̫�s���ɶ��B�̫�ק�ɶ�***************************/
Public Function GetFileTime(ByVal FilePath As String, ByVal FileClass As FileDateClass) As Date
    Dim FSO As New FileSystemObject
    Dim objFile As Scripting.File
    
    
    Set objFile = FSO.GetFile(FilePath)
    
    
    Select Case FileClass
    Case FileDateClass.FileDateCreated
        GetFileTime = objFile.DateCreated
    Case FileDateClass.FileDateLastAccessed
        GetFileTime = objFile.DateLastAccessed
    Case FileDateClass.FileDateLastModified
        GetFileTime = objFile.DateLastModified
    Case Else
        GetFileTime = DateTime.Now
    End Select
End Function
'/***************************�p�حק諸(2009/11/09)***************************/



'/*�q�ɮפ����ɮ�Ū�X�ӭק�j�p��A�s�^�h*/
Function SetJpegFileResize(ByVal frm As Form, ByVal SourcePath As String, ByVal SourceName As String, ByVal TargetPath As String, ByVal TargetName As String, ByVal TargetX As Double, ByVal TargetY As Double, ByVal TargetWidth As Double, ByVal TargetHeight As Double) As Boolean
    On Error GoTo errout:
    
    Dim PICName As String
    Dim PicClass As String
    Dim PicBoxA As PictureBox
    Dim PicBoxB As PictureBox
    Dim FSO As New FileSystemObject
    
    
    '/*�T�O�ɮ׸��|�O���T��*/
    If Right(SourcePath, 1) <> "\" Then
        SourcePath = SourcePath & "\"
    End If
    If Right(TargetPath, 1) <> "\" Then
        TargetPath = TargetPath & "\"
    End If
    '/**/
    
    '/**/
    Set PicBoxA = frm.Controls.Add("VB.PictureBox", "PicBoxA")
    Set PicBoxB = frm.Controls.Add("VB.PictureBox", "PicBoxB")
    '/**/
    
    '/**/
    PicBoxA.ScaleMode = 3
    PicBoxA.width = TargetWidth * Screen.TwipsPerPixelX
    PicBoxA.height = TargetHeight * Screen.TwipsPerPixelY
    PicBoxA.AutoRedraw = True
    PicBoxB.ScaleMode = 3
    PicBoxB.width = TargetWidth * Screen.TwipsPerPixelX
    PicBoxB.height = TargetHeight * Screen.TwipsPerPixelY
    PicBoxB.AutoRedraw = True
    '/**/
    
    '/**/
    Call DivisionFileName(TargetName, PICName, PicClass)
    PicBoxA.Picture = LoadPicture(SourcePath & SourceName)
    Call PicBoxB.PaintPicture(PicBoxA.Picture, TargetX, TargetY, TargetWidth, TargetHeight)
    Call SavePicture(PicBoxB.Image, TargetPath & PICName & DOT_BMP_CLASS)
    '/**/

    '/**/
    Call frm.Controls.Remove("PicBoxA")
    Call frm.Controls.Remove("PicBoxB")
    '/**/
    
    '/**/
    Call DI_BMP_TO_JPG(TargetPath, PICName & DOT_BMP_CLASS, TargetPath, PICName & DOT_JPG_CLASS)
    Call FSO.DeleteFile(TargetPath & PICName & DOT_BMP_CLASS)
    '/**/
    
    SetJpegFileResize = True
    
    If False Then
errout:
        SetJpegFileResize = False
        Call PrintLog("SetJpegFileResize-Not Import A Current Image File!")
    End If
End Function
'/*2010/01/04*/




'/*�ק�t��Ū��./�����|���{�����檺���|*/
Public Function ChangeSystemPath2ApplicationPath() As Boolean
    On Error GoTo errout:
    
    Dim TargetPath As String
    
    
    '/*�T�O�ɮ׸��|�O���T��*/
    TargetPath = App.Path
    If Right(TargetPath, 1) <> "\" Then
        TargetPath = TargetPath & "\"
    End If
    '/**/
    
    Call FileSystem.ChDrive(FSO.GetDriveName(TargetPath))
    Call FileSystem.ChDir(FSO.GetAbsolutePathName(TargetPath))

    
    ChangeSystemPath2ApplicationPath = True
    
    If False Then
errout:
        ChangeSystemPath2ApplicationPath = False
        Call PrintLog("ChangeSystemPath2ApplicationPath-Can't Change Path!")
    End If
End Function
'/*20100521*/



'/*�ק�t��Ū��./�����|�����w�����|*/
Public Function ChangeSystemPath2DestinationPath(ByVal TargetPath As String) As Boolean
    On Error GoTo errout:
      
    
    '/*�T�O�ɮ׸��|�O���T��*/
    If Right(TargetPath, 1) <> "\" Then
        TargetPath = TargetPath & "\"
    End If
    '/**/
    
    Call FileSystem.ChDrive(FSO.GetDriveName(TargetPath))
    Call FileSystem.ChDir(FSO.GetAbsolutePathName(TargetPath))

    
    ChangeSystemPath2DestinationPath = True
    
    If False Then
errout:
        ChangeSystemPath2DestinationPath = False
        Call PrintLog("ChangeSystemPath2DestinationPath-Can't Change Path!")
    End If
End Function
'/*20100521*/



