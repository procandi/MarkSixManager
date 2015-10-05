Attribute VB_Name = "basError"
'/******************************************************************/
'/*�����G�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*��m�B�z�o�ͨҥ~���p�ɥΨ쪺�禡�B�ܼơB�`�Ƶ����a��C          */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�ҲաG�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@    �@�@�@�@�@�@�@�@�@�@�@�@*/
'/*basFile.bas��basVariable.bas�C                                  */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�ѦҡG�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*scrrun.dll�C�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@  */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*����G�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�`�N�ƶ��G�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*                                           Edit By Edward Hsieh */
'/*                                      Last Edit Date 2009/03/13 */
'/******************************************************************/
Option Explicit



'/**************************�L�XLog�ɦܫ��w�����|�A�Y���|���s�b�h�|�۰ʻs�@***************************/
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
        Print #FreeFilePort, SaveDate, SaveTime, "����-" & Log_String
        Print #FreeFilePort, SaveDate, SaveTime, "�N�X-" & err.Number
        Print #FreeFilePort, SaveDate, SaveTime, "�T��-" & err.Description
    Close #FreeFilePort
    
    PrintLog = SavePath & SaveFile
End Function
'/**************************�p�حק諸(2009/03/13)***********************************/



'/**************************��ܿ��~�T���A�ðO����Log�ɤ��A�H�K�ư����~***********************************/
Public Function ErrorOut(ByVal Log_String As String, Optional ByVal FileName As String = "") As Boolean
    Dim Log_File As String
    
    Log_File = PrintLog(Log_String, FileName)
    MsgBox "�o�ͨt�ο��~�G" & vbCrLf & "����-" & Log_String & vbCrLf & "�N�X-" & err.Number & vbCrLf & "�T��-" & err.Description & vbCrLf & vbCrLf & "�бN���~�O����" & Log_File & "�q���ñH���a�Q�t�Τu�{�v", vbCritical
    
    ErrorOut = True
End Function
'/**************************�p�حק諸(2009/03/13)***********************************/
