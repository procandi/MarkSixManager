Attribute VB_Name = "basDICOM"
'/******************************************************************/
'/*�����G�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*��m��Ҧ�DICOM�������a��C                                     */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�ҲաG�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@    �@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�ѦҡG�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@    �@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*����G    �@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*dicomobjects.ocx�C                                              */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�`�N�ƶ��G�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*                                           Edit By Edward Hsieh */
'/*                                      Last Edit Date 2009/09/30 */
'/******************************************************************/
Option Explicit



'/**/
Public DicomImg As New DicomImage
Public DicomImgs As New DicomImages
'/**/

'/**/
Public Const DEFAULT_DICOM_IP As String = "localhost" 'Dicom�n�s�u���ؼЪ�IP
Public Const DEFAULT_DICOM_PORT As String = "104" 'Dicom�n�s�u���ؼЪ�Port
'/**/


'/*�פJ�DDICOM�榡���ɮ�*/
Public Function ImportNonDicomFile(ByVal LoadAs As String, ByRef DcmImg As DicomImage, ByRef DcmImgs As DicomImages) As Boolean
    On Error GoTo errout:
    
    Call DcmImg.FileImport(LoadAs, "")
    Call DcmImgs.Add(DcmImg)

    ImportNonDicomFile = True
    
    If False Then
errout:
        ImportNonDicomFile = False
    End If
End Function
'/***********************�p�حק諸(2009/09/30)***************************/

'/*�N�DDICOM�榡���ɮסA�ץX���t�@��DICOM�榡���ɮ�*/
Public Function ExportNonDicomToDicomFile(ByVal SaveAs As String, ByRef DcmImg As DicomImage, ByRef DcmImgs As DicomImages) As Boolean
    On Error GoTo errout:
    
    DcmImgs.Clear
    Call DcmImgs.Add(DcmImg)
    Call DcmImgs.Item(DcmImgs.Count).WriteFile(SaveAs, True)
    
    ExportNonDicomToDicomFile = True
    
    If False Then
errout:
        ExportNonDicomToDicomFile = False
    End If
End Function
'/***********************�p�حק諸(2009/09/30)***************************/


'/*   �NBMP���নDCM�榡���禡�A�Ǧ^�Ȭ��`�@��F�X�iJPG��         */
'/*   Example Input BMP_TO_DCM("C:","doc99.bmp","C:","doc*.dcm")   */
Public Function BMP_TO_DCM(ByVal Source_FilePath As String, ByVal Source_FileName As String, ByVal Target_FilePath As String, ByVal Target_FileName As String) As Boolean
    On Error GoTo errout:


    Dim DcmImg As New DicomImage
    Dim DcmImgs As New DicomImages
    
    '/*�T�O�ɮ׸��|�O���T��*/
    If Right(Source_FilePath, 1) <> "\" Then
        Source_FilePath = Source_FilePath & "\"
    End If
    If Right(Target_FilePath, 1) <> "\" Then
        Target_FilePath = Target_FilePath & "\"
    End If
    '/**/
    

    '/*�T�O�ɮ׸��|�O���T��*/
    If Right(Source_FilePath, 1) <> "\" Then
        Source_FilePath = Source_FilePath & "\"
    End If
    If Right(Target_FilePath, 1) <> "\" Then
        Target_FilePath = Target_FilePath & "\"
    End If
    '/**/
    
    
    
    '/*����*/
    Call ImportNonDicomFile(Source_FilePath & Source_FileName, DcmImg, DcmImgs)
    Call ExportNonDicomToDicomFile(Target_FilePath & Target_FileName, DcmImg, DcmImgs)
    '/**/
    
    
    BMP_TO_DCM = True
    
    If False Then
errout:
        Call PrintLog("BMP_TO_DCM-Not Import A Current Image File!!")
        BMP_TO_DCM = False
    End If
End Function
'/***********************�p�حק諸(2009/09/30)***************************/

