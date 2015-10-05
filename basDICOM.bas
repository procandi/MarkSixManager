Attribute VB_Name = "basDICOM"
'/******************************************************************/
'/*說明：　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*放置跟所有DICOM有關的地方。                                     */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*模組：　　　　　　　　　　　　　　　    　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*參考：　　　　　　　　　　　　　　　    　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*元件：    　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*dicomobjects.ocx。                                              */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*注意事項：　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*                                           Edit By Edward Hsieh */
'/*                                      Last Edit Date 2009/09/30 */
'/******************************************************************/
Option Explicit



'/**/
Public DicomImg As New DicomImage
Public DicomImgs As New DicomImages
'/**/

'/**/
Public Const DEFAULT_DICOM_IP As String = "localhost" 'Dicom要連線的目標的IP
Public Const DEFAULT_DICOM_PORT As String = "104" 'Dicom要連線的目標的Port
'/**/


'/*匯入非DICOM格式的檔案*/
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
'/***********************小華修改的(2009/09/30)***************************/

'/*將非DICOM格式的檔案，匯出成另一個DICOM格式的檔案*/
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
'/***********************小華修改的(2009/09/30)***************************/


'/*   將BMP檔轉成DCM格式的函式，傳回值為總共轉了幾張JPG檔         */
'/*   Example Input BMP_TO_DCM("C:","doc99.bmp","C:","doc*.dcm")   */
Public Function BMP_TO_DCM(ByVal Source_FilePath As String, ByVal Source_FileName As String, ByVal Target_FilePath As String, ByVal Target_FileName As String) As Boolean
    On Error GoTo errout:


    Dim DcmImg As New DicomImage
    Dim DcmImgs As New DicomImages
    
    '/*確保檔案路徑是正確的*/
    If Right(Source_FilePath, 1) <> "\" Then
        Source_FilePath = Source_FilePath & "\"
    End If
    If Right(Target_FilePath, 1) <> "\" Then
        Target_FilePath = Target_FilePath & "\"
    End If
    '/**/
    

    '/*確保檔案路徑是正確的*/
    If Right(Source_FilePath, 1) <> "\" Then
        Source_FilePath = Source_FilePath & "\"
    End If
    If Right(Target_FilePath, 1) <> "\" Then
        Target_FilePath = Target_FilePath & "\"
    End If
    '/**/
    
    
    
    '/*轉檔*/
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
'/***********************小華修改的(2009/09/30)***************************/

