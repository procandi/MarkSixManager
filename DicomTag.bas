Attribute VB_Name = "BasDicomTag"
'/定義DicomTag，值為16進位之數值

'/~~~~~此區為Dicom必要值~~~~~~/

'一、<Patient Level↓>
'/SOP Instance UID
'/Public Const TagSOPUID As String = "0008,0010"
'/Patient Name
Public Const TagPatientName As String = "0010,0010"
'/Issuer of Patient ID 病例號發證院所
Public Const TagIssuerID As String = "0010,0021"
'/~Other Patient IDs 病例號
Public Const TagPatientID As String = "0010,0020"
'/~Patient’s Sex
Public Const TagPatientSex As String = "0010,0040"
'/Patient’s Birth Date
Public Const TagPatientBirthday As String = "0010,0030"
'/~Other Patient IDs 身分證
Public Const TagOtherID As String = "0010,1000"
'/~Patient Age  病人年齡
Public Const TagAge As String = "0010,1010"
'<Patient Level↑>

'二、<Study Level↓>
'/~Study Instance UID
Public Const TagStudyUID As String = "0020,000D"
'/~Study Description 檢查描述
Public Const TagDescription As String = "0008,1030"
'~Accession Number 申請序號
Public Const TagAccession As String = "0008,0050"
'/Study Time 影像擷取時間
Public Const TagStudyTime As String = "0008,0030"
'/~Study Date 影像擷取日期
Public Const TagStudyDate As String = "0008,0020"
'/Modalities in Study 診斷儀器種類
Public Const TagModality As String = "0008,0060"
'<Study Level↑>

'三、<Series Level ( 00041430 directory record type : SERIES)↓>

'/~Series Instance UID
Public Const TagSeriesUID As String = "0020,000E"
'/~Body Part Examined 身體受檢部位
Public Const TagExamBody As String = "0018,0015"
'/~Series Description 序列說明
Public Const TagSerDecription As String = "0018,0015"
'/~Series Number 序列號碼
Public Const TagSerNum As String = "0020,0011"
'/~Laterality 側向性
Public Const TagLaterality As String = "0020,0060"

'<Series Level ( 00041430 directory record type : SERIES↑>

'四、<Instance Level ( 00041430 directory record type : IMAGE)↓>
'/~SOP Instance UID
Public Const TagUID As String = "0008,0018"
'/~Acquisition Time 擷取時間
Public Const TagAcquisitionTime As String = "0008,0032"
'/~Content Time 內容時間
Public Const TagContentTime As String = "0008,0033"
'/~Instance Number 影像號碼
Public Const TagInsNum As String = "0020,0013"
'/Window Width 視窗寬度
Public Const TagWinWidth As String = "0028,1051"
'/Window Center 視窗中心
Public Const TagWinCenter As String = "0028,1050"
'/~Number of Frames 訊框數
Public Const TagNoF As String = "0028,0008"
'/Rows 列
Public Const TagRow As String = "0028,0010"
'/Columns 行
Public Const TagColumns As String = "0028,0011"
'/~ImageNumber
Public Const TagImageNumber As String = "0020,0013"
'/參考醫生 Referring Physician's Name
Public Const TagReferenceDoctor As String = "0008,0090"

'<Instance Level ( 00041430 directory record type : IMAGE)↑>
Public Sub GetDicomTag(ByVal DicomTag As String, ByRef Group As String, ByRef Elem As String)
    Group = Left(DicomTag, 4)
    Elem = Right(DicomTag, 4)
End Sub

Public Function GetDicomTagValue(ByVal Tag As String, ByVal DcmAtt As Object) As String
    Dim v As Variant
    Dim atts As DicomAttribute
    Group = Convert16to10(Left(Tag, 4))
    Elem = Convert16to10(Right(Tag, 4))
    
    For Each atts In DcmAtt
        If atts.Group = Group And atts.Element = Elem Then
            v = atts.Value
            Exit For
        End If
    Next
    GetDicomTagValue = v
End Function


'寫入Dicom 必須為10進制
Public Sub WriteDicomTag(ByRef Image As DicomImage, ByVal Tag As String, ByVal TagValue As String)
    On Error GoTo errWriteTag
    Dim Group As String
    Dim Element As String
    Call GetDicomTag(Tag, Group, Element)
    
    '/寫入TAG
    
    'SaveLog (Group & " + " & Element & "　　　" & Convert16to10(Group) & "," & Convert16to10(Element) & "====> " & TagValue)
    Image.Attributes.Add Convert16to10(Group), Convert16to10(Element), TagValue
    Exit Sub
    
errWriteTag:
    PrintLog (err.Number & ":" & err.Description)
End Sub

'/轉換為10進制
Public Function Convert16to10(ByVal Va As String) As Long
    For i = 1 To Len(Va)
        Select Case Mid(Va, i, 1)
            Case 0 To 9
                Sum = Sum + Mid(Va, i, 1) * (16 ^ (4 - i))
            Case "A"
                Sum = Sum + 10 * (16 ^ (4 - i))
            Case "B"
                Sum = Sum + 11 * (16 ^ (4 - i))
            Case "C"
                Sum = Sum + 12 * (16 ^ (4 - i))
            Case "D"
                Sum = Sum + 13 * (16 ^ (4 - i))
            Case "E"
                Sum = Sum + 14 * (16 ^ (4 - i))
            Case "F"
                Sum = Sum + 15 * (16 ^ (4 - i))
        End Select
    Next
    Convert16to10 = Sum
End Function


