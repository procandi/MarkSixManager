Attribute VB_Name = "basOCR"
'/******************************************************************/
'/*說明：　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*放置跟光學文字辨識有關的地方。                                  */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*模組：　　　　　　　　　　　　　　　    　　　　　　　　　　　　*/
'/*無。                                                            */
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


'/**************************公用變數的部份***********************************/
Public OCR_Word As String '公用來放置讀取出來的光學文字用的變數
Public OCR_Length As String '公用來放置讀取出來的光學文字長度用的變數
Public OCR_Supply As Boolean '公用來放置讀取出來的光學文字決定是否要補置足碼
'/**************************小華修改的(2009/04/09)***********************************/



'/***************讀取有支援光學文字辨識的檔案的光學文字*********************/
Public Function ReadOCR(ByVal FilePath As String, ByVal FileName As String) As String()
    On Error GoTo errout:
    
    Dim i As Long
    Dim miDoc As New MODI.Document
    Dim miImg As MODI.Image
    Dim nWordCount As Long
    Dim ResultString() As String
    
    
    If Right(FilePath, 1) <> "\" Then
        FilePath = FilePath & "\"
    End If
    
    Call miDoc.Create(FilePath & FileName)
    ReDim ResultString(miDoc.Images.Count - 1)
    
    For i = 0 To miDoc.Images.Count - 1
        Set miImg = miDoc.Images(i)
    
        nWordCount = miImg.Layout.NumChars
    
        If nWordCount > -1 Then
            ResultString(i) = miImg.Layout.Text
        Else
            ResultString(i) = ""
        End If
    Next
    
    Set miImg = Nothing
    Call miDoc.Close(False)
    Set miDoc = Nothing
    
    ReadOCR = ResultString
    
    If False Then
errout:
        Call PrintLog("ReadOCR-Load Image OCR is Faild,Image Path=" & FilePath & FileName & "!")
        ReDim ResultString(0)
        ResultString(0) = "Error"
        ReadOCR = ResultString
    End If
End Function
'/**************************小華修改的(2009/04/09)***********************************/





'/***************讀取有支援光學文字辨識的檔案的光學文字*********************/
Public Function InputOCR(ByVal FilePath As String, ByVal FileName As String, ByVal PageNum As Long) As String()
    On Error GoTo errout:
    
    Dim i As Long
    Dim j As Long
    Dim miDoc As New MODI.Document
    Dim miImg As MODI.Image
    Dim miWord As MODI.Word
    Dim nWordCount As Long
    Dim ResultString() As String
    
    
    If Right(FilePath, 1) <> "\" Then
        FilePath = FilePath & "\"
    End If
    
    Call miDoc.Create(FilePath & FileName)
    Set miImg = miDoc.Images(PageNum)
    
    ReDim ResultString(miImg.Layout.Words.Count - 1)
    For i = 0 To miImg.Layout.Words.Count - 1
        Set miWord = miImg.Layout.Words(i)
        ResultString(i) = miWord.Text
    Next
    
    
    Set miImg = Nothing
    Call miDoc.Close(False)
    Set miDoc = Nothing
    
    InputOCR = ResultString
    
    If False Then
errout:
        Call PrintLog("InputOCR-Load Image OCR is Faild,Image Path=" & FilePath & FileName & "!")
        ReDim ResultString(0)
        ResultString(0) = "Error"
        InputOCR = ResultString
    End If
End Function
'/**************************小華修改的(2009/04/09)***********************************/

