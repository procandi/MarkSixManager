Attribute VB_Name = "basString"
'/******************************************************************/
'/*說明：　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*放置跟字串相關的變數、常數、函式等的地方。                      */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*模組：　　　　　　　　　　　　　　　    　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*參考：    　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*元件：    　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*注意事項：　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*                                           Edit By Edward Hsieh */
'/*                                      Last Edit Date 2009/03/03 */
'/******************************************************************/
Option Explicit


'/**************************字串相關的常數資料***********************************/
Public Const ASCII_UPPER_A As Integer = 65
Public Const ASCII_LOWER_A As Integer = 97
Public Const ASCII_NULL As Integer = 0
Public Const ASCII_MAX As Integer = 128

Public Const CHAR_EN_MAX As Integer = 26
'/**************************小華修改的(20100222)***********************************/


'/**************************字串相關的函數資料***********************************/
Public Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
'/**************************小華修改的(20100604)***********************************/




'/**************************找出字串Source中，字元為Start到End之間的字元分類出來代入Target，其餘丟到Else，最後回傳Target共有幾個字元(例如Start輸入"0"、End輸入"9"，則"0"~"9"會被放入Target，其餘在Else)***********************************/
Public Function Str_Classify(ByRef Str_Source As String, ByRef Str_Target As String, ByRef Str_Else As String, ByRef Str_Start As Long, ByRef Str_End As Long) As Long
    Dim i As Integer
    Dim Count As Long
    
    Str_Target = ""
    Str_Else = ""
    Count = 0

    For i = 1 To Len(Str_Source)
        If Asc(Mid(Str_Source, i, 1)) >= Str_Start And Asc(Mid(Str_Source, i, 1)) <= Str_End Then
            Str_Target = Str_Target & Mid(Str_Source, i, 1)
            Count = Count + 1
        Else
            Str_Else = Str_Else & Mid(Str_Source, i, 1)
        End If
    Next
    
    Str_Classify = Count
End Function
'/**************************小華修改的(2009/03/03)***********************************/


'/**************************用於將字串補位的函式，Insert_Count為要補成幾位？負數代表要把Fill_Source_Char補在Fill_Target_String前面，正數則為後面。(例：三個值由左至右分別是-4,0、99，則傳回值會為0099)***********************************/
Public Function Ch_Fill(ByVal Insert_Count As Long, ByVal Fill_Source_Char As String, ByVal Fill_Target_String As String) As String
    Dim i As Long
    Dim Init_Count As Long
    
    
    Init_Count = Len(Fill_Target_String)
    
    If Insert_Count > 0 Then
        For i = Init_Count To Insert_Count - 1
            Fill_Target_String = Fill_Target_String & Fill_Source_Char
        Next
    ElseIf Insert_Count < 0 Then
        For i = Insert_Count + 1 To -Init_Count
            Fill_Target_String = Fill_Source_Char & Fill_Target_String
        Next
    End If
    
    Ch_Fill = Fill_Target_String
End Function
'/**************************小華修改的(2009/03/16)***********************************/



'/**************************用於將字串補位的函式，Insert_Count為要補成幾位？負數代表要把Fill_Source_String補在Fill_Target_String前面，正數則為後面。(例：三個值由左至右分別是-5,12、99，則傳回值會為121299)***********************************/
Public Function Str_Fill(ByVal Insert_Count As Long, ByVal Fill_Source_String As String, ByVal Fill_Target_String As String) As String
    Dim i As Long
    Dim Init_Count As Long
    Dim StringLen As Long
    
    
    Init_Count = Len(Fill_Target_String)
    StringLen = Len(Fill_Source_String)
    
    If Insert_Count > 0 Then
        For i = Init_Count To Insert_Count - 1 Step StringLen
            Fill_Target_String = Fill_Target_String & Fill_Source_String
        Next
    ElseIf Insert_Count < 0 Then
        For i = Insert_Count + 1 To -Init_Count Step StringLen
            Fill_Target_String = Fill_Source_String & Fill_Target_String
        Next
    End If
    
    Str_Fill = Fill_Target_String
End Function
'/**************************小華修改的(2009/03/16)***********************************/




'/*****************找出在Source_String中，共有幾個Search_String，傳回數量**************/
Public Function Str_SearchCount(ByVal Source_String As String, ByVal Search_Char As String) As Long
    Dim Search_Count As Long
    Dim Search_Start As Long
    
    Search_Count = 0
    Search_Start = 0
    Do
        Search_Start = InStr(Search_Start + 1, Source_String, Search_Char)
        Search_Count = Search_Count + 1
    Loop Until Search_Start = 0
    
    Str_SearchCount = Search_Count - 1
End Function
'/**************************小華修改的(2009/03/24)***********************************/




'/*****************計算含有中英文字或有UniCode、BIG5Code混合的字串的真實字元數*******************/
Public Function Str_TrueLen(ByVal Source_String As String) As Long
    Dim SingleChar As String
    Dim DoubleChar As String
    
    Call Str_Classify(Source_String, SingleChar, DoubleChar, 0, 128)
    Str_TrueLen = Len(SingleChar) + (LenB(DoubleChar) / 2)
End Function
'/**************************小華修改的(2009/05/15)***********************************/



'/*把英文字依二十六進位轉換成數字*/
Public Function CharEN2Numeric(ByVal CharEN As String) As Long
    On Error GoTo err:
    
    
    Dim i As Long
    Dim j As Long
    Dim Answer As Long
    
    
    CharEN = UCase(CharEN)
    Answer = 0
    j = Len(CharEN) - 1
    For i = 0 To Len(CharEN) - 1
        Answer = Answer + ((1 + Asc(Mid(CharEN, i + 1, 1)) - ASCII_UPPER_A) * (CHAR_EN_MAX ^ j))
        j = j - 1
    Next
    
    CharEN2Numeric = Answer
    
    If False Then
err:
        CharEN2Numeric = -1
    End If
End Function
'/*20100122*/


'/*把數字依二十六進位轉換成英文字*/
Public Function Numeric2CharEN(ByVal Numeric As Long) As String
    On Error GoTo err:
    
    Dim Answer As String
        
    Answer = ""
    Do While Numeric > 0
        Numeric = Numeric - 1
        Answer = Chr((Numeric Mod CHAR_EN_MAX) + ASCII_UPPER_A) & Answer
        Numeric = Numeric \ CHAR_EN_MAX
    Loop
    
    Numeric2CharEN = Answer
    
    If False Then
err:
        Numeric2CharEN = "error"
    End If
End Function
'/*20100225*/



'/*傳回從文章中找到的關鍵字後的所有字串*/
Public Function ReturnKeyWordValue(ByVal Article As String, ByVal keyWord As String) As String
    On Error GoTo err:
    
    Dim WhereKeyWord As Long
    
    WhereKeyWord = InStr(Article, keyWord)
    If WhereKeyWord > 0 Then
        Dim LenArticle As Long
        Dim LenKeyWord As Long
        
        LenArticle = Len(Article)
        LenKeyWord = Len(keyWord)
    
        ReturnKeyWordValue = Right(Article, LenArticle - WhereKeyWord - LenKeyWord + 1)
    Else
        ReturnKeyWordValue = ""
    End If
    
    If False Then
err:
        ReturnKeyWordValue = ""
    End If
End Function
'/*20100122*/



'/*將指定的記憶體位置內的值以字串形式抓出來*/
Public Function PointerValueToString(ByVal Pointer As Long) As String
    Dim BufferString As String * 256
   
    '/*如果傳進來的位置是記憶體0的位置的話，則需要把它去掉。其餘的照常處理*/
    If Pointer = 0 Then
        PointerValueToString = "error"
    Else
        Call CopyMemory(ByVal BufferString, ByVal Pointer, 256)
        
        PointerValueToString = StripNulls(BufferString)
    End If
End Function
'/*20100527*/

'/*把字串中第一個找到的null字串以後的字全部都去掉，只保留前面的*/
Public Function StripNulls(ByVal OriginalString As String) As String
   If (InStr(OriginalString, Chr(0)) > 0) Then
      OriginalString = Left(OriginalString, InStr(OriginalString, Chr(0)) - 1)
   End If

   StripNulls = OriginalString
End Function
'/*20100527*/

'/*將指定的記憶體位置本身的值以字串形式抓出來*/
Public Function PointerCodeToString(ByVal Pointer As Long) As String
    Dim temp As String * 512

    Call lstrcpy(temp, Pointer)
    
    '/*如果進來的不是一個正確的記憶體位置，則直接傳回error。其餘才繼續處理*/
    If (InStr(1, temp, Chr(0)) = 0) Then
         PointerCodeToString = ""
    Else
         PointerCodeToString = Left(temp, InStr(1, temp, Chr(0)) - 1)
    End If
End Function
'/*20100527*/

