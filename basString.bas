Attribute VB_Name = "basString"
'/******************************************************************/
'/*�����G�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*��m��r��������ܼơB�`�ơB�禡�����a��C                      */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�ҲաG�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@    �@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�ѦҡG    �@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*����G    �@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�`�N�ƶ��G�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*                                           Edit By Edward Hsieh */
'/*                                      Last Edit Date 2009/03/03 */
'/******************************************************************/
Option Explicit


'/**************************�r��������`�Ƹ��***********************************/
Public Const ASCII_UPPER_A As Integer = 65
Public Const ASCII_LOWER_A As Integer = 97
Public Const ASCII_NULL As Integer = 0
Public Const ASCII_MAX As Integer = 128

Public Const CHAR_EN_MAX As Integer = 26
'/**************************�p�حק諸(20100222)***********************************/


'/**************************�r���������Ƹ��***********************************/
Public Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
'/**************************�p�حק諸(20100604)***********************************/




'/**************************��X�r��Source���A�r����Start��End�������r�������X�ӥN�JTarget�A��l���Else�A�̫�^��Target�@���X�Ӧr��(�ҦpStart��J"0"�BEnd��J"9"�A�h"0"~"9"�|�Q��JTarget�A��l�bElse)***********************************/
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
'/**************************�p�حק諸(2009/03/03)***********************************/


'/**************************�Ω�N�r��ɦ쪺�禡�AInsert_Count���n�ɦ��X��H�t�ƥN��n��Fill_Source_Char�ɦbFill_Target_String�e���A���ƫh���᭱�C(�ҡG�T�ӭȥѥ��ܥk���O�O-4,0�B99�A�h�Ǧ^�ȷ|��0099)***********************************/
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
'/**************************�p�حק諸(2009/03/16)***********************************/



'/**************************�Ω�N�r��ɦ쪺�禡�AInsert_Count���n�ɦ��X��H�t�ƥN��n��Fill_Source_String�ɦbFill_Target_String�e���A���ƫh���᭱�C(�ҡG�T�ӭȥѥ��ܥk���O�O-5,12�B99�A�h�Ǧ^�ȷ|��121299)***********************************/
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
'/**************************�p�حק諸(2009/03/16)***********************************/




'/*****************��X�bSource_String���A�@���X��Search_String�A�Ǧ^�ƶq**************/
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
'/**************************�p�حק諸(2009/03/24)***********************************/




'/*****************�p��t�����^��r�Φ�UniCode�BBIG5Code�V�X���r�ꪺ�u��r����*******************/
Public Function Str_TrueLen(ByVal Source_String As String) As Long
    Dim SingleChar As String
    Dim DoubleChar As String
    
    Call Str_Classify(Source_String, SingleChar, DoubleChar, 0, 128)
    Str_TrueLen = Len(SingleChar) + (LenB(DoubleChar) / 2)
End Function
'/**************************�p�حק諸(2009/05/15)***********************************/



'/*��^��r�̤G�Q���i���ഫ���Ʀr*/
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


'/*��Ʀr�̤G�Q���i���ഫ���^��r*/
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



'/*�Ǧ^�q�峹����쪺����r�᪺�Ҧ��r��*/
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



'/*�N���w���O�����m�����ȥH�r��Φ���X��*/
Public Function PointerValueToString(ByVal Pointer As Long) As String
    Dim BufferString As String * 256
   
    '/*�p�G�Ƕi�Ӫ���m�O�O����0����m���ܡA�h�ݭn�⥦�h���C��l���ӱ`�B�z*/
    If Pointer = 0 Then
        PointerValueToString = "error"
    Else
        Call CopyMemory(ByVal BufferString, ByVal Pointer, 256)
        
        PointerValueToString = StripNulls(BufferString)
    End If
End Function
'/*20100527*/

'/*��r�ꤤ�Ĥ@�ӧ�쪺null�r��H�᪺�r�������h���A�u�O�d�e����*/
Public Function StripNulls(ByVal OriginalString As String) As String
   If (InStr(OriginalString, Chr(0)) > 0) Then
      OriginalString = Left(OriginalString, InStr(OriginalString, Chr(0)) - 1)
   End If

   StripNulls = OriginalString
End Function
'/*20100527*/

'/*�N���w���O�����m�������ȥH�r��Φ���X��*/
Public Function PointerCodeToString(ByVal Pointer As Long) As String
    Dim temp As String * 512

    Call lstrcpy(temp, Pointer)
    
    '/*�p�G�i�Ӫ����O�@�ӥ��T���O�����m�A�h�����Ǧ^error�C��l�~�~��B�z*/
    If (InStr(1, temp, Chr(0)) = 0) Then
         PointerCodeToString = ""
    Else
         PointerCodeToString = Left(temp, InStr(1, temp, Chr(0)) - 1)
    End If
End Function
'/*20100527*/

