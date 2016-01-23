Attribute VB_Name = "basNew"
'/******************************************************************/
'/*說明：　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*放置新的函式。                                                  */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*模組：　　　　　　　　　　　　　　　    　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*參考：　　　　　　　　　　　　　　　    　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*元件：    　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*注意事項：　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*                                           Edit By Edward Hsieh */
'/*                                      Last Edit Date 2009/03/02 */
'/******************************************************************/
Option Explicit

Function SimpleRound(ByVal SourceNumber As Double, ByVal Length As Integer) As Double
    Dim number As String
    Dim i As Integer, addition As Integer, zero As String, temp() As String, result1 As Double, result2 As Double
    Dim last As Integer, begin As Integer, position As String
    
    
    If SourceNumber < 0 Then
        position = "-"
    Else
        position = ""
    End If
    number = Abs(SourceNumber)
    
    zero = ""
    For i = 1 To Length - 1
        zero = zero & "0"
    Next
    
    temp = Split(number, ".")
    If UBound(temp) > 0 Then
        result1 = Val("0." & temp(1))
    Else
        result1 = 0
    End If
    
    If result1 > 0 Then
        addition = 0
        last = Len(number) + 1
        begin = Length + Len(temp(0)) + 2
        For i = last To begin Step -1
            result2 = Val(Mid(number, i, 1)) + addition
            If result2 >= 5 Then
                addition = 1
            Else
                addition = 0
            End If
        Next
        
        If Length = 0 Then
            SimpleRound = Val(position & Int(number) + addition)
        Else
            SimpleRound = Val(position & Int(number) + Val("0." & Left(temp(1), Length)) + Val("0." + zero + Str(addition)))
        End If
    Else
        SimpleRound = SourceNumber
    End If
End Function
