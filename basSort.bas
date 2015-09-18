Attribute VB_Name = "basSort"
'/******************************************************************/
'/*說明：　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*放置跟排序相關的變數、常數、函式等的地方。                      */
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
'/*                                      Last Edit Date 2009/03/13 */
'/******************************************************************/
Option Explicit



'/********************************泡沫排序法***************************/
Public Function Bubble_Sort(ByRef Source_String() As String)
    Dim i As Long, j As Long, k As Long
    
    For i = 1 To UBound(Source_String) - 1
        For j = 0 To i - 1
            For k = 1 To Len(Source_String(i))
                If k > Len(Source_String(j)) Then
                    Exit For
                End If
            
                If Mid(Source_String(i), k, 1) < Mid(Source_String(j), k, 1) Then
                    Call swap(Source_String(i), Source_String(j))
                    Exit For
                ElseIf Mid(Source_String(i), k, 1) > Mid(Source_String(j), k, 1) Then
                    Exit For
                End If
            Next
        Next
    Next
    
    Bubble_Sort = True
End Function
'/*************************小華修改的(2009/03/13)********************************/

