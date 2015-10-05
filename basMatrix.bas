Attribute VB_Name = "basMatrix"
'/******************************************************************/
'/*說明：　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*放置跟所有矩陣運算有關的地方。                                  */
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
'/*                                      Last Edit Date 2009/10/29 */
'/******************************************************************/
Option Explicit



'/*矩陣加法*/
Public Sub MatrixAddition(ByRef M() As Double, ByRef n() As Double, ByRef ReturnValue() As Double)
    Dim i As Long, j As Long, row As Long, column As Long
     
     
    row = UBound(M, 1)
    column = UBound(M, 2)
    ReDim ReturnValue(1 To row, 1 To column)
     
     
    For i = 1 To row
        For j = 1 To column
            ReturnValue(i, j) = M(i, j) + n(i, j)
        Next
    Next
End Sub
'/*小華修改的(20091029)*/


'/*矩陣減法*/
Public Sub MatrixSubtraction(ByRef M() As Double, ByRef n() As Double, ByRef ReturnValue() As Double)
    Dim i As Long, j As Long, row As Long, column As Long
     
     
    row = UBound(M, 1)
    column = UBound(M, 2)
    ReDim ReturnValue(1 To row, 1 To column)
     
     
    For i = 1 To row
        For j = 1 To column
            ReturnValue(i, j) = M(i, j) - n(i, j)
        Next
    Next
End Sub
'/*小華修改的(20091029)*/


'/*矩陣乘法*/
Public Sub MatrixMultiply(ByRef M() As Double, ByRef n() As Double, ByRef ReturnValue() As Double)
    Dim i As Long, j As Long, k As Long, row As Long, column As Long, max As Long
     
     
    row = UBound(M, 1)
    column = UBound(n, 2)
    max = UBound(M, 2)
    ReDim ReturnValue(1 To row, 1 To column)
     
     
    For i = 1 To row
        For j = 1 To column
            For k = 1 To max
                ReturnValue(i, j) = ReturnValue(i, j) + M(i, k) * n(k, j)
            Next
        Next
    Next
End Sub
'/*小華修改的(20091029)*/



'/*矩陣轉置*/
Public Sub MatrixTranspose(ByRef M() As Double, ByRef ReturnValue() As Double)
    Dim i As Long, j As Long, row As Long, column As Long
    
    
    row = UBound(M, 1)
    column = UBound(M, 2)
    ReDim ReturnValue(1 To column, 1 To row)
     
     
    For i = 1 To row
        For j = 1 To column
            ReturnValue(j, i) = M(i, j)
        Next
    Next
End Sub
'/*小華修改的(20091029)*/


'/*比對是否為對稱矩陣*/
Public Function IsSymmetric(ByRef M() As Double) As Boolean
    Dim i As Long, j As Long
    
    
    For i = 1 To UBound(M, 1)
        For j = 1 To UBound(M, 1)
            If M(j, i) <> M(i, j) Then
                IsSymmetric = False
                Exit Function
            End If
        Next
    Next
    
    IsSymmetric = True
End Function
'/*小華修改的(20091029)*/
