Attribute VB_Name = "basMatrix"
'/******************************************************************/
'/*�����G�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*��m��Ҧ��x�}�B�⦳�����a��C                                  */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�ҲաG�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@    �@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�ѦҡG�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@    �@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*����G    �@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�`�N�ƶ��G�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*                                           Edit By Edward Hsieh */
'/*                                      Last Edit Date 2009/10/29 */
'/******************************************************************/
Option Explicit



'/*�x�}�[�k*/
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
'/*�p�حק諸(20091029)*/


'/*�x�}��k*/
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
'/*�p�حק諸(20091029)*/


'/*�x�}���k*/
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
'/*�p�حק諸(20091029)*/



'/*�x�}��m*/
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
'/*�p�حק諸(20091029)*/


'/*���O�_����ٯx�}*/
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
'/*�p�حק諸(20091029)*/
