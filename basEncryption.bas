Attribute VB_Name = "basEncryption"
'/******************************************************************/
'/*�����G�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*��m��Ҧ��[�K�������a��C                                      */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�ҲաG�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@    �@�@�@�@�@�@�@�@�@�@�@�@*/
'/*basMatrix.bas�C                                                 */
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



'/*�x�}�[�K*/
Public Sub MatrixEncode(ByRef strSource As String, ByRef M() As Double, ByRef dblCoded() As Double)
     Dim i As Long, j As Long, n As Long, temp As Long, strM() As Double, strC() As Double
     
     
    n = UBound(M, 2)
    temp = Len(strSource) Mod n
    strSource = strSource & String(IIf(temp = 0, 0, n - temp), " ")
    ReDim strM(1 To n, 1 To 1), strC(1 To n, 1 To 1), dblCoded(1 To Len(strSource))
     
     
    For i = 1 To Len(strSource)
        If i Mod n = 0 Then
            strM(n, 1) = AscW(Mid(strSource, i, 1))
            Call MatrixMultiply(M, strM, strC)
            For j = 1 To n
                dblCoded(i + j - n) = strC(j, 1)
            Next
        Else
            strM(i Mod n, 1) = AscW(Mid(strSource, i, 1))
        End If
    Next
End Sub
'/*�p�حק諸(20091029)*/


'/*�x�}�[�K����l�]�w*/
Public Function MakeMatrix(ByRef M() As Double, ByRef iM() As Double) As Boolean
     Dim i, strWant As String
     ReDim M(1 To 3, 1 To 3)
     ReDim iM(1 To 3, 1 To 3)
     
     
     M(1, 1) = 1: M(1, 2) = 2: M(1, 3) = 3: M(2, 1) = 4: M(2, 2) = 5
     M(2, 3) = 6: M(3, 1) = 7: M(3, 2) = 8: M(3, 3) = 10
     iM(1, 1) = -0.666666667: iM(1, 2) = -1.333333333: iM(1, 3) = 1
     iM(2, 1) = -0.666666667: iM(2, 2) = 3.666666667: iM(2, 3) = -2
     iM(3, 1) = 1: iM(3, 2) = -2: iM(3, 3) = 1
     
     
     MakeMatrix = True
End Function
'/*�p�حק諸(20091029)*/
