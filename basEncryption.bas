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

