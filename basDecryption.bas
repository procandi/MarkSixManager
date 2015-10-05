Attribute VB_Name = "basDecryption"
'/******************************************************************/
'/*�����G�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*��m��Ҧ��ѱK�������a��C                                      */
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



'/*�x�}�ѱK*/
Public Sub MatrixDecode(ByRef strSource As String, ByRef iM() As Double, ByRef dblCoded() As Double)
    Dim i As Long, j As Long, n As Long, strM() As Double, strC() As Double
     
     
    n = UBound(iM, 2)
    ReDim strM(1 To n, 1 To 1) As Double, strC(1 To n, 1 To 1) As Double
     
     
    For i = 1 To UBound(dblCoded, 1)
        If i Mod n = 0 Then
            strM(n, 1) = dblCoded(i)
            Call MatrixMultiply(iM, strM, strC)
            For j = 1 To n
                strSource = strSource & ChrW(CLng(strC(j, 1)))
            Next
        Else
            strM(i Mod n, 1) = dblCoded(i)
        End If
    Next
End Sub
'/*�p�حק諸(20091029)*/
