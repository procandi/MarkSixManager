Attribute VB_Name = "basEncryption"
'/******************************************************************/
'/*說明：　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*放置跟所有加密有關的地方。                                      */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*模組：　　　　　　　　　　　　　　　    　　　　　　　　　　　　*/
'/*basMatrix.bas。                                                 */
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



'/*矩陣加密*/
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
'/*小華修改的(20091029)*/

