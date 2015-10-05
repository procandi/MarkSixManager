Attribute VB_Name = "basDecryption"
'/******************************************************************/
'/*說明：　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*放置跟所有解密有關的地方。                                      */
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



'/*矩陣解密*/
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
'/*小華修改的(20091029)*/
