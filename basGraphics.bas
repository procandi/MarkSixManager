Attribute VB_Name = "basGraphics"
'/******************************************************************/
'/*�����G�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*��m��v���B�z�B����ø�Ϥ��ܧε����a��C                        */
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
'/*                                      Last Edit Date 2009/07/07 */
'/******************************************************************/
Option Explicit


'/*�B�z�v���B�z��ø�s�Ϊ�Win32API�禡*/
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'/**/

'/**/
Public Image_Rotate As Integer '�v���n��m�����ת��ܼ�
'/**/


'/*     �]�w�����ܦ��Ϥ��H�~�A���Q�z�����A�H���X�S�O�Ϊ��������A�åH���h���覡�������      */
'/*        Example Input SetWindowLayoutByImageAndThread(form1,0,"c:/test.bmp")       */
Public Function SetWindowLayoutByImageAndThread(ByRef frm As Form, ByVal Clr As ColorConstants, ByVal pic As String) As Boolean
    On Error GoTo errout:
    
    Dim i As Double
    
    frm.BackColor = Clr
    frm.Picture = LoadPicture(pic)
    frm.Show
    For i = 0 To 255
        Call SetWindowTransparentAndOpacity(frm, Clr, i)
        DoEvents
    Next
    
    SetWindowLayoutByImageAndThread = True
    
    If False Then
errout:
        SetWindowLayoutByImageAndThread = False
    End If
End Function
'/***********************�p�حק諸(2009/08/26)***************************/

'/*           �]�w�����ܦ��Ϥ��H�~�A���Q�z�����A�H���X�S�O�Ϊ�������      */
'/*        Example Input SetWindowLayoutByImage(form1,0,"c:/test.bmp")       */
Public Function SetWindowLayoutByImage(ByRef frm As Form, ByVal Clr As ColorConstants, ByVal pic As String) As Boolean
    On Error GoTo errout:
    
    frm.BackColor = Clr
    frm.Picture = LoadPicture(pic)
    Call SetWindowTransparent(frm, Clr)
    
    SetWindowLayoutByImage = True
    
    If False Then
errout:
        SetWindowLayoutByImage = False
    End If
End Function
'/***********************�p�حק諸(2009/07/07)***************************/


'/*          �]�w�����S�w���C���ܳz���A�z���פ��i�աA���i���w�n�z�����C��      */
'/*             Example Input SetWindowTransparent(form1,0)       */
Public Function SetWindowTransparent(ByRef frm As Form, ByVal Clr As ColorConstants) As Boolean
    On Error GoTo errout:
    
    Call SetWindowLong(frm.hWnd, GWL_EXSTYLE, GetWindowLong(frm.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
    Call SetLayeredWindowAttributes(frm.hWnd, Clr, 0, LWA_COLORKEY)
    
    SetWindowTransparent = True
    
    If False Then
errout:
        SetWindowTransparent = False
    End If
End Function
'/***********************�p�حק諸(2009/07/07)***************************/


'/*          �]�w�����ܳz���A���i���w�C��A���z���ץi��Prec�h�վ�         */
'/*             Example Input SetWindowOpacity(form1,100)       */
Public Function SetWindowOpacity(ByRef frm As Form, ByVal Prec As Integer) As Boolean
    On Error GoTo errout:
    
    If (Prec >= 0 Or Prec <= 255) Then
        Call SetWindowLong(frm.hWnd, GWL_EXSTYLE, GetWindowLong(frm.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
        Call SetLayeredWindowAttributes(frm.hWnd, 0, Prec, LWA_ALPHA)
    End If
    
    SetWindowOpacity = True
    
    If False Then
errout:
        SetWindowOpacity = False
    End If
End Function
'/***********************�p�حק諸(2009/07/07)***************************/

'/*          �]�w�����ܳz���A��i���w�C��z��         */
'/*             Example Input SetWindowTransparentAndOpacity(form1,0,100)       */
Public Function SetWindowTransparentAndOpacity(ByRef frm As Form, ByVal Clr As ColorConstants, ByVal Prec As Integer) As Boolean
    On Error GoTo errout:
    
    If (Prec >= 0 Or Prec <= 255) Then
        Call SetWindowLong(frm.hWnd, GWL_EXSTYLE, GetWindowLong(frm.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
        Call SetLayeredWindowAttributes(frm.hWnd, Clr, Prec, LWA_COLORKEY Or LWA_ALPHA)
    End If
    
    SetWindowTransparentAndOpacity = True
    
    If False Then
errout:
        SetWindowTransparentAndOpacity = False
    End If
End Function
'/***********************�p�حק諸(2009/08/26)***************************/









'/*     �]�wHwnd�ܦ��Ϥ��H�~�A���Q�z�����A�H���X�S�O�Ϊ��������A�åH���h���覡�������      */
'/*        Example Input SetHwndLayoutByImageAndThread(Picture1,0,"c:/test.bmp")       */
Public Function SetHwndLayoutByImageAndThread(ByRef picbox As PictureBox, ByVal Clr As ColorConstants, ByVal pic As String) As Boolean
    On Error GoTo errout:
    
    Dim i As Double
    
    picbox.BackColor = Clr
    picbox.Picture = LoadPicture(pic)
    For i = 0 To 255
        Call SetHwndTransparentAndOpacity(picbox.hWnd, Clr, i)
        DoEvents
    Next
    
    SetHwndLayoutByImageAndThread = True
    
    If False Then
errout:
        SetHwndLayoutByImageAndThread = False
    End If
End Function
'/***********************�p�حק諸(2009/08/26)***************************/


'/*           �]�wHwnd�ܦ��Ϥ��H�~�A���Q�z�����A�H���X�S�O�Ϊ�������      */
'/*        Example Input SetHwndLayoutByImage(Picture1.hWnd,0,"c:/test.bmp")       */
Public Function SetHwndLayoutByImage(ByRef picbox As PictureBox, ByVal Clr As ColorConstants, ByVal pic As String) As Boolean
    On Error GoTo errout:
    
    picbox.BackColor = Clr
    picbox.Picture = LoadPicture(pic)
    Call SetHwndTransparent(picbox.hWnd, Clr)
    
    SetHwndLayoutByImage = True
    
    If False Then
errout:
        SetHwndLayoutByImage = False
    End If
End Function
'/***********************�p�حק諸(2009/07/07)***************************/



'/*          �]�w�S�w��Hwnd���C���ܳz���A�z���פ��i�աA���i���w�n�z�����C��      */
'/*             Example Input SetHwndTransparent(Picture1.Hwnd,0)       */
Public Function SetHwndTransparent(ByRef hWnd As Long, ByVal Clr As ColorConstants) As Boolean
    On Error GoTo errout:
    
    Call SetWindowLong(hWnd, GWL_EXSTYLE, GetWindowLong(hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
    Call SetLayeredWindowAttributes(hWnd, Clr, 0, LWA_COLORKEY)
    
    SetHwndTransparent = True
    
    If False Then
errout:
        SetHwndTransparent = False
    End If
End Function
'/***********************�p�حק諸(2009/07/07)***************************/



'/*          �]�wHwnd�ܳz���A���i���w�C��A���z���ץi��Prec�h�վ�         */
'/*             Example Input SetHwndOpacity(Picture1.hWnd,100)       */
Public Function SetHwndOpacity(ByRef hWnd As Long, ByVal Prec As Integer) As Boolean
    On Error GoTo errout:
    
    If (Prec >= 0 Or Prec <= 255) Then
        Call SetWindowLong(hWnd, GWL_EXSTYLE, GetWindowLong(hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
        Call SetLayeredWindowAttributes(hWnd, 0, Prec, LWA_ALPHA)
    End If
    
    SetHwndOpacity = True
    
    If False Then
errout:
        SetHwndOpacity = False
    End If
End Function
'/***********************�p�حק諸(2009/07/07)***************************/



'/*          �]�wHwnd�ܳz���A��i���w�C��z��         */
'/*             Example Input SetHwndTransparentAndOpacity(Picture1.hWnd,0,100)       */
Public Function SetHwndTransparentAndOpacity(ByRef hWnd As Long, ByVal Clr As ColorConstants, ByVal Prec As Integer) As Boolean
    On Error GoTo errout:
    
    If (Prec >= 0 Or Prec <= 255) Then
        Call SetWindowLong(hWnd, GWL_EXSTYLE, GetWindowLong(hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
        Call SetLayeredWindowAttributes(hWnd, Clr, Prec, LWA_COLORKEY Or LWA_ALPHA)
    End If
    
    SetHwndTransparentAndOpacity = True
    
    If False Then
errout:
        SetHwndTransparentAndOpacity = False
    End If
End Function
'/***********************�p�حק諸(2009/08/26)***************************/



'/*�]�w�Ϥ���½�ਤ��*/
Public Function SetPictureRotate(ByRef picboxSource As PictureBox, ByRef picboxTarget As PictureBox, ByVal RotateAngle As Integer) As Boolean
    On Error GoTo errout:
    
    Dim intX As Integer
    Dim intY As Integer
    Dim intX1 As Integer
    Dim intY1 As Integer
    Dim dblX2 As Double
    Dim dblY2 As Double
    Dim dblX3 As Double
    Dim dblY3 As Double
    Dim dblThetaDeg As Double
    Dim dblThetaRad As Double
    
    'Initialize rotation angle
    dblThetaDeg = RotateAngle
    
    'Compute angle in radians
    dblThetaRad = dblThetaDeg * PI / 180
    
    'Set scale modes to pixels
    picboxSource.ScaleMode = vbPixels
    picboxTarget.ScaleMode = vbPixels
    For intX = 0 To picboxTarget.ScaleWidth
        intX1 = intX - picboxTarget.ScaleWidth \ 2
        For intY = 0 To picboxTarget.ScaleHeight
            intY1 = intY - picboxTarget.ScaleHeight \ 2
            
            'Rotate picture by dblThetaRad
            dblX2 = intX1 * Cos(-dblThetaRad) + intY1 * Sin(-dblThetaRad)
            dblY2 = intY1 * Cos(-dblThetaRad) - intX1 * Sin(-dblThetaRad)
            
            'Translate to center of picture box
            dblX3 = dblX2 + picboxSource.ScaleWidth \ 2
            dblY3 = dblY2 + picboxSource.ScaleHeight \ 2
            
            'If data point is in picboxSource, set its color in picboxTarget
            If dblX3 > 0 And dblX3 < picboxSource.ScaleWidth - 1 And dblY3 > 0 And dblY3 < picboxSource.ScaleHeight - 1 Then
                picboxTarget.PSet (intX, intY), picboxSource.Point(dblX3, dblY3)
            End If
        Next
    Next
    
    SetPictureRotate = True
    
    If False Then
errout:
        SetPictureRotate = False
    End If
End Function
'/*20100224*/


'/*��pictuebox���ϵe��t�@�ipicturebox���ϤW�A���쥻��picturebox���Ϸ|�⥦�z��*/
Public Function DrawPictureOpacity(ByRef pic As PictureBox, ByRef picMask As PictureBox, ByRef picReplace As PictureBox, ByRef picBackground As PictureBox, ByVal lngX As Long, ByVal lngY As Long) As Boolean
    On Error GoTo errout:
    
    Dim lngW As Long
    Dim lngH As Long
    
    'Save sizes in local variables once for speed
    lngW = pic.ScaleWidth
    lngH = pic.ScaleHeight
        
    'Save background at new location
    Call BitBlt(picReplace.hDC, 0, 0, lngW, lngH, picBackground.hDC, lngX, lngY, vbSrcCopy)
        
    'Apply mask
    Call BitBlt(picBackground.hDC, lngX, lngY, lngW, lngH, picMask.hDC, 0, 0, vbSrcAnd)
    
    'Draw picture
    Call BitBlt(picBackground.hDC, lngX, lngY, lngW, lngH, pic.hDC, 0, 0, vbSrcPaint)
    picBackground.Refresh
    
    DrawPictureOpacity = True
    
    If False Then
errout:
        DrawPictureOpacity = False
    End If
End Function
'/**/


'/*�u�n��m��DrawPictureOpacity�@�˪�����A�N����٭�DrawPictureOpacity�ҳy�����ʧ@*/
Public Function RestorePictureOpacity(ByRef picReplace As PictureBox, ByRef picBackground As PictureBox, ByVal lngX As Long, ByVal lngY As Long) As Boolean
    On Error GoTo errout:
    
    Dim lngW As Long
    Dim lngH As Long
    
    'Save sizes in local variables once for speed
    lngW = picReplace.ScaleWidth
    lngH = picReplace.ScaleHeight
            
    'Restore picture
    Call BitBlt(picBackground.hDC, lngX, lngY, lngW, lngH, picReplace.hDC, 0, 0, vbSrcCopy)
    picBackground.Refresh
    
    RestorePictureOpacity = True
    
    If False Then
errout:
        RestorePictureOpacity = False
    End If
End Function
'/*20100225*/
