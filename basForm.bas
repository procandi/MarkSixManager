Attribute VB_Name = "basForm"
'/******************************************************************/
'/*說明：　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*放置跟表單處理有關的函式。                                      */
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
'/*                                      Last Edit Date 2009/07/21 */
'/******************************************************************/
Option Explicit


'/*處理表單的Win32API常數*/

'/*MessageBox用來決定要含哪些功能用的*/
Public Const MB_OK As Long = &H0&
Public Const MB_OKCANCEL As Long = &H1&
Public Const MB_RETRYCANCEL As Long = &H5&
Public Const MB_ABORTRETRYIGNORE As Long = &H2&
Public Const MB_YESNO As Long = &H4&
Public Const MB_YESNOCANCEL As Long = &H3&
Public Const MB_TOPMOST As Long = &H40000 '可以強制把MesageBox移動到作業系統的最上方
'/**/

'/*MessageBox用來判斷結果用的*/
Public Const IDYES As Long = 6
Public Const IDRETRY As Long = 4
Public Const IDOK As Long = 1
Public Const IDNO As Long = 7
Public Const IDIGNORE As Long = 5
Public Const IDCANCEL As Long = 2
Public Const IDABORT As Long = 3
'/**/

'/*20100528*/

'/*處理表單的Win32API函式*/
Public Declare Function MessageBox Lib "user32.dll" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Public Declare Function MessageBoxEx Lib "user32.dll" Alias "MessageBoxExA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal uType As Long, ByVal wLanguageId As Long) As Long
'/*20100528*/



'/*用來記錄表單要移動到哪邊的變數*/
Public goX As Long
Public goY As Long
'/**/

'/*用來在無框模式下移動表單用的函式*/
Public Sub FormMouseDown(ByVal frm As Form, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
     goX = x
     goY = y
End Sub
Public Sub FormMouseMove(ByVal frm As Form, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
     If Button = vbLeftButton Then
        Dim TargetX As Long, TargetY As Long, BorderX As Long, BorderY As Long
        
        TargetY = y - goY
        TargetX = x - goX
        BorderX = (Screen.Width - frm.Left - frm.Width)
        BorderY = (Screen.Height - frm.Top - frm.Height)
        
        If TargetY > 0 And TargetY > BorderY Then
            TargetY = BorderY
        End If
        If TargetY < 0 And Abs(TargetY) > frm.Top Then
            TargetY = -frm.Top
        End If
        If TargetX > 0 And TargetX > BorderX Then
            TargetX = BorderX
        End If
        If TargetX < 0 And Abs(TargetX) > frm.Left Then
            TargetX = -frm.Left
        End If
        
        Call frm.Move(frm.Left + TargetX, frm.Top + TargetY)
     End If
End Sub
'/**/
