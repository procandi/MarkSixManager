Attribute VB_Name = "Module5"
'/******************************************************************/
'/*說明：　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*放置跟版本控制有關的地方。                                      */
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
'/*                                      Last Edit Date 2015/09/22 */
'/******************************************************************/
Option Explicit

'get current app building version
Public Function GetVersion() As String
    Static strVer As String
    If strVer = "" Then
        strVer = Trim$(Str$(App.Major)) & "." & Format$(App.Minor, "##00") & "." & Format$(App.Revision, "0000")
    End If
    GetVersion = strVer
End Function
