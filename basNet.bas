Attribute VB_Name = "basNet"
'/******************************************************************/
'/*說明：　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*放置跟網路處理相關資料的地方。                                  */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*模組：　　　　　　　　　　　　　　　    　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*參考：    　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*元件：    　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*MSWINSCK.OCX。                                                  */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*注意事項：　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*                                           Edit By Edward Hsieh */
'/*                                      Last Edit Date 2009/04/02 */
'/******************************************************************/
Option Explicit



'/**************************跟網路處理有關的常數***********************************/
Public Const CONNECT_TIME_LIMIT = 5 '連線的延遲時間，超過時間的話，即算連不上，而不再重試
Public Const LISTEN_TIME_LIMIT = 5 '被連線的延遲時間，超過時間的話，即算連不上，而不再重試

Public Const DEFAULT_SOCKET_IP As String = "127.0.0.1" '預設的要連線的目標要用的IP位址
Public Const DEFAULT_SOCKET_PORT As String = "105" '預設的要連線的目標要走的Port
'/*******************************小華修改的(2009/04/02)**************************/



'/********************************用於連線別人電腦的函式**************************/
Public Function ConnectSocket(ByRef Socket As Winsock, ByVal IP As String, ByVal Port As String) As Boolean
    Dim i As Integer
    Dim temp() As String
    
    temp = Split(IP, ".")
    If UBound(temp) <> 3 Then
        ConnectSocket = False
        Exit Function
    End If
    For i = 0 To 3
        If IsNumeric(temp(i)) Then
            If Val(temp(i)) >= 256 Then
                ConnectSocket = False
                Exit Function
            End If
        Else
            ConnectSocket = False
            Exit Function
        End If
    Next
    
    Socket.RemoteHost = IP
    Socket.RemotePort = Port
    Socket.Connect
    
    ConnectSocket = True
End Function
'/*******************************小華修改的(2009/04/02)**************************/



'/*********************************用於等待連線的函式**************************/
Public Function ListenSocket(ByRef Socket As Winsock, ByVal Port As String) As Boolean
    Socket.LocalPort = Port
    Socket.Listen
    
    ListenSocket = True
End Function
'/*******************************小華修改的(2009/04/02)**************************/
