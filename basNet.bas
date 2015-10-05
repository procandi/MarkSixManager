Attribute VB_Name = "basNet"
'/******************************************************************/
'/*�����G�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*��m������B�z������ƪ��a��C                                  */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�ҲաG�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@    �@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�ѦҡG    �@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*����G    �@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*MSWINSCK.OCX�C                                                  */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�`�N�ƶ��G�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*                                           Edit By Edward Hsieh */
'/*                                      Last Edit Date 2009/04/02 */
'/******************************************************************/
Option Explicit



'/**************************������B�z�������`��***********************************/
Public Const CONNECT_TIME_LIMIT = 5 '�s�u������ɶ��A�W�L�ɶ����ܡA�Y��s���W�A�Ӥ��A����
Public Const LISTEN_TIME_LIMIT = 5 '�Q�s�u������ɶ��A�W�L�ɶ����ܡA�Y��s���W�A�Ӥ��A����

Public Const DEFAULT_SOCKET_IP As String = "127.0.0.1" '�w�]���n�s�u���ؼЭn�Ϊ�IP��}
Public Const DEFAULT_SOCKET_PORT As String = "105" '�w�]���n�s�u���ؼЭn����Port
'/*******************************�p�حק諸(2009/04/02)**************************/



'/********************************�Ω�s�u�O�H�q�����禡**************************/
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
'/*******************************�p�حק諸(2009/04/02)**************************/



'/*********************************�Ω󵥫ݳs�u���禡**************************/
Public Function ListenSocket(ByRef Socket As Winsock, ByVal Port As String) As Boolean
    Socket.LocalPort = Port
    Socket.Listen
    
    ListenSocket = True
End Function
'/*******************************�p�حק諸(2009/04/02)**************************/
