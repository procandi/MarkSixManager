Attribute VB_Name = "basDataBase"
'/******************************************************************/
'/*�����G�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*��m���Ʈw�������Ҧ��ܼơB�`�ơB�禡�����a��C                */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�ҲաG�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@    �@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�ѦҡG    �@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*msado15.dll�C                                                   */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*����G    �@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�`�N�ƶ��G�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*                                           Edit By Edward Hsieh */
'/*                                      Last Edit Date 2009/02/26 */
'/******************************************************************/
Option Explicit


'/**************************���Ϊ���Ʈw�B��ƪ�s���ܼƤΪ���***********************************/
Public Connection_String As String '�̰�����Өt�ΨϥΪ���Ʈw�s�u�r��
Public Connection As New ADODB.Connection '�̰�����Өt�ΨϥΪ���Ʈw�s�u����
Public Recordset As New ADODB.Recordset '�̰�����Өt�ΨϥΪ���ƪ���

Public Source_Connection_String As String '�U�@�H�W�������ΡA�t�@������Өt�ΨϥΪ���Ʈw�s�u�r��
Public Source_Connection As New ADODB.Connection '�U�@�H�W�������ΡA�t�@��������Өt�ΨϥΪ���Ʈw�s�u����
Public Source_Recordset As New ADODB.Recordset '�U�@�H�W�������ΡA�t�@��������Өt�ΨϥΪ���ƪ���

Public Target_Connection_String As String '�U�@�H�W�������ΡA�t�@������Өt�ΨϥΪ���Ʈw�s�u�r��
Public Target_Connection As New ADODB.Connection '�U�@�H�W�������ΡA�t�@��������Өt�ΨϥΪ���Ʈw�s�u����
Public Target_Recordset As New ADODB.Recordset '�U�@�H�W�������ΡA�t�@��������Өt�ΨϥΪ���ƪ���
'/**************************�p�حק諸(2009/02/25)***********************************/



'/*                 �Ω�s�u��SQLSERVER���禡�A�Ǧ^�u�N���\�A���N����                             */
'/*            Example Input Connect2SQLServer("127.0.0.1","DB1","User","Pass",Connection)            */
Public Function Connect2SQLServer(ByRef IP As String, ByRef DB As String, ByRef ID As String, ByRef PW As String, ByRef Connection As ADODB.Connection) As Boolean
    On Error GoTo err

    If Connection.State <> ADODB.adStateClosed Then
        Connection.Close
    End If
    
    Connection.Open "Provider=SQLOLEDB;Data Source=" & IP & ";Initial Catalog=" & DB & ";User Id=" & ID & ";Password=" & PW & ";"

    If Connection.State Then
        Connect2SQLServer = True
    Else
        Connect2SQLServer = False
    End If

    If False Then
err:
        Call ErrorOut("Function Connect2SQLServer Error!")
        Connect2SQLServer = False
    End If
End Function
'/**************************�p�حק諸(2009/02/03)***********************************/



'/*                     �Ω�s�u��U�ظ�Ʈw���禡�A�Ǧ^�u�N���\�A���N����                    */
'/*            Example Input Connect2DataBase("Server=127.0.0.1;DRIVER={SQL Server};UID=User;PWD=Pass;DATABASE=DB1;",,Connection)            */
Public Function Connect2DataBase(ByRef Connection_String As String, ByRef Connection As ADODB.Connection) As Boolean
    On Error GoTo err

    If Connection.State <> ADODB.adStateClosed Then
        Connection.Close
    End If
    
    Connection.Open Connection_String

    If Connection.State Then
        Connect2DataBase = True
    Else
        Connect2DataBase = False
    End If

    If False Then
err:
        Call ErrorOut("Function Connect2DataBase Error!")
        Connect2DataBase = False
    End If
End Function
'/**************************�p�حק諸(2009/02/03)***********************************/




'/*                      �}�Ҹ�ƪ��禡�A�Ǧ^�u�N���\�A���N����                */
'/*            Example Input OpenRecordset("select * from NorthWind",Connection,Recordset)            */
Public Function OpenRecordset(ByRef SQL As String, ByRef Connection As ADODB.Connection, ByRef Recordset As ADODB.Recordset, Optional ByVal CursorType As Integer, Optional ByVal LockType As Integer) As Boolean
    On Error GoTo err
    
    If Recordset.State <> ADODB.adStateClosed Then
        Recordset.Close
    End If
     
    If CursorType = 0 Then
        CursorType = ADODB.adOpenStatic
    End If
    If LockType = 0 Then
        LockType = ADODB.adLockOptimistic
    End If
    
    Recordset.Open SQL, Connection, CursorType, LockType
    
    If Recordset.State Then
        OpenRecordset = True
    Else
        OpenRecordset = False
    End If

    If False Then
err:
        Call ErrorOut("Function OpenRecordset Error!")
        OpenRecordset = False
    End If
End Function
'/**************************�p�حק諸(2009/02/04)***********************************/



'/*    �D�n�γ~���Ω��Ʈw��where�P�_�����A��]�ǤJ�����O���Onull�ӭק�Ǧ^��sql���O           */
'/*               Example Input OpenRecordset(null)            */
Public Function isFieldsNull(ByRef Fields As String) As String
    If Fields = "Null" Then
        isFieldsNull = " is null"
    Else
        isFieldsNull = "='" & Fields & "'"
    End If
End Function
'/**************************�p�حק諸(2009/02/25)***********************************/

