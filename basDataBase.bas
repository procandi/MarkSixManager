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
Public Connection As New adoDB.Connection '�̰�����Өt�ΨϥΪ���Ʈw�s�u����
Public Recordset As New adoDB.Recordset '�̰�����Өt�ΨϥΪ���ƪ���

Public Source_Connection_String As String '�U�@�H�W�������ΡA�t�@������Өt�ΨϥΪ���Ʈw�s�u�r��
Public Source_Connection As New adoDB.Connection '�U�@�H�W�������ΡA�t�@��������Өt�ΨϥΪ���Ʈw�s�u����
Public Source_Recordset As New adoDB.Recordset '�U�@�H�W�������ΡA�t�@��������Өt�ΨϥΪ���ƪ���

Public Target_Connection_String As String '�U�@�H�W�������ΡA�t�@������Өt�ΨϥΪ���Ʈw�s�u�r��
Public Target_Connection As New adoDB.Connection '�U�@�H�W�������ΡA�t�@��������Өt�ΨϥΪ���Ʈw�s�u����
Public Target_Recordset As New adoDB.Recordset '�U�@�H�W�������ΡA�t�@��������Өt�ΨϥΪ���ƪ���
'/**************************�p�حק諸(2009/02/25)***********************************/



'/*                 �Ω�s�u��SQLSERVER���禡�A�Ǧ^�u�N���\�A���N����                             */
'/*            Example Input Connect2SQLServer("127.0.0.1","DB1","User","Pass",Connection)            */
Public Function Connect2SQLServer(ByRef IP As String, ByRef DB As String, ByRef ID As String, ByRef PW As String, ByRef Connection As adoDB.Connection) As Boolean
    On Error GoTo err

    If Connection.State <> adoDB.adStateClosed Then
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
Public Function Connect2DataBase(ByRef Connection_String As String, ByRef Connection As adoDB.Connection) As Boolean
    On Error GoTo err

    If Connection.State <> adoDB.adStateClosed Then
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
Public Function OpenRecordset(ByRef SQL As String, ByRef Connection As adoDB.Connection, ByRef Recordset As adoDB.Recordset, Optional ByVal CursorType As Integer, Optional ByVal LockType As Integer) As Boolean
    On Error GoTo err
    
    If Recordset.State <> adoDB.adStateClosed Then
        Recordset.Close
    End If
     
    If CursorType = 0 Then
        CursorType = adoDB.adOpenStatic
    End If
    If LockType = 0 Then
        LockType = adoDB.adLockOptimistic
    End If
    
    Recordset.Open SQL, Connection, CursorType, LockType
    
    If Recordset.State Then
        OpenRecordset = True
    Else
        OpenRecordset = False
    End If

    If False Then
err:
        Call ErrorOut("Function OpenRecordset Error! :" & SQL)
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

'2013/03/22
'�W�[��s��Ʈw���e�ɪ�LOG�O���\��
Public Sub DBRecordLog(SQLType, SQLString, LogNote)
    Dim SQL$
    Dim adoDB As adoDB.Connection
    
    Set adoDB = New adoDB.Connection
    adoDB.Open dbConnection$
    
    On Error GoTo err
        SQL$ = "insert into cris_ReportLog ( "
        SQL$ = SQL$ & "uni_key, chartno, SqlString, SqlType, Logdate, Logtime, UserID, UserType, LogIP, LogNote "
        SQL$ = SQL$ & " ) values ( "
        SQL$ = SQL$ & " '" & curr_Record.uni_key & "', '" & curr_Record.chartno & "', "
        SQL$ = SQL$ & " '" & Replace(SQLString, "'", "''") & "', '" & SQLType & "', "
        SQL$ = SQL$ & " '" & Format(Now, "YYYY/MM/DD") & "', '" & Format(Now, "hh:mm:ss") & "', "
        SQL$ = SQL$ & " '" & UserID$ & "', '" & UserType$ & "', '" & GetFullIPAddress & "', "
        SQL$ = SQL$ & " '" & LogNote & "' )"
        Call PrintLog(SQL$)
'        Call Connection.Execute(sql$)
        
        adoDB.Execute SQL$
        adoDB.Close
        Set adoDB = Nothing
    
    If False Then
err:
        Call PrintLog("DBRecordLog error!")
    End If
    
End Sub

Public Sub DBRecordLogA(uni_key, chartno, SQLType, SQLString, LogNote)
    Dim SQL$
    Dim adoDB As adoDB.Connection
    
    Set adoDB = New adoDB.Connection
    adoDB.Open dbConnection$
    
    On Error GoTo err
        SQL$ = "insert into cris_ReportLog ( "
        SQL$ = SQL$ & "uni_key, chartno, SqlString, SqlType, Logdate, Logtime, UserID, UserType, LogIP, LogNote "
        SQL$ = SQL$ & " ) values ( "
        SQL$ = SQL$ & " '" & uni_key & "', '" & chartno & "', "
        SQL$ = SQL$ & " '" & Replace(SQLString, "'", "''") & "', '" & SQLType & "', "
        SQL$ = SQL$ & " '" & Format(Now, "YYYY/MM/DD") & "', '" & Format(Now, "hh:mm:ss") & "', "
        SQL$ = SQL$ & " '" & UserID$ & "', '" & UserType$ & "', '" & GetFullIPAddress & "', "
        SQL$ = SQL$ & " '" & LogNote & "' )"
        Call PrintLog(SQL$)
'        Call Connection.Execute(sql$)
        adoDB.Execute SQL$
        adoDB.Close
        Set adoDB = Nothing
    If False Then
err:
        Call PrintLog("DBRecordLog error!")
    End If
End Sub

'�Y�ثe�b��Ʈw�������A�P�{��O�������A���šA��ܥi��b�t�@�x�q���w���ܧ�L/�s��/�W�ǳ��i�A�h�ݸT����i�Q�л\
Public Function CheckStatus() As Boolean
    Dim SQL$, t As Boolean
        
    t = False
    SQL$ = "select * from cris_exam_online "
    SQL$ = SQL$ & " WHERE status<>'�w�R��' and ChartNo='" & curr_Record.chartno & "' "
    SQL$ = SQL$ & " AND Type='" & curr_Record.Type & "' AND Uni_key='" & curr_Record.uni_key & "'"
    Call OpenRecordset(SQL$, Connection, Recordset)
    If Not Recordset.EOF Then
        If NoNull(Recordset("status")) = curr_Record.Status Then
            t = True
        End If
    End If
    CheckStatus = t
End Function

