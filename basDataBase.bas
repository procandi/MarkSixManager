Attribute VB_Name = "basDataBase"
'/******************************************************************/
'/*說明：　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*放置跟資料庫有關的所有變數、常數、函式等的地方。                */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*模組：　　　　　　　　　　　　　　　    　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*參考：    　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*msado15.dll。                                                   */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*元件：    　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*注意事項：　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*無。　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*                                           Edit By Edward Hsieh */
'/*                                      Last Edit Date 2009/02/26 */
'/******************************************************************/
Option Explicit


'/**************************公用的資料庫、資料表連結變數及物件***********************************/
Public Connection_String As String '最基本讓整個系統使用的資料庫連線字串
Public Connection As New adoDB.Connection '最基本讓整個系統使用的資料庫連線物件
Public Recordset As New adoDB.Recordset '最基本讓整個系統使用的資料表物件

Public Source_Connection_String As String '萬一以上的不夠用，另一個讓整個系統使用的資料庫連線字串
Public Source_Connection As New adoDB.Connection '萬一以上的不夠用，另一個讓讓整個系統使用的資料庫連線物件
Public Source_Recordset As New adoDB.Recordset '萬一以上的不夠用，另一個讓讓整個系統使用的資料表物件

Public Target_Connection_String As String '萬一以上的不夠用，另一個讓整個系統使用的資料庫連線字串
Public Target_Connection As New adoDB.Connection '萬一以上的不夠用，另一個讓讓整個系統使用的資料庫連線物件
Public Target_Recordset As New adoDB.Recordset '萬一以上的不夠用，另一個讓讓整個系統使用的資料表物件
'/**************************小華修改的(2009/02/25)***********************************/



'/*                 用於連線到SQLSERVER的函式，傳回真代表成功，假代表失敗                             */
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
'/**************************小華修改的(2009/02/03)***********************************/



'/*                     用於連線到各種資料庫的函式，傳回真代表成功，假代表失敗                    */
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
'/**************************小華修改的(2009/02/03)***********************************/




'/*                      開啟資料表的函式，傳回真代表成功，假代表失敗                */
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
'/**************************小華修改的(2009/02/04)***********************************/



'/*    主要用途為用於資料庫的where判斷部份，能因傳入的欄位是不是null而修改傳回的sql指令           */
'/*               Example Input OpenRecordset(null)            */
Public Function isFieldsNull(ByRef Fields As String) As String
    If Fields = "Null" Then
        isFieldsNull = " is null"
    Else
        isFieldsNull = "='" & Fields & "'"
    End If
End Function
'/**************************小華修改的(2009/02/25)***********************************/

'2013/03/22
'增加更新資料庫內容時的LOG記錄功能
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

'若目前在資料庫內的狀態與現行記錄的狀態不符，表示可能在另一台電腦已先變更過/存檔/上傳報告，則需禁止報告被覆蓋
Public Function CheckStatus() As Boolean
    Dim SQL$, t As Boolean
        
    t = False
    SQL$ = "select * from cris_exam_online "
    SQL$ = SQL$ & " WHERE status<>'已刪除' and ChartNo='" & curr_Record.chartno & "' "
    SQL$ = SQL$ & " AND Type='" & curr_Record.Type & "' AND Uni_key='" & curr_Record.uni_key & "'"
    Call OpenRecordset(SQL$, Connection, Recordset)
    If Not Recordset.EOF Then
        If NoNull(Recordset("status")) = curr_Record.Status Then
            t = True
        End If
    End If
    CheckStatus = t
End Function

