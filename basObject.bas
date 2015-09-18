Attribute VB_Name = "basObject"
'/******************************************************************/
'/*說明：　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*放置跟物件處理相關資料的地方。                                  */
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
'/*                                      Last Edit Date 2009/04/02 */
'/******************************************************************/
Option Explicit



'/**************************跟物件處理有關的常數***********************************/
Public Const VARTYPE_STRING As Integer = 8 '用VarType()函式去判斷物件，若得到這個常數的數值，即代表為字串物件
Public Const VARTYPE_COMMONDIALOG As Integer = 9 '用VarType()函式去判斷物件，若得到這個常數的數值，即代表為對話盒物件
Public Const VARTYPE_PICTUREBOX As Integer = 3 '用VarType()函式去判斷物件，若得到這個常數的數值，即代表為PictureBox物件
Public Const VARTYPE_COMMAND As Integer = 11 '用VarType()函式去判斷物件，若得到這個常數的數值，即代表為按鈕物件
'/**************************小華修改的(2009/05/02)***********************************/



'/****************************用於將資料庫的資料填入ListBox_Sourece該物件的函式**************************/
Public Function ListBox_LoadFrom_DataBase(ByRef ListBox_Source As ListBox, ByVal FieldsValue As String, ByVal RecordsetValue As String, ByVal WhereValue As String, ByVal OrderByValue As String, ByVal FileName As String) As Boolean
    Dim Connection As New adoDB.Connection
    Dim Recordset As New adoDB.Recordset
    Dim SQL_String As String
    
    Connection_String = InputINI("Database", "Connection", FileName)
    
    If Connect2DataBase(Connection_String, Connection) Then
        SQL_String = "select " & FieldsValue & " "
        SQL_String = SQL_String & "from " & RecordsetValue & " "
        If WhereValue <> "" Then
            SQL_String = SQL_String & "where " & WhereValue & " "
        End If
        If OrderByValue <> "" Then
            SQL_String = SQL_String & "order by " & OrderByValue & " "
        End If

        If OpenRecordset(SQL_String, Connection, Recordset, adOpenStatic, adLockReadOnly) Then
            Dim i As Long
            Dim String_Merge As String
            Dim SelectCount As Long
            
            SelectCount = Str_SearchCount(FieldsValue, ",")
                      
            Do Until Recordset.EOF
                String_Merge = ""
                For i = 0 To SelectCount
                    String_Merge = String_Merge & Recordset(i) & " "
                Next
                ListBox_Source.AddItem String_Merge
                
                Recordset.MoveNext
            Loop
        
            ListBox_LoadFrom_DataBase = True
        Else
            ListBox_LoadFrom_DataBase = False
        End If
    Else
        ListBox_LoadFrom_DataBase = False
    End If
End Function
'/*******************************小華修改的(2009/04/02)**************************/


'/****************************用於將資料庫的資料填入ComboBox_Sourece該物件的函式**************************/
Public Function ComboBox_LoadFrom_DataBase(ByRef ComboBox_Source As ComboBox, ByVal FieldsValue As String, ByVal RecordsetValue As String, ByVal WhereValue As String, ByVal OrderByValue As String, ByVal FileName As String) As Boolean
    Dim Connection As New adoDB.Connection
    Dim Recordset As New adoDB.Recordset
    Dim SQL_String As String
    
    Connection_String = InputINI("Database", "Connection", FileName)
    
    If Connect2DataBase(Connection_String, Connection) Then
        SQL_String = "select " & FieldsValue & " "
        SQL_String = SQL_String & "from " & RecordsetValue & " "
        If WhereValue <> "" Then
            SQL_String = SQL_String & "where " & WhereValue & " "
        End If
        If OrderByValue <> "" Then
            SQL_String = SQL_String & "order by " & OrderByValue & " "
        End If
        
        If OpenRecordset(SQL_String, Connection, Recordset, adOpenStatic, adLockReadOnly) Then
            Dim i As Long
            Dim String_Merge As String
            Dim SelectCount As Long
            
            SelectCount = Str_SearchCount(FieldsValue, ",")
                      
            Do Until Recordset.EOF
                String_Merge = ""
                For i = 0 To SelectCount
                    String_Merge = String_Merge & Recordset(i) & " "
                Next
                ComboBox_Source.AddItem String_Merge
                
                Recordset.MoveNext
            Loop
        
            ComboBox_LoadFrom_DataBase = True
        Else
            ComboBox_LoadFrom_DataBase = False
        End If
    Else
        ComboBox_LoadFrom_DataBase = False
    End If
End Function
'/*******************************小華修改的(2009/04/02)**************************/


