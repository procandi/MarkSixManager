Attribute VB_Name = "basObject"
'/******************************************************************/
'/*�����G�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*��m�򪫥�B�z������ƪ��a��C                                  */
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
'/*                                      Last Edit Date 2009/04/02 */
'/******************************************************************/
Option Explicit



'/**************************�򪫥�B�z�������`��***********************************/
Public Const VARTYPE_STRING As Integer = 8 '��VarType()�禡�h�P�_����A�Y�o��o�ӱ`�ƪ��ƭȡA�Y�N���r�ꪫ��
Public Const VARTYPE_COMMONDIALOG As Integer = 9 '��VarType()�禡�h�P�_����A�Y�o��o�ӱ`�ƪ��ƭȡA�Y�N����ܲ�����
Public Const VARTYPE_PICTUREBOX As Integer = 3 '��VarType()�禡�h�P�_����A�Y�o��o�ӱ`�ƪ��ƭȡA�Y�N��PictureBox����
Public Const VARTYPE_COMMAND As Integer = 11 '��VarType()�禡�h�P�_����A�Y�o��o�ӱ`�ƪ��ƭȡA�Y�N�����s����
'/**************************�p�حק諸(2009/05/02)***********************************/



'/****************************�Ω�N��Ʈw����ƶ�JListBox_Sourece�Ӫ��󪺨禡**************************/
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
'/*******************************�p�حק諸(2009/04/02)**************************/


'/****************************�Ω�N��Ʈw����ƶ�JComboBox_Sourece�Ӫ��󪺨禡**************************/
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
'/*******************************�p�حק諸(2009/04/02)**************************/


