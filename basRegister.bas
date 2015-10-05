Attribute VB_Name = "basRegister"
'/******************************************************************/
'/*�����G�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*��m��Ҧ����U�X�������a��C                                    */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�ҲաG�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@    �@�@�@�@�@�@�@�@�@�@�@�@*/
'/*basWindowsAPI.bas�C                                             */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�ѦҡG�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@    �@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*����G    �@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�`�N�ƶ��G�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L                                                              */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*                                           Edit By Edward Hsieh */
'/*                                      Last Edit Date 2009/10/29 */
'/******************************************************************/
Option Explicit


'/*�B�z���U�ɥ�Win32API�`��*/
Public Const ERROR_SUCCESS As Long = 0&

Public Const HKEY_CLASSES_ROOT As Long = &H80000000
Public Const HKEY_CURRENT_USER As Long = &H80000001
Public Const HKEY_LOCAL_MACHINE As Long = &H80000002
Public Const HKEY_USERS As Long = &H80000003

Public Const KEY_QUERY_VALUE As Long = &H1
Public Const KEY_SET_VALUE As Long = &H2
Public Const KEY_ALL_ACCESS As Long = &H3F

Public Const REG_SZ As Long = 1
Public Const REG_BINARY As Long = 3                    ' Free form binary
Public Const REG_DWORD As Long = 4
'/**/


'/*�B�z���U�ɥΪ�Win32API�禡*/
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Public Declare Function RegQueryValueExBinary Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Public Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
 
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
'/**/


'/*�B�z���U�ɥΪ��`��*/
Public Const REGKEY_BUFFER As Long = 65535 'Ū���μg�J���U�X�ɡA�]�w���r�ꪺ�̰��w�Ī���
'/**/



'/*Ū�����U�X*/
Public Function QueryRegKey(ByVal Class As Long, ByVal Path As String, ByVal Name As String, ByVal Format As Long) As Variant
    Dim hKey As Long
    Dim sLength As Long
    Dim ReturnResult As Variant
    Dim ReturnCode As Long
    
    ReturnCode = RegOpenKeyEx(Class, Path, 0, KEY_QUERY_VALUE, hKey)
    If ReturnCode = 0 Then
        sLength = REGKEY_BUFFER
        
        Select Case Format
        Case REG_SZ
            Dim sValueString As String
            
            sValueString = String(sLength, Chr(0))
            ReturnCode = RegQueryValueExString(hKey, Name, 0&, REG_SZ, sValueString, sLength)
            
            ReturnResult = Left(sValueString, sLength - 1)
        Case REG_BINARY
            Dim i As Integer
            Dim sValueByte(REGKEY_BUFFER) As Byte
            
            ReturnCode = RegQueryValueExBinary(hKey, Name, 0&, REG_BINARY, sValueByte(0), sLength)
            
            For i = 0 To sLength - 1
                ReturnResult = ReturnResult & " " & sValueByte(i)
            Next
        Case REG_DWORD
            Dim sValueLong As Long
            
            ReturnCode = RegQueryValueExLong(hKey, Name, 0&, REG_DWORD, sValueLong, sLength)
            
            ReturnResult = sValueLong
        Case Else
            ReturnResult = -1
        End Select
        
        
        ReturnCode = RegCloseKey(hKey)
        
        
        If ReturnCode = ERROR_SUCCESS Then
            QueryRegKey = ReturnResult
        Else
            QueryRegKey = ReturnCode
        End If
    Else
        QueryRegKey = ReturnCode
    End If
End Function
'/**/



'/*�g�J���U�X*/
Public Function SetRegKey(ByVal sValue As Variant, ByVal sLength As Long, ByVal Class As Long, ByVal Path As String, ByVal Name As String, ByVal Format As Long) As Long
    Dim hKey As Long
    Dim ReturnCode As Long
    
    ReturnCode = RegOpenKeyEx(Class, Path, 0, KEY_SET_VALUE, hKey)
    If ReturnCode = 0 Then
        Select Case Format
        Case REG_SZ
            Dim sValueString As String
            
            sValueString = sValue & Chr(0)
            If sLength = 0 Then
                sLength = Len(sValueString)
            End If
            ReturnCode = RegSetValueExString(hKey, Name, 0&, REG_SZ, sValueString, sLength)
        Case REG_DWORD
            Dim sValueLong As Long
            
            sValueLong = sValue
            If sLength = 0 Then
                sLength = 4
            End If
            ReturnCode = RegSetValueExLong(hKey, Name, 0&, REG_DWORD, sValueLong, sLength)
        End Select
        
        
        ReturnCode = RegCloseKey(hKey)
        
        
        If ReturnCode = ERROR_SUCCESS Then
            SetRegKey = 0
        Else
            SetRegKey = ReturnCode
        End If
    Else
        SetRegKey = ReturnCode
    End If
End Function
'/**/

