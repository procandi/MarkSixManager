Attribute VB_Name = "basDecoding"
'�[�ѱK���ؤl�s��0~15�A0�����[�K
Public Const Decoding_Seed_Number = 15

Public Const MAX_HOSTNAME_LEN = 132
Public Const MAX_DOMAIN_NAME_LEN = 132
Public Const MAX_SCOPE_ID_LEN = 260
Public Const MAX_ADAPTER_NAME_LENGTH = 260
Public Const MAX_ADAPTER_ADDRESS_LENGTH = 8
Public Const MAX_ADAPTER_DESCRIPTION_LENGTH = 132

Public Type IP_ADDR_STRING
            Next As Long
            IpAddress As String * 16
            IpMask As String * 16
            Context As Long
End Type

Public Type IP_ADAPTER_INFO
            Next As Long
            ComboIndex As Long
            AdapterName As String * MAX_ADAPTER_NAME_LENGTH
            Description As String * MAX_ADAPTER_DESCRIPTION_LENGTH
            AddressLength As Long
            Address(MAX_ADAPTER_ADDRESS_LENGTH - 1) As Byte
            Index As Long
            Type As Long
            DhcpEnabled As Long
            CurrentIpAddress As Long
            IpAddressList As IP_ADDR_STRING
            GatewayList As IP_ADDR_STRING
            DhcpServer As IP_ADDR_STRING
            HaveWins As Boolean
            PrimaryWinsServer As IP_ADDR_STRING
            SecondaryWinsServer As IP_ADDR_STRING
            LeaseObtained As Long
            LeaseExpires As Long
End Type

'Public Declare Function Netbios Lib "netapi32.dll" (pncb As NET_CONTROL_BLOCK) As Byte
'Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Public Declare Function GetProcessHeap Lib "kernel32" () As Long
Public Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Public Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal _
lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Integer, _
lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal _
lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public Declare Function GetAdaptersInfo Lib "IPHlpApi" (IpAdapterInfo As Any, pOutBufLen As Long) As Long

'���i�t�ε{���W�١A�Ω�g�Jcris_ini_log��ƪ�ɨϥΡA���P���{�����W�٧����P
Public Const ProgramName = "MINIPACS"

'INI�ɤ��ܼƪ��ƶq�W��
Public Const MAX_INI_VARIABLES = 200
'INI�ɤ����U�س]�w�W�ٻP��(0:Tag name, 1:Var name, 2:Value)
Global Ini_Variables_Name(MAX_INI_VARIABLES, 2) As String
Global Ini_Variables_Count As Integer
Global flg_INI_Trans As Boolean     '�O�_�Ĥ@��Ū��
'Global flg_IS_INI_Trans As Boolean  '�O�_�ҥ�ini��Ʈw
  
Public INIConnection_String As String           'INI�M�Ϊ���Ʈw�s�u�r��
Public INIConnection As New adoDB.Connection    'INI�M�Ϊ���Ʈw�s�u����
Public INIRecordset As New adoDB.Recordset      'INI�M�Ϊ���ƪ���
Public INISetupName As String                   '��Ʈw��INI�]�w�ɦW��
Public INIIP As String

' ���o �w�� �Ǹ�GetDiskSerialNumber("C:\")
Function GetDiskSerialNumber(strDrive As String) As String
    Dim SerialNum As Long
    GetVolumeInformation strDrive, vbNullString, _
    0, SerialNum, 0, 0, vbNullString, 0
    GetDiskSerialNumber = Hex(SerialNum)
End Function

' ���o �D���O �Ǹ�
Function Get_MB_SNo() As String
    Dim strCls As String, strKey As String
    Dim WMI As Object
    Set WMI = GetObject("winmgmts:")
    strCls = "Win32_BaseBoard" ' WMI ���O
    strKey = strCls & ".Tag=""Base Board"""
    Get_MB_SNo = Trim(WMI.InstancesOf(strCls)(strKey).SerialNumber)
End Function

'���o���d�Ǹ�
Public Function GetPhysicalAddress() As String
    Dim AdapterInfoSize As Long
    Dim i As Integer
    Dim PhysicalAddress  As String
    Dim AdapterInfo As IP_ADAPTER_INFO
    Dim Buffer2 As IP_ADAPTER_INFO

    On Error GoTo ErrMsg

    GetAdaptersInfo ByVal 0&, AdapterInfoSize

    ReDim AdapterInfoBuffer(AdapterInfoSize - 1)

    GetAdaptersInfo AdapterInfoBuffer(0), AdapterInfoSize

    CopyMemory AdapterInfo, AdapterInfoBuffer(0), Len(AdapterInfo)

    CopyMemory Buffer2, AdapterInfo, Len(Buffer2)
    
    For i = 0 To Buffer2.AddressLength - 1
        PhysicalAddress = PhysicalAddress & Right("0" & Hex(Buffer2.Address(i)), 2)
        If i < Buffer2.AddressLength - 1 Then
            PhysicalAddress = PhysicalAddress & "-"
        End If
    Next
    GetPhysicalAddress = PhysicalAddress
    Exit Function
ErrMsg:
    GetPhysicalAddress = "Error"
End Function

'���oIP�Ǹ�
Public Function GetIPAddress() As String
    Dim AdapterInfoSize As Long
    Dim i As Integer
    Dim PhysicalAddress  As String, Y$
    Dim AdapterInfo As IP_ADAPTER_INFO
    Dim Buffer2 As IP_ADAPTER_INFO

    On Error GoTo ErrMsg

    GetAdaptersInfo ByVal 0&, AdapterInfoSize

    ReDim AdapterInfoBuffer(AdapterInfoSize - 1)

    GetAdaptersInfo AdapterInfoBuffer(0), AdapterInfoSize

    CopyMemory AdapterInfo, AdapterInfoBuffer(0), Len(AdapterInfo)

    CopyMemory Buffer2, AdapterInfo, Len(Buffer2)
    PhysicalAddress = ""
    For i = 1 To Len(Buffer2.IpAddressList.IpAddress)
        Y$ = Mid(Buffer2.IpAddressList.IpAddress, i, 1)
        If Y$ >= "0" And Y$ <= "9" Then
            PhysicalAddress = PhysicalAddress & Y$
        End If
    Next
    
    GetIPAddress = PhysicalAddress
'    For i = 0 To Buffer2.AddressLength - 1
'        PhysicalAddress = PhysicalAddress & Right("0" & Hex(Buffer2.Address(i)), 2)
'        If i < Buffer2.AddressLength - 1 Then
'            PhysicalAddress = PhysicalAddress & "-"
'        End If
'    Next
'    GetPhysicalAddress = PhysicalAddress
    Exit Function
ErrMsg:
    GetIPAddress = "Error"
End Function

Public Function GetFullIPAddress() As String
    Dim AdapterInfoSize As Long
    Dim i As Integer
    Dim PhysicalAddress  As String, Y$
    Dim AdapterInfo As IP_ADAPTER_INFO
    Dim Buffer2 As IP_ADAPTER_INFO

    On Error GoTo ErrMsg

    GetAdaptersInfo ByVal 0&, AdapterInfoSize

    ReDim AdapterInfoBuffer(AdapterInfoSize - 1)

    GetAdaptersInfo AdapterInfoBuffer(0), AdapterInfoSize

    CopyMemory AdapterInfo, AdapterInfoBuffer(0), Len(AdapterInfo)

    CopyMemory Buffer2, AdapterInfo, Len(Buffer2)
    PhysicalAddress = ""
    For i = 1 To Len(Buffer2.IpAddressList.IpAddress)
        Y$ = Mid(Buffer2.IpAddressList.IpAddress, i, 1)
        If (Y$ >= "0" And Y$ <= "9") Or Y$ = "." Then
            PhysicalAddress = PhysicalAddress & Y$
        End If
    Next
'
'    GetIPAddress = PhysicalAddress
'    PhysicalAddress = ""
'    For i = 0 To Buffer2.AddressLength - 1
'        If PhysicalAddress <> "" Then
'            PhysicalAddress = PhysicalAddress & "."
'        End If
'        PhysicalAddress = PhysicalAddress & Right("000" & Buffer2.Address(i), 3)
'    Next
    GetFullIPAddress = PhysicalAddress
'    GetFullIPAddress = Trim(Buffer2.IpAddressList.IpAddress)
    Exit Function
ErrMsg:
    GetFullIPAddress = "Error"
End Function

Function Show_Seed(n As Byte) As String
    Dim strx As String
    Dim Str_Decoding(3) As String
    Dim i As Byte
    
    strx = ""
    Str_Decoding(0) = GetDiskSerialNumber("C:\")
    Str_Decoding(1) = Replace(GetPhysicalAddress, "-", "")
    Str_Decoding(2) = Replace(Get_MB_SNo, "/", "")
    Str_Decoding(3) = GetIPAddress
    For i = 0 To 3
        If (n And (2 ^ i)) <> 0 Then
            strx = strx & Trim(Str_Decoding(i))
        End If
    Next
    
    Show_Seed = strx
End Function

Public Function xDecoding(ByRef bIn() As Byte, ByRef bOut() As Byte) As Boolean
    Dim tg As Boolean
    Dim psw As String
    Dim pswB() As Byte
    Dim i As Integer, n As Integer, t As Integer, q As Integer
    
    tg = True
    On Error GoTo err_p
    psw = Show_Seed(Decoding_Seed_Number)
    t = Len(psw)
    If t > 0 Then
'        ReDim pswB(t - 1)
        pswB = StrConv(psw, vbUpperCase)
        q = UBound(pswB) + 1
    End If
    ReDim bOut(UBound(bIn))
    n = 0
'    q = UBound(pswB) + 1
    For i = 0 To UBound(bIn)
        If t > 0 Then
            bOut(i) = pswB(n) Xor bIn(i)
            n = (n + 1) Mod q
        Else
            bOut(i) = bIn(i)
        End If
    Next

    If False Then
err_p:
        tg = False
    End If
    xDecoding = tg
End Function

'��ĶINX���e
Function Trans_Ini_Array(tx As String) As Boolean
    Dim Result As Boolean
    Dim r() As String
    Dim i As Integer, x As Integer, n As Integer
    Dim Tag$, vName$, vValue$
    
    Result = True
    On Error GoTo Err_Proc
    r = Split(tx, vbCrLf)
    Tag$ = ""
    Ini_Variables_Count = 0
    For i = 0 To UBound(r)
        r(i) = Trim(r(i))
        If Left(r(i), 1) = "[" And InStr(r(i), "]") > 3 Then
            x = InStr(r(i), "]")
            Tag$ = Mid(r(i), 2, x - 2)
        ElseIf Left(r(i), 1) <> "\" And Left(r(i), 1) <> "/" And Left(r(i), 1) <> "'" And InStr(r(i), "=") > 3 Then
            x = InStr(r(i), "=")
            vName$ = Left(r(i), x - 1)
            vValue$ = Right(r(i), Len(r(i)) - x)
            Ini_Variables_Name(Ini_Variables_Count, 0) = UCase(Trim(Tag$))
            Ini_Variables_Name(Ini_Variables_Count, 1) = UCase(Trim(vName$))
            Ini_Variables_Name(Ini_Variables_Count, 2) = Trim(vValue$)
            Ini_Variables_Count = Ini_Variables_Count + 1
        End If
    Next
    If False Then
Err_Proc:
        Result = False
    End If
    Trans_Ini_Array = Result
End Function

'�d�ߨt�γ]�w�ȡA�Ĥ@������ɷ|Ū���ɮסA����N�u�|�j�M�}�C
Public Function xInputINI(ByVal ClassName As String, ByVal TitleName As String, ByVal FileName As String) As String
    Dim Inbyte() As Byte
    Dim Outbyte() As Byte
    Dim xStr As String
    Dim i As Integer
    
    If Not flg_INI_Trans Then
    '�Ĥ@������ɡA��Ū��INX��
        If FSO.FileExists(FileName) Then
            Open FileName For Binary Access Read As #1
                ReDim Inbyte(LOF(1) - 1)
                Get #1, , Inbyte
            Close #1
        Else
            MsgBox "�䤣��t�γ]�w�ɮסA���p���a�Q�u�{�v!!!"
            End
        End If
        If xDecoding(Inbyte(), Outbyte()) Then
            xStr = StrConv(Outbyte, vbUnicode)
            If Not Trans_Ini_Array(xStr) Then
                MsgBox "�t�γ]�w�ѪR���~�A���p���a�Q�u�{�v!!!"
                End
            End If
        Else
            MsgBox "�t�γ]�w�ѱK���~�A���p���a�Q�u�{�v!!!"
            End
        End If
        flg_INI_Trans = True
    End If
    
    'INI�����]�w��Ʈɤ~�d��
    xStr = ""
    If Ini_Variables_Count > 0 Then
        For i = 0 To Ini_Variables_Count
            If UCase(ClassName) = Ini_Variables_Name(i, 0) And UCase(TitleName) = Ini_Variables_Name(i, 1) Then
                xStr = Ini_Variables_Name(i, 2)
                Exit For
            End If
        Next
    End If
    xInputINI = xStr
End Function
