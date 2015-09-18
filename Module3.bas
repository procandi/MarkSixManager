Attribute VB_Name = "Module3"
Global OSVersion$
Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
        (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias _
        "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32s = 0

Function getOSInfo() As String
        Dim len5 As Long, aa As Long
        Dim cmprName As String
        Dim osver As OSVERSIONINFO
        Dim tmp$
        
        '取得Computer Name
        'cmprName = String(255, 0)
        'len5 = 256
        'aa = GetComputerName(cmprName, len5)
        'cmprName = VBA.left(cmprName, InStr(1, cmprName, Chr(0)) - 1)
        'Debug.Print "Computer Name = "; cmprName
        
        '取得OS的版本
        osver.dwOSVersionInfoSize = Len(osver)
        aa = GetVersionEx(osver)
        'Debug.Print "MajorVersion "; osver.dwMajorVersion
        'Debug.Print "MinorVersion "; osver.dwMinorVersion
        
        tmp$ = ""
        Select Case osver.dwPlatformId
        Case ER_PLATFORM_WIN32s
             tmp$ = "Microsoft Win32s "
        
        Case VER_PLATFORM_WIN32_WINDOWS
            If (osver.dwMajorVersion = 4) And (osver.dwMinorVersion = 0) Then
                tmp$ = "Microsoft Windows 95 "
                If (Mid(osver.szCSDVersion, 2, 1) = "C") Then
                    tmp$ = tmp$ & "OSR2 "
                End If
            ElseIf (osver.dwMajorVersion = 4) And (osver.dwMinorVersion = 10) Then
                If Mid(osver.szCSDVersion, 2, 1) = "A" Then
                    tmp$ = "Microsoft Windows 98 SE"
                Else
                    tmp$ = "Microsoft Windows 98"
                End If
        
            ElseIf (osver.dwMajorVersion = 4) And (osver.dwMinorVersion = 90) Then
                tmp$ = "Microsoft Windows Me "
            End If
        Case VER_PLATFORM_WIN32_NT
            If osver.dwMajorVersion <= 4 Then
                    tmp$ = "Microsoft Windows NT "
        
            ElseIf (osver.dwMajorVersion = 5) And (osver.dwMinorVersion) = 0 Then
                    tmp$ = "Microsoft Windows 2000 "
        
            ElseIf (osver.dwMajorVersion = 5) And (osver.dwMinorVersion = 1) Then
                    tmp$ = "Windows XP"
            End If
        End Select
        
        getOSInfo = tmp$
        
End Function


