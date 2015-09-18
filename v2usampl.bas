Attribute VB_Name = "VGA2USB"
'VGA2USB IOCTL codes
Private Const IOCTL_VGA2USB_VIDEOMODE As Long = &H224027
Private Const IOCTL_VGA2USB_GETPARAMS As Long = &H22401B
Private Const IOCTL_VGA2USB_SETPARAMS As Long = &H228023
Private Const IOCTL_VGA2USB_GRABFRAME As Long = &H22002B
Private Const IOCTL_VGA2USB_GETSN As Long = &H22401F

'Parameter for IOCTL_VGA2USB_VIDEOMODE
Public Type V2U_VideoMode
  width As Long
  height As Long
  vfreg As Long
End Type

'Parameter for IOCTL_VGA2USB_GRABFRAME
Public Type V2U_GrabFrame
    pixbuf As Long
    pixbuflen As Long
    width As Long
    height As Long
    flags As Long
End Type

'V2U_GrabFrame flags
Public Const V2U_GRABFRAME_FORMAT_RGB16 As Long = &H10
Public Const V2U_GRABFRAME_FORMAT_RGB24 As Long = &H18
Public Const V2U_GRABFRAME_FORMAT_YUY2 As Long = &H100
Public Const V2U_GRABFRAME_FORMAT_YV12 As Long = &H200
Public Const V2U_GRABFRAME_FORMAT_2VUY As Long = &H300
Public Const V2U_GRABFRAME_BOTTOM_UP_FLAG As Long = &H80000000

'BMP file format constants
Private Const BMP_FILE_HEADER_SIZE As Long = 14
Public Const BMP_INFO_HEADER_SIZE As Long = 40
Public Const PELS_PER_METER As Long = 2835
Public Const BI_RGB As Long = 0
Public Const BI_BITFIELDS As Long = 3
Public Const RED_MASK As Integer = &H1F
Public Const GREEN_MASK As Integer = &H7E0
Public Const BLUE_MASK As Integer = &HF800

'Win32 constants
Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000
Private Const CREATE_ALWAYS As Long = 2
Private Const OPEN_EXISTING As Long = 3
Public Const INVALID_HANDLE_VALUE As Long = -1

'Kernel32 functions
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" ( _
    ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByVal lpSecurityAttributes As Any, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) _
    As Long

Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long) _
    As Long

Private Declare Function DeviceIoControl Lib "kernel32" ( _
    ByVal hDevice As Long, _
    ByVal dwIoControlCode As Long, _
    ByRef lpInBuffer As Any, _
    ByVal nInBufferSize As Long, _
    ByRef lpOutBuffer As Any, _
    ByVal nOutBufferSize As Long, _
    ByRef lpBytesReturned As Long, _
    ByVal lpOverlapped As Any) _
    As Long
   
Private Declare Function WriteFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByRef lpBuffer As Any, _
    ByVal nNumberOfBytesToWrite As Long, _
    ByRef lpNumberOfBytesWritten As Long, _
    ByVal lpOverlapped As Any) _
    As Long
    
Private Declare Function GetProcessHeap Lib "kernel32" () _
    As Long
    
Private Declare Function HeapAlloc Lib "kernel32" ( _
    ByVal hHeap As Long, _
    ByVal dwFlags As Long, _
    ByVal dwBytes As Long) _
    As Long
    
Private Declare Function HeapFree Lib "kernel32" ( _
    ByVal hHeap As Long, _
    ByVal dwFlags As Long, _
    lpMem As Any) _
    As Long
'Opens handle to VGA2USB driver. Returns INVALID_HANDLE_VALUE on failure.
'The caller must close the handle with VGA2USB_Close
Public Function VGA2USB_Open() As Long
    Dim i As Integer
    For i = 0 To 15
        Dim Name As String
        Name = "\\.\EpiphanVga2usb" & Trim$(str$(i))
        VGA2USB_Open = CreateFile(Name, _
            GENERIC_READ Or GENERIC_WRITE, _
            0&, _
            ByVal 0&, _
            OPEN_EXISTING, _
            0&, _
            0&)
        If VGA2USB_Open <> INVALID_HANDLE_VALUE Then
'            Debug.Print "VGA2USB: opened " & Name
            Exit Function
        End If
'        Debug.Print "VGA2USB: can't open " & Name & ", error" _
            & str$(err.LastDllError)
    Next i
End Function
'Closes the handle returns by VGA2USB_Open
Public Sub VGA2USB_Close(hDriver As Long)
    If hDriver <> INVALID_HANDLE_VALUE Then
        CloseHandle hDriver
    End If
End Sub
'Detects video mode
Public Function VGA2USB_GetVideoMode(hDevice As Long, _
                                     ByRef mode As V2U_VideoMode) As Boolean
    Dim bytesReturned As Long
    If DeviceIoControl(hDevice, IOCTL_VGA2USB_VIDEOMODE, _
        mode, Len(mode), mode, Len(mode), bytesReturned, ByVal 0&) <> 0 Then
        If mode.width <> 0 And mode.height <> 0 Then
            VGA2USB_GetVideoMode = True
'            Debug.Print "VGA2USB: detected " & VGA2USB_DescribeVideoMode(mode)
        Else
            VGA2USB_GetVideoMode = False
'            Debug.Print "VGA2USB: no signal detected"
        End If
    Else
        VGA2USB_GetVideoMode = False
'        Debug.Print "VGA2USB: IOCTL_VGA2USB_VIDEOMODE error " _
            & str$(err.LastDllError)
    End If
End Function
'Formats video mode string
Public Function VGA2USB_DescribeVideoMode(ByRef mode As V2U_VideoMode) As String
    If mode.width <> 0 And mode.height <> 0 Then
        VGA2USB_DescribeVideoMode = Trim$(str$(mode.width)) & " x" _
                & str$(mode.height) & " (" & str$(mode.vfreg / 1000) & " Hz )"
    Else
        VGA2USB_DescribeVideoMode = "No signal"
    End If
End Function
'Captures a single frame
Public Function VGA2USB_Capture(hDevice As Long, _
                                mode As V2U_VideoMode, _
                                ByRef frame As V2U_GrabFrame) As Boolean
                                      
    On Error GoTo capture_Error
    
    frame.pixbuflen = mode.width * mode.height * 2
    frame.flags = V2U_GRABFRAME_FORMAT_RGB16
'    Debug.Print "VGA2USB: allocating" & str$(frame.pixbuflen) _
        & " bytes for capture"
    frame.pixbuf = HeapAlloc(GetProcessHeap(), 0, frame.pixbuflen)
    If frame.pixbuf <> 0 Then
        Dim bytesReturned As Long
'        Debug.Print "VGA2USB: capturing the frame..."
        If DeviceIoControl(hDevice, IOCTL_VGA2USB_GRABFRAME, _
            frame, Len(frame), frame, Len(frame), bytesReturned, _
            ByVal 0&) <> 0 Then
            
            VGA2USB_Capture = True
'            Debug.Print "VGA2USB: captured" _
                & str$(frame.width) & " x" & str$(frame.height) & "," _
                & str$(frame.pixbuflen) & " bytes"
        Else
            VGA2USB_Capture = False
'            Debug.Print "VGA2USB: IOCTL_VGA2USB_GRABFRAME error " _
                & str$(err.LastDllError)
            HeapFree GetProcessHeap(), 0, grab.pixbuf
            grab.pixbuf = 0
        End If
    Else
        VGA2USB_Capture = False
'        Debug.Print "VGA2USB: failed to allocate" _
            & str$(grab.pixbuflen) & " bytes"
    End If
    On Error GoTo 0
    Exit Function
    
capture_Error:
    Resume Next
    
End Function
'Deallocates the buffer allocated by VGA2USB_GrabFrame
Public Sub VGA2USB_FreeBuffer(ByRef Buffer As V2U_GrabFrame)
    If Buffer.pixbuf <> 0 Then
        HeapFree GetProcessHeap(), 0, Buffer.pixbuf
        Buffer.pixbuf = 0
    End If
End Sub
'Writes 32-bit number in (hopefully) little endian byte order to a file
Private Sub WriteInt32(hFile As Long, data As Long)
    Dim n As Long
    WriteFile hFile, data, Len(data), n, ByVal 0&
End Sub
'Writes 16-bit number in (hopefully) little endian byte order to a file
Private Sub WriteInt16(hFile As Long, data As Integer)
    Dim n As Long
    WriteFile hFile, data, Len(data), n, ByVal 0&
End Sub
'Writes ASCII string to a file
Private Sub WriteStr(hFile As Long, data As String)
    Dim n As Long
    WriteFile hFile, ByVal data, Len(data), n, ByVal 0&
End Sub
'Writes V2U_GRABFRAME_FORMAT_RGB16 frame into a file
Private Function VGA2USB_SaveRGB16(hFile As Long, _
                                   ByRef frame As V2U_GrabFrame) As Boolean
    Dim y As Long
    Dim bmpsize As Long
    Dim rowsize As Long
    Dim datasize As Long
    Dim offbits As Long
    Dim filesize As Long
    
    bmpsize = BMP_INFO_HEADER_SIZE + 12
    rowsize = frame.width * 2
    datasize = frame.height * rowsize
    offbits = BMP_FILE_HEADER_SIZE + bmpsize
    filesize = offbits + datasize
    
    'Write file header
    WriteStr hFile, "BM"                    ' signature
    WriteInt32 hFile, filesize              ' bfSize
    WriteInt32 hFile, 0                     ' reserved
    WriteInt32 hFile, offbits               ' bfOffBits

    'Write BITMAP Header
    WriteInt32 hFile, BMP_INFO_HEADER_SIZE  ' biSize
    WriteInt32 hFile, frame.width           ' biWidth
    WriteInt32 hFile, frame.height          ' biHeight
    WriteInt16 hFile, 1                     ' biPlanes
    WriteInt16 hFile, 16                    ' biBitCount
    WriteInt32 hFile, BI_BITFIELDS          ' biCompression
    WriteInt32 hFile, datasize              ' biSizeImage
    WriteInt32 hFile, PELS_PER_METER        ' biXPelsPerMeter
    WriteInt32 hFile, PELS_PER_METER        ' biYPelsPerMeter
    WriteInt32 hFile, 0                     ' biClrUsed
    WriteInt32 hFile, 0                     ' biClrImportant

    'Write the RGB masks
    WriteInt16 hFile, RED_MASK
    WriteInt16 hFile, 0
    WriteInt16 hFile, GREEN_MASK
    WriteInt16 hFile, 0
    WriteInt16 hFile, BLUE_MASK
    WriteInt16 hFile, 0

    'Write the bitmap data
    For y = frame.height - 1 To 0 Step -1
        Dim row As Long
        row = frame.pixbuf + rowsize * y
        WriteFile hFile, ByVal row, rowsize, n, ByVal 0&
    Next y
    VGA2USB_SaveRGB16 = True
End Function
'Writes frame into a file
Public Function VGA2USB_SaveFrame(sFile As String, _
                                  ByRef frame As V2U_GrabFrame) As Boolean
    If frame.flags = V2U_GRABFRAME_FORMAT_RGB16 Then
        Dim hFile As Long
        hFile = CreateFile(sFile, GENERIC_WRITE, 0&, ByVal 0&, _
            CREATE_ALWAYS, 0&, 0&)
        If hFile <> INVALID_HANDLE_VALUE Then
'            Debug.Print "VGA2USB: writing " + sFile
            VGA2USB_SaveFrame = VGA2USB_SaveRGB16(hFile, frame)
            CloseHandle hFile
        Else
            VGA2USB_SaveFrame = False
'            Debug.Print "VGA2USB: can't create " + sFile + ", error " _
                & str$(err.LastDllError)
        End If
    Else
'        Debug.Print "VGA2USB: unsupported image format"
        VGA2USB_SaveFrame = False
    End If
End Function

