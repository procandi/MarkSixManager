Attribute VB_Name = "Module2"
'Put in .bas module.  Or, change declarations to private for
'use in .cls or .frm modules

Public Declare Function MessageBeep Lib "user32" _
  (ByVal wType As Long) As Long
Public Const MB_ICONASTERISK = &H40&
Public Const MB_ICONEXCLAMATION = &H30&
Public Const MB_ICONHAND = &H10&
Public Const MB_ICONINFORMATION = MB_ICONASTERISK
Public Const MB_ICONMASK = &HF0&
Public Const MB_ICONQUESTION = &H20&
Public Const MB_ICONSTOP = MB_ICONHAND

Public Sub wavePlay()
    Dim sSoundName As String
    'sSoundName = "SystemStart" '�Ұ�Windows
    '�]�w�ȥi�H�O*****************************
    'sSoundName = ".DEFAULT" '�w�]
    'sSoundName = "SystemStart" '�Ұ�Windows
    'sSoundName = "SystemExit" '����Windows
    'sSoundName = "SystemHand" '��氱��
    'sSoundName = "SystemQuestion" '�ݸ�
    'sSoundName = "SystemExclamation" '����n
   ' sSoundName = "SystemAsterisk" '�P��
    'sSoundName = "Open" '�}�ҵ{��
    'sSoundName = "Close" '�����{��
    'sSoundName = "Maximize" '���̤j
    'sSoundName = "Minimize" '�Y��̤p
    'sSoundName = "RestoreDown" '�V�U�٭�
    'sSoundName = "RestoreUp" '�V�W�٭�
    'sSoundName = "AppGPFault" '�{�����~
    'sSoundName = "MenuCommand" '�\�����O
    'sSoundName = "MenuPopup" '�۲{�\���
    sSoundName = "MailBeep" '�s�l��q��
    '---------------------------------------------
    sndPlaySound sSoundName, &H1 '����

End Sub

Sub Log(X As String)
    
    Logger = Logger & vbCrLf & Now() & " " & X
    Open App.Path & "\wl_log" & Day(Date) & ".txt" For Append As #7
         Print #7, Now() & " " & X & vbCrLf
    Close #7
    If Len(Logger) > 64000 Then Logger = Right(Logger, 60000)

End Sub
