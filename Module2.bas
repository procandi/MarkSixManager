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
    'sSoundName = "SystemStart" '啟動Windows
    '設定值可以是*****************************
    'sSoundName = ".DEFAULT" '預設
    'sSoundName = "SystemStart" '啟動Windows
    'sSoundName = "SystemExit" '結束Windows
    'sSoundName = "SystemHand" '緊急停止
    'sSoundName = "SystemQuestion" '問號
    'sSoundName = "SystemExclamation" '驚嘆聲
   ' sSoundName = "SystemAsterisk" '星號
    'sSoundName = "Open" '開啟程式
    'sSoundName = "Close" '關閉程式
    'sSoundName = "Maximize" '放到最大
    'sSoundName = "Minimize" '縮到最小
    'sSoundName = "RestoreDown" '向下還原
    'sSoundName = "RestoreUp" '向上還原
    'sSoundName = "AppGPFault" '程式錯誤
    'sSoundName = "MenuCommand" '功能表指令
    'sSoundName = "MenuPopup" '蹦現功能表
    sSoundName = "MailBeep" '新郵件通知
    '---------------------------------------------
    sndPlaySound sSoundName, &H1 '播放

End Sub

Sub Log(X As String)
    
    Logger = Logger & vbCrLf & Now() & " " & X
    Open App.Path & "\wl_log" & Day(Date) & ".txt" For Append As #7
         Print #7, Now() & " " & X & vbCrLf
    Close #7
    If Len(Logger) > 64000 Then Logger = Right(Logger, 60000)

End Sub
