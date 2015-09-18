Attribute VB_Name = "Module4"
Option Explicit

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const GWL_STYLE& = (-16)
Public Const BS_ICON& = &H40&
Public Const BS_BITMAP& = &H80&


Public Sub AddCommandButton(ByVal strIconPath As String)

Static intCount As Integer

Dim objCmdButton As VB.CommandButton
Dim lngStyle As Long
Dim lngRet As Long

'Increment Count to make unique button name
intCount = intCount + 1

'Create button
Set objCmdButton = frmButtonTest.Controls.Add("VB.CommandButton", "Test" + LTrim$(Str(intCount)))

'Set Style to Icon
lngStyle = GetWindowLong(objCmdButton.hwnd, GWL_STYLE)
lngRet = SetWindowLong(objCmdButton.hwnd, GWL_STYLE, lngStyle Or BS_ICON Or BS_BITMAP)

'Adjust buttons top and left positions to cascade multiple creates
objCmdButton.Top = objCmdButton.Top + ((intCount - 1) * 50)
objCmdButton.Left = objCmdButton.Left + ((intCount - 1) * 50)

'Set properties
objCmdButton.Enabled = True
objCmdButton.Visible = True
objCmdButton.ZOrder 0

'Load icon into picture property
objCmdButton.Picture = LoadPicture(strIconPath)

'refresh button
objCmdButton.Refresh
End Sub
