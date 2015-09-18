Attribute VB_Name = "customFunctions"
Public Function GetWeeks(ByVal imm As Integer, ary() As Double, ByVal iMULTI As Integer, _
                         ByVal dMIN As Double, ByVal dMAX As Double, _
                         ByVal sMIN As String, ByVal sMAX As String) As String
'傳入值:    imm - 輸入值
'           ary - 二維陣列,第一維為要傳回結果的數據, 第二維為根據輸入值判斷依據的數據
'        iMULTI - 二維數據與輸入值所差的倍數, 如:10
'          dMIN - 輸入值小於二維數據最小值時的數據, 如:-1
'          dMAX - 輸入值大於二維數據最大值時的數據, 如:-99
'          sMIN - 計算結果若為 dMIN 時, 所要傳回字串, 如:"20以下"
'          sMAX - 計算結果若為 dMAX 時, 所要傳回字串, 如:"40以上"
'傳回值: 計算結果或 sMIN, sMAX
'
    Dim i As Integer
    Dim a1S, a1E, a2S, a2E As Integer
    Dim dWeeks As Double
    Dim sweeks As String
    a1S = LBound(ary, 1)
    a1E = UBound(ary, 1)
    a2S = LBound(ary, 2)
    a2E = UBound(ary, 2)
    '
    If imm = 0 Then
       dWeeks = 0
    ElseIf imm < ary(a1E, a2S) * iMULTI Then
       dWeeks = dMIN
    ElseIf imm > ary(a1E, a2E) * iMULTI Then
       dWeeks = dMAX
'    ElseIf imm < ary(1, 0) * 10 Or imm > ary(1, 25) * 10 Then
'       dWeeks = 5.11726 + 1.01918 * (imm / 10)
    Else
        For i = a2S To a2E
            If imm = ary(a1E, i) * iMULTI Then
               dWeeks = ary(a1S, i)
               Exit For
            ElseIf imm < ary(a1E, i) * iMULTI Then
               dWeeks = ary(a1S, i - 1) + (ary(a1S, i) - ary(a1S, i - 1)) * ((imm - ary(a1E, i - 1) * iMULTI) / ((ary(a1E, i) - ary(a1E, i - 1)) * iMULTI))
               Exit For
            Else
            End If
        Next i
    End If
    
    Select Case dWeeks
           Case dMIN: sweeks = sMIN
           Case dMAX: sweeks = sMAX
           Case Else:
                sweeks = str(Round(dWeeks, 1))
    End Select
    '
    GetWeeks = sweeks
End Function

Public Function GetBPDWeeks(ByVal imm As Integer) As String
'輸入 BPD 的 mm 值, 求出預估的週數
'傳回 15-: 若小於最小值
'傳回 40+: 若大於最大值
'
    Dim dWeeks As Double
    Dim ary(0 To 1, 0 To 25) As Double
    Dim i As Integer
    'ary(0,--) 週數值
    For i = 0 To 25
        ary(0, i) = 15 + i
    Next i
    'ary(1,--) 輸入值
    ary(1, 0) = 3.3
    ary(1, 1) = 3.6
    ary(1, 2) = 3.9
    ary(1, 3) = 4.3
    ary(1, 4) = 4.6
    ary(1, 5) = 4.9
    ary(1, 6) = 5.2
    ary(1, 7) = 5.5
    ary(1, 8) = 5.8
    ary(1, 9) = 6.1
    ary(1, 10) = 6.4
    ary(1, 11) = 6.7
    ary(1, 12) = 7
    ary(1, 13) = 7.3
    ary(1, 14) = 7.5
    ary(1, 15) = 7.8
    ary(1, 16) = 8
    ary(1, 17) = 8.2
    ary(1, 18) = 8.4
    ary(1, 19) = 8.6
    ary(1, 20) = 8.8
    ary(1, 21) = 9
    ary(1, 22) = 9.1
    ary(1, 23) = 9.2
    ary(1, 24) = 9.3
    ary(1, 25) = 9.3
    
    GetBPDWeeks = GetWeeks(imm, ary(), 10, -1, 99, Format(ary(0, LBound(ary, 2)), "0") & "-", Format(ary(0, UBound(ary, 2)), "0") & "+")

End Function

Public Function GetFLWeeks(ByVal imm As Integer) As String
'輸入 FL 的 mm 值, 求出預估的週數
'傳回 15-: 若小於最小值
'傳回 40+: 若大於最大值
'
    Dim dWeeks As Double
    Dim ary(0 To 1, 0 To 25) As Double
    Dim i As Integer
    For i = 0 To 25
        ary(0, i) = 15 + i
    Next i
    ary(1, 0) = 2
    ary(1, 1) = 2.1
    ary(1, 2) = 2.6
    ary(1, 3) = 2.6
    ary(1, 4) = 3
    ary(1, 5) = 3.2
    ary(1, 6) = 3.5
    ary(1, 7) = 3.7
    ary(1, 8) = 4
    ary(1, 9) = 4.1
    ary(1, 10) = 4.5
    ary(1, 11) = 4.8
    ary(1, 12) = 5
    ary(1, 13) = 5.2
    ary(1, 14) = 5.4
    ary(1, 15) = 5.7
    ary(1, 16) = 5.8
    ary(1, 17) = 6.1
    ary(1, 18) = 6.2
    ary(1, 19) = 6.4
    ary(1, 20) = 6.5
    ary(1, 21) = 6.6
    ary(1, 22) = 6.7
    ary(1, 23) = 6.8
    ary(1, 24) = 6.8
    ary(1, 25) = 6.9
    
    GetFLWeeks = GetWeeks(imm, ary(), 10, -1, 99, Format(ary(0, LBound(ary, 2)), "0") & "-", Format(ary(0, UBound(ary, 2)), "0") & "+")
    
End Function

Public Function GetACWeeks(ByVal imm As Integer) As String
'輸入 AC 的 mm 值, 求出預估的週數
'傳回 15-: 若小於最小值
'傳回 40+: 若大於最大值
'
    Dim dWeeks As Double
    Dim ary(0 To 1, 0 To 25) As Double
    Dim i As Integer
    For i = 0 To 25
        ary(0, i) = 15 + i
    Next i
    ary(1, 0) = 9.6
    ary(1, 1) = 10.4
    ary(1, 2) = 11.6
    ary(1, 3) = 12.6
    ary(1, 4) = 13.6
    ary(1, 5) = 14.6
    ary(1, 6) = 15.6
    ary(1, 7) = 16.6
    ary(1, 8) = 17.6
    ary(1, 9) = 18.6
    ary(1, 10) = 19.6
    ary(1, 11) = 20.6
    ary(1, 12) = 21.6
    ary(1, 13) = 22.6
    ary(1, 14) = 23.5
    ary(1, 15) = 24.5
    ary(1, 16) = 25.5
    ary(1, 17) = 26.4
    ary(1, 18) = 27.4
    ary(1, 19) = 28.4
    ary(1, 20) = 29.3
    ary(1, 21) = 30.3
    ary(1, 22) = 31.2
    ary(1, 23) = 32.2
    ary(1, 24) = 33.1
    ary(1, 25) = 34
    
'    GetACWeeks = GetWeeks(imm, ary(), 10, -1, 99, "15-", "40+")
    GetACWeeks = GetWeeks(imm, ary(), 10, -1, 99, Format(ary(0, LBound(ary, 2)), "0") & "-", Format(ary(0, UBound(ary, 2)), "0") & "+")

End Function


Public Function GetGSWeeks(ByVal imm As Integer) As String
'輸入 GS 的 mm 值, 求出預估的週數
'傳回 6-: 若小於最小值
'傳回 8+: 若大於最大值
'
    Dim dWeeks As Double
    Dim ary(0 To 1, 0 To 2) As Double
    Dim i As Integer
    For i = 0 To 2
        ary(0, i) = 6 + i
    Next i
    ary(1, 0) = 1.5
    ary(1, 1) = 2.2
    ary(1, 2) = 2.5
    
    GetGSWeeks = GetWeeks(imm, ary(), 10, -1, 99, Format(ary(0, LBound(ary, 2)), "0") & "-", Format(ary(0, UBound(ary, 2)), "0") & "+")
    
End Function
Public Function GetCRLWeeks(ByVal imm As Integer) As String
'輸入 CRL 的 mm 值, 求出預估的週數
'傳回  5-: 若小於最小值
'傳回 14.3+: 若大於最大值
'
    Dim dWeeks As Double
    Dim ary(0 To 1, 0 To 73) As Double
    Dim i As Integer
    For i = 0 To 73
        ary(1, i) = 5 + i
    Next i
    ary(0, 0) = 6
    ary(0, 1) = 6.1
    ary(0, 2) = 6.3
    ary(0, 3) = 6.4
    ary(0, 4) = 6.6
    ary(0, 5) = 7
    ary(0, 6) = 7.1
    ary(0, 7) = 7.3
    ary(0, 8) = 7.4
    ary(0, 9) = 7.5
    ary(0, 10) = 7.6
    ary(0, 11) = 8
    ary(0, 12) = 8.1
    ary(0, 13) = 8.2
    ary(0, 14) = 8.3
    ary(0, 15) = 8.4
    ary(0, 16) = 8.5
    ary(0, 17) = 8.6
    ary(0, 18) = 8.6
    ary(0, 19) = 9
    ary(0, 20) = 9.1
    ary(0, 21) = 9.2
    ary(0, 22) = 9.3
    ary(0, 23) = 9.4
    ary(0, 24) = 9.4
    ary(0, 25) = 9.5
    ary(0, 26) = 9.6
    ary(0, 27) = 9.6
    ary(0, 28) = 10
    ary(0, 29) = 10.1
    ary(0, 30) = 10.1
    ary(0, 31) = 10.2
    ary(0, 32) = 10.3
    ary(0, 33) = 10.4
    ary(0, 34) = 10.4
    ary(0, 35) = 10.5
    ary(0, 36) = 10.6
    ary(0, 37) = 10.6
    ary(0, 38) = 10.6
    ary(0, 39) = 11
    ary(0, 40) = 11.1
    ary(0, 41) = 11.1
    ary(0, 42) = 11.2
    ary(0, 43) = 11.3
    ary(0, 44) = 11.3
    ary(0, 45) = 11.4
    ary(0, 46) = 11.4
    ary(0, 47) = 11.5
    ary(0, 48) = 11.6
    ary(0, 49) = 11.6
    For i = 50 To 73
        ary(0, i) = Format((5 + i + 65) / 10, "0.0")
'        Debug.Print i, 5 + i, ary(0, i)
    Next i
    
'    GetACWeeks = GetWeeks(imm, ary(), 10, -1, 99, "15-", "40+")
    GetCRLWeeks = GetWeeks(imm, ary(), 1, -1, 99, Format(ary(0, LBound(ary, 2)), "0") & "-", Format(ary(0, UBound(ary, 2)), "0.0") & "+")

End Function
Public Function GetHCWeeks(ByVal imm As Integer) As String
'輸入 HC 的 mm 值, 求出預估的週數
'傳回 16-: 若小於最小值
'傳回 41+: 若大於最大值
'
    Dim dWeeks As Double
    Dim ary(0 To 1, 0 To 25) As Double
    Dim i As Integer
    For i = 0 To 25
        ary(0, i) = 16 + i
    Next i
    ary(1, 0) = 11.4 '16.0
    ary(1, 1) = 12.7 '17.0
    ary(1, 2) = 14 '18.0
    ary(1, 3) = 15.2 '19.0
    ary(1, 4) = 16.4 '20.0
    ary(1, 5) = 17.6 '21.0
    ary(1, 6) = 18.7 '22.0
    ary(1, 7) = 19.8 '23.0
    ary(1, 8) = 20.9 '24.0
    ary(1, 9) = 21.9 '25.0
    ary(1, 10) = 22.8 '26.0
    ary(1, 11) = 23.8 '27.0
    ary(1, 12) = 24.7 '28.0
    ary(1, 13) = 25.5 '29.0
    ary(1, 14) = 26.3 '30.0
    ary(1, 15) = 27.1 '31.0
    ary(1, 16) = 27.9 '32.0
    ary(1, 17) = 28.5 '33.0
    ary(1, 18) = 29.2 '34.0
    ary(1, 19) = 29.8 '35.0
    ary(1, 20) = 30.4 '36.0
    ary(1, 21) = 31  '37.0
    ary(1, 22) = 31.5 '38.0
    ary(1, 23) = 31.9 '39.0
    ary(1, 24) = 32.3 '40.0
    ary(1, 25) = 32.7 '41.0
    
'    GetACWeeks = GetWeeks(imm, ary(), 10, -1, 99, "15-", "40+")
    GetHCWeeks = GetWeeks(imm, ary(), 10, -1, 99, Format(ary(0, LBound(ary, 2)), "0") & "-", Format(ary(0, UBound(ary, 2)), "0") & "+")

End Function

Public Function GetHumerusWeeks(ByVal imm As Integer) As String
'輸入 Humerus 的 mm 值, 求出預估的週數
'傳回 13-: 若小於最小值
'傳回 42+: 若大於最大值
'
    Dim dWeeks As Double
    Dim ary(0 To 1, 0 To 29) As Double
    Dim i As Integer
    For i = 0 To 29
        ary(0, i) = 13 + i
    Next i
    ary(1, 0) = 1 '13.0
    ary(1, 1) = 1.2 '14.0
    ary(1, 2) = 1.4 '15.0
    ary(1, 3) = 1.7 '16.0
    ary(1, 4) = 2  '17.0
    ary(1, 5) = 2.3 '18.0
    ary(1, 6) = 2.6 '19.0
    ary(1, 7) = 2.9 '20.0
    ary(1, 8) = 3.2 '21.0
    ary(1, 9) = 3.3 '22.0
    ary(1, 10) = 3.7 '23.0
    ary(1, 11) = 3.8 '24.0
    ary(1, 12) = 4.2 '25.0
    ary(1, 13) = 4.3 '26.0
    ary(1, 14) = 4.5 '27.0
    ary(1, 15) = 4.7 '28.0
    ary(1, 16) = 4.8 '29.0
    ary(1, 17) = 5 '30.0
    ary(1, 18) = 5.3 '31.0
    ary(1, 19) = 5.4 '32.0
    ary(1, 20) = 5.6 '33.0
    ary(1, 21) = 5.8 '34.0
    ary(1, 22) = 5.9 '35.0
    ary(1, 23) = 6 '36.0
    ary(1, 24) = 6.1 '37.0
    ary(1, 25) = 6.4 '38.0
    ary(1, 26) = 6.5 '39.0
    ary(1, 27) = 6.6 '40.0
    ary(1, 28) = 6.6 '41.0
    ary(1, 29) = 6.8 '42.0
    
'    GetACWeeks = GetWeeks(imm, ary(), 10, -1, 99, "15-", "40+")
    GetHumerusWeeks = GetWeeks(imm, ary(), 10, -1, 99, Format(ary(0, LBound(ary, 2)), "0") & "-", Format(ary(0, UBound(ary, 2)), "0") & "+")

End Function

Public Function GetUlnaWeeks(ByVal imm As Integer) As String
'輸入 Ulna 的 mm 值, 求出預估的週數
'傳回 13-: 若小於最小值
'傳回 42+: 若大於最大值
'
    Dim dWeeks As Double
    Dim ary(0 To 1, 0 To 29) As Double
    Dim i As Integer
    For i = 0 To 29
        ary(0, i) = 13 + i
    Next i
    ary(1, 0) = 0.8 '13.0
    ary(1, 1) = 1 '14.0
    ary(1, 2) = 1.2 '15.0
    ary(1, 3) = 1.6 '16.0
    ary(1, 4) = 1.7 '17.0
    ary(1, 5) = 2.2 '18.0
    ary(1, 6) = 2.4 '19.0
    ary(1, 7) = 2.7 '20.0
    ary(1, 8) = 3 '21.0
    ary(1, 9) = 3.1 '22.0
    ary(1, 10) = 3.5 '23.0
    ary(1, 11) = 3.6 '24.0
    ary(1, 12) = 3.9 '25.0
    ary(1, 13) = 4 '26.0
    ary(1, 14) = 4.1 '27.0
    ary(1, 15) = 4.4 '28.0
    ary(1, 16) = 4.5 '29.0
    ary(1, 17) = 4.7 '30.0
    ary(1, 18) = 4.9 '31.0
    ary(1, 19) = 5 '32.0
    ary(1, 20) = 5.2 '33.0
    ary(1, 21) = 5.4 '34.0
    ary(1, 22) = 5.4 '35.0
    ary(1, 23) = 5.5 '36.0
    ary(1, 24) = 5.6 '37.0
    ary(1, 25) = 5.8 '38.0
    ary(1, 26) = 6 '39.0
    ary(1, 27) = 6 '40.0
    ary(1, 28) = 6.3 '41.0
    ary(1, 29) = 6.5 '42.0
    
'    GetACWeeks = GetWeeks(imm, ary(), 10, -1, 99, "15-", "40+")
    GetUlnaWeeks = GetWeeks(imm, ary(), 10, -1, 99, Format(ary(0, LBound(ary, 2)), "0") & "-", Format(ary(0, UBound(ary, 2)), "0") & "+")

End Function

Public Function GetTibiaWeeks(ByVal imm As Integer) As String
'輸入 Tibia 的 mm 值, 求出預估的週數
'傳回 13-: 若小於最小值
'傳回 42+: 若大於最大值
'
    Dim dWeeks As Double
    Dim ary(0 To 1, 0 To 29) As Double
    Dim i As Integer
    For i = 0 To 29
        ary(0, i) = 13 + i
    Next i
    ary(1, 0) = 0.9 '13.0
    ary(1, 1) = 1 '14.0
    ary(1, 2) = 1.3 '15.0
    ary(1, 3) = 1.6 '16.0
    ary(1, 4) = 1.8 '17.0
    ary(1, 5) = 2.2 '18.0
    ary(1, 6) = 2.5 '19.0
    ary(1, 7) = 2.7 '20.0
    ary(1, 8) = 3 '21.0
    ary(1, 9) = 3.2 '22.0
    ary(1, 10) = 3.6 '23.0
    ary(1, 11) = 3.7 '24.0
    ary(1, 12) = 4 '25.0
    ary(1, 13) = 4.2 '26.0
    ary(1, 14) = 4.4 '27.0
    ary(1, 15) = 4.5 '28.0
    ary(1, 16) = 4.6 '29.0
    ary(1, 17) = 4.8 '30.0
    ary(1, 18) = 5.1 '31.0
    ary(1, 19) = 5.2 '32.0
    ary(1, 20) = 5.4 '33.0
    ary(1, 21) = 5.7 '34.0
    ary(1, 22) = 5.8 '35.0
    ary(1, 23) = 6 '36.0
    ary(1, 24) = 6.1 '37.0
    ary(1, 25) = 6.2 '38.0
    ary(1, 26) = 6.4 '39.0
    ary(1, 27) = 6.5 '40.0
    ary(1, 28) = 6.6 '41.0
    ary(1, 29) = 6.8 '42.0
    
'    GetACWeeks = GetWeeks(imm, ary(), 10, -1, 99, "15-", "40+")
    GetTibiaWeeks = GetWeeks(imm, ary(), 10, -1, 99, Format(ary(0, LBound(ary, 2)), "0") & "-", Format(ary(0, UBound(ary, 2)), "0") & "+")

End Function

Public Function GetFibulaWeeks(ByVal imm As Integer) As String
'輸入 Fibula 的 mm 值, 求出預估的週數
'傳回 13-: 若小於最小值
'傳回 42+: 若大於最大值
'
    Dim dWeeks As Double
    Dim ary(0 To 1, 0 To 29) As Double
    Dim i As Integer
    For i = 0 To 29
        ary(0, i) = 13 + i
    Next i
    ary(1, 0) = 0.8 '13.0
    ary(1, 1) = 0.9 '14.0
    ary(1, 2) = 1.2 '15.0
    ary(1, 3) = 1.5 '16.0
    ary(1, 4) = 1.7 '17.0
    ary(1, 5) = 2.1 '18.0
    ary(1, 6) = 2.3 '19.0
    ary(1, 7) = 2.6 '20.0
    ary(1, 8) = 2.9 '21.0
    ary(1, 9) = 3.1 '22.0
    ary(1, 10) = 3.4 '23.0
    ary(1, 11) = 3.6 '24.0
    ary(1, 12) = 3.9 '25.0
    ary(1, 13) = 4 '26.0
    ary(1, 14) = 4.2 '27.0
    ary(1, 15) = 4.4 '28.0
    ary(1, 16) = 4.5 '29.0
    ary(1, 17) = 4.7 '30.0
    ary(1, 18) = 4.9 '31.0
    ary(1, 19) = 5.1 '32.0
    ary(1, 20) = 5.3 '33.0
    ary(1, 21) = 5.5 '34.0
    ary(1, 22) = 5.6 '35.0
    ary(1, 23) = 5.6 '36.0
    ary(1, 24) = 6  '37.0
    ary(1, 25) = 6  '38.0
    ary(1, 26) = 6.1 '39.0
    ary(1, 27) = 6.2 '40.0
    ary(1, 28) = 6.3 '41.0
    ary(1, 29) = 6.7 '42.0
    
'    GetACWeeks = GetWeeks(imm, ary(), 10, -1, 99, "15-", "40+")
    GetFibulaWeeks = GetWeeks(imm, ary(), 10, -1, 99, Format(ary(0, LBound(ary, 2)), "0") & "-", Format(ary(0, UBound(ary, 2)), "0") & "+")

End Function

Public Function GetRadiusWeeks(ByVal imm As Integer) As String
'輸入 Radius 的 mm 值, 求出預估的週數
'傳回 13-: 若小於最小值
'傳回 42+: 若大於最大值
'
    Dim dWeeks As Double
    Dim ary(0 To 1, 0 To 29) As Double
    Dim i As Integer
    For i = 0 To 29
        ary(0, i) = 13 + i
    Next i
    ary(1, 0) = 0.6  '13.0
    ary(1, 1) = 0.8 '14.0
    ary(1, 2) = 1.1 '15.0
    ary(1, 3) = 1.4 '16.0
    ary(1, 4) = 1.5 '17.0
    ary(1, 5) = 1.9 '18.0
    ary(1, 6) = 2.1 '19.0
    ary(1, 7) = 2.4 '20.0
    ary(1, 8) = 2.7 '21.0
    ary(1, 9) = 2.8 '22.0
    ary(1, 10) = 3.1 '23.0
    ary(1, 11) = 3.3 '24.0
    ary(1, 12) = 3.5 '25.0
    ary(1, 13) = 3.6 '26.0
    ary(1, 14) = 3.7 '27.0
    ary(1, 15) = 3.9 '28.0
    ary(1, 16) = 4 '29.0
    ary(1, 17) = 4.1 '30.0
    ary(1, 18) = 4.2 '31.0
    ary(1, 19) = 4.4 '32.0
    ary(1, 20) = 4.5 '33.0
    ary(1, 21) = 4.7 '34.0
    ary(1, 22) = 4.8 '35.0
    ary(1, 23) = 4.9 '36.0
    ary(1, 24) = 5.1 '37.0
    ary(1, 25) = 5.1 '38.0
    ary(1, 26) = 5.3 '39.0
    ary(1, 27) = 5.3 '40.0
    ary(1, 28) = 5.6 '41.0
    ary(1, 29) = 5.7 '42.0
    
'    GetACWeeks = GetWeeks(imm, ary(), 10, -1, 99, "15-", "40+")
    GetRadiusWeeks = GetWeeks(imm, ary(), 10, -1, 99, Format(ary(0, LBound(ary, 2)), "0") & "-", Format(ary(0, UBound(ary, 2)), "0") & "+")

End Function

Public Function SetNumeric(ByVal arg As String) As Variant

       If IsNumeric(arg) Then
          SetNumeric = True
       Else
          SetNumeric = False
       End If

End Function
