Attribute VB_Name = "basDateTime"
'/******************************************************************/
'/*說明：　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*放置跟日期、時間有關的所有變數、常數、函式等的地方。            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*模組：　　　　　　　　　　　　　　　    　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*參考：　　　　　　　　　　　　　　　    　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*元件：    　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*注意事項：　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*                                           Edit By Edward Hsieh */
'/*                                      Last Edit Date 2009/03/02 */
'/******************************************************************/
Option Explicit



'/**************************考慮所有因素後計算出兩個日期間的日期差，並分以年、月、日回傳記錄，傳回的布林值代表計算成功與否***********************************/
Public Function Year_Month_Day_Diff(ByVal Date1 As Date, ByVal Date2 As Date, ByRef YearDiff As Integer, ByRef MonthDiff As Integer, ByRef DayDiff As Integer) As Boolean
    Dim Date_Year As Integer
    Dim Date_Month As Integer
    Dim Date_Day As Integer
    Dim DATE_NOW As Date
    
    
    DayDiff = DateDiff("d", Date1, Date2)
    If DayDiff < 0 Then
        If swap(Date1, Date2) Then
            DayDiff = Abs(DayDiff)
        Else
            Year_Month_Day_Diff = False
        End If
    End If
    
    DATE_NOW = Date2
    Date_Year = Year(DATE_NOW)
    Date_Month = Month(DATE_NOW)
    Date_Day = Day(DATE_NOW)
    
    YearDiff = 0
    MonthDiff = 0
    Do Until DayDiff < Date_Is_28_29_30_31(DATE_NOW)
        DayDiff = DayDiff - Date_Is_28_29_30_31(DATE_NOW)
        MonthDiff = MonthDiff + 1
        If MonthDiff > 12 Then
            MonthDiff = 1
            YearDiff = YearDiff + 1
        End If
        
        DATE_NOW = DateTime.DateAdd("m", -1, DATE_NOW)
    Loop
    
    Year_Month_Day_Diff = True
End Function
'/**************************小華修改的(2009/03/02)***********************************/



'/**************************計算該日期的年份，是否為閏年？是的話傳回29，不是的話傳回28***********************************/
Public Function DateLeap(ByVal DATE_NOW As Date) As Integer
    If (Year(DATE_NOW) Mod 4 = 0 And Year(DATE_NOW) Mod 100 <> 0) Or Year(DATE_NOW) Mod 400 = 0 Then
        DateLeap = 29
    Else
        DateLeap = 28
    End If
End Function
'/**************************小華修改的(2009/03/02)***********************************/



'/**************************計算該日期的月份，有幾天？並傳回該月份應有的日期***********************************/
Public Function Date_Is_28_29_30_31(ByVal DATE_NOW As Date) As Integer
    Select Case Month(DATE_NOW)
    Case 1, 3, 5, 7, 8, 10, 12
        Date_Is_28_29_30_31 = 31
    Case 2
        Date_Is_28_29_30_31 = DateLeap(DATE_NOW)
    Case 4, 6, 9, 11
        Date_Is_28_29_30_31 = 30
    End Select
End Function
'/**************************小華修改的(2009/03/02)***********************************/



'/**********************用於Delay的函式，DelaySecond的計算單位是秒************/
Public Sub Delay(ByVal DelaySecond As Double)
    If DelaySecond > 0 Then
        Dim NowTime As Double
        Dim LoadTime As Double
    
        DelaySecond = DelaySecond / 100000
        LoadTime = DateTime.Now
        NowTime = DateTime.Now
        Do Until NowTime - LoadTime > DelaySecond
            NowTime = DateTime.Now
            DoEvents
        Loop
    End If
End Sub
'/**************************小華修改的(2009/03/27)***********************************/



'/**********************用於Delay的函式，DelaySecond的計算單位是秒************/
Public Function DelayBreakWhenFileExist(ByVal DelaySecond As Double, ByVal FilePath As String) As Boolean
    Dim FSO As New FileSystemObject
    
    If DelaySecond > 0 Then
        Dim NowTime As Double
        Dim LoadTime As Double
    
        DelaySecond = DelaySecond / 100000
        LoadTime = DateTime.Now
        NowTime = DateTime.Now
        Do Until NowTime - LoadTime > DelaySecond Or FSO.FileExists(FilePath)
            NowTime = DateTime.Now
            DoEvents
        Loop
        
        If NowTime - LoadTime > DelaySecond And Not FSO.FileExists(FilePath) Then
            DelayBreakWhenFileExist = False
        Else
            DelayBreakWhenFileExist = True
        End If
    End If
End Function
'/**************************小華修改的(2009/05/14)***********************************/




'/***********************用於計算Time1跟Time2之間的時間差*******************/
Public Function TimeDiff(ByVal Time1 As Double, ByVal Time2 As Double) As Double
    Dim SecondDiff As Double
    
    SecondDiff = Abs(Time1 - Time2)
    TimeDiff = SecondDiff * 100000
End Function
'/**************************小華修改的(2009/04/02)***********************************/

