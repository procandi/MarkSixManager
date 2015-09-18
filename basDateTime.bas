Attribute VB_Name = "basDateTime"
'/******************************************************************/
'/*�����G�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*��m�����B�ɶ��������Ҧ��ܼơB�`�ơB�禡�����a��C            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�ҲաG�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@    �@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�ѦҡG�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@    �@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*����G    �@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�`�N�ƶ��G�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*                                           Edit By Edward Hsieh */
'/*                                      Last Edit Date 2009/03/02 */
'/******************************************************************/
Option Explicit



'/**************************�Ҽ{�Ҧ��]����p��X��Ӥ����������t�A�ä��H�~�B��B��^�ǰO���A�Ǧ^�����L�ȥN��p�⦨�\�P�_***********************************/
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
'/**************************�p�حק諸(2009/03/02)***********************************/



'/**************************�p��Ӥ�����~���A�O�_���|�~�H�O���ܶǦ^29�A���O���ܶǦ^28***********************************/
Public Function DateLeap(ByVal DATE_NOW As Date) As Integer
    If (Year(DATE_NOW) Mod 4 = 0 And Year(DATE_NOW) Mod 100 <> 0) Or Year(DATE_NOW) Mod 400 = 0 Then
        DateLeap = 29
    Else
        DateLeap = 28
    End If
End Function
'/**************************�p�حק諸(2009/03/02)***********************************/



'/**************************�p��Ӥ��������A���X�ѡH�öǦ^�Ӥ�����������***********************************/
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
'/**************************�p�حק諸(2009/03/02)***********************************/



'/**********************�Ω�Delay���禡�ADelaySecond���p����O��************/
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
'/**************************�p�حק諸(2009/03/27)***********************************/



'/**********************�Ω�Delay���禡�ADelaySecond���p����O��************/
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
'/**************************�p�حק諸(2009/05/14)***********************************/




'/***********************�Ω�p��Time1��Time2�������ɶ��t*******************/
Public Function TimeDiff(ByVal Time1 As Double, ByVal Time2 As Double) As Double
    Dim SecondDiff As Double
    
    SecondDiff = Abs(Time1 - Time2)
    TimeDiff = SecondDiff * 100000
End Function
'/**************************�p�حק諸(2009/04/02)***********************************/

