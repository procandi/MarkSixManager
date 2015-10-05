Attribute VB_Name = "basMath"
'/******************************************************************/
'/*�����G�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*��m��ƾǬ������ܼơB�`�ơB�禡�����a��C                      */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�ҲաG�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@    �@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�ѦҡG    �@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*����G    �@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�`�N�ƶ��G�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*                                           Edit By Edward Hsieh */
'/*                                      Last Edit Date 2010/05/18 */
'/******************************************************************/
Option Explicit

'/**************************�ƾǬ������`�Ƹ��***********************************/
Public Const DOUBLEEX_MAX As Double = 1.79769313486231E+308
'/**************************�p�حק諸(20100222)***********************************/




'/*********************���o�̤p�ȡA���i�H�bCondition�������Φr��U�B�~������***********************************/
'/*Example                                  */
'/*Dim Data(2) As Double,Condition As String*/
'/*Data(0) = 0                              */
'/*Data(1) = 10                             */
'/*Data(2) = 2                              */
'/*Condition = "Data[i]>0 && Data[i]<5"     */
'/*Debug.Print MinEx(Data,Condition)        */
'/*                                         */
Public Function MinEX(ByRef Data() As Double, ByVal Condition As String) As Double
    Dim i As Integer
    Dim flag As Boolean
    Dim min As Double
    
    flag = False
    min = DOUBLEEX_MAX
    
    Dim MSSC As New MSScriptControl.ScriptControl
    MSSC.Language = "JavaScript"
    Call MSSC.Eval("var i;")
    Call MSSC.Eval("var Data=new Array(" & UBound(Data) & ");")
   
    For i = 0 To UBound(Data) - 1
        Call MSSC.Eval("i=" & i & ";")
        Call MSSC.Eval("Data[i]=" & Data(i) & ";")
        
        If flag And min > Data(i) And MSSC.Eval(Condition & "?true:false;") Then
            min = Data(i)
        ElseIf Not flag And MSSC.Eval(Condition & "?true:false;") Then
            min = Data(i)
            flag = True
        End If
    Next

    MinEX = min
End Function
'/**************************�p�حק諸(2010/05/18)***********************************/
