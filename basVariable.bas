Attribute VB_Name = "basVariable"
'/******************************************************************/
'/*說明：　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*放置公用結構、常數及變數的地方。                                */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*模組：　　　　　　　　　　　　　　　    　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*參考：    　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*元件：    　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*注意事項：　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*若該公用結構、常數及變數為某功能才會用到的結構、常數及變數數，請*/
'/*歸類於該功能的模組下。                                          */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*                                           Edit By Edward Hsieh */
'/*                                      Last Edit Date 2009/02/26 */
'/******************************************************************/
Option Explicit


'/**************************公用常數的部份***********************************/
Public Const MIN_GRID_ROWS As Integer = 1 '建立Grid物件時，最少要有的Row數
Public Const MIN_GRID_COLS As Integer = 1 '建立Grid物件時，最少要有的Col數
Public Const PI As Double = 3.14159265358979
'/**************************小華修改的(2009/02/25)***********************************/


'/**************************公用變數的部份***********************************/
Public args() As String '用於放置切割讀入的Command的內容的字串變數
Public argc As Long '用於記錄切割讀入的Command的長度的長整數變數

Public FreeFilePort As Integer '用於記錄目前還空閒可以開的port的代號
'/**************************小華修改的(2009/04/02)***********************************/


'/*****************************系統關鍵變數***********************************************/
Public Action As String
Public Parameter As String
Public SelectCID As String
Public SelectCName As String
Public SelectPID As String
Public CurrentSwiftCode As Integer
Public SelectDate As String
'/**************************小華修改的(2015/10/05)***********************************/
