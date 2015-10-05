Attribute VB_Name = "basArtificailNeuralNetwork"
'/******************************************************************/
'/*說明：　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*放置類神經網路相關函式。                                        */
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
'/*                                      Last Edit Date 2010/02/04 */
'/******************************************************************/
Option Explicit


'/*用於類神經網路使用的結構*/
Public Type ANN_Module
    ANN_p() As Double '使用者給予的題目
    ANN_t As Integer '使用者給予的答案
    ANN_width() As Double '加權
    ANN_Bias As Double '偏壓
    ANN_a As Integer '經過學習後計算出的答案
    ANN_e As Integer '計算使用者給予的答案跟程式學習後的答案是否相同(如果使用hardlims，-2為比答案還小，2為比答案還大，0為代表正確答案。如果使用hardlim，-為比答案還小，1為比答案還大，0為代表正確答案。)
    ANN_width_edit() As Double 'width_edit=修正加權
    ANN_Bias_edit As Double '修正偏壓
    ANN_Lumbda As Double '設定給認知器使用的學習因子的值
    
    ANN_p_w_count As Long '儲存輸入值及加權值有多少
    ANN_init As Boolean '告訴使用者這群變數是否初始化過
End Type
'/**/

'/*用於類神經網路使用的變數*/
Public ANN_HL As ANN_Module
'/**/


'/*hardlims=雙邊硬極限(為一函式，會產生1及-1兩種分類，分類方式為輸入的值>=0即是1，<0即是-1。)*/
Public Function hardlims(ByRef num As Double) As Integer
    If num >= 0 Then
        hardlims = 1
    Else
        hardlims = -1
    End If
End Function
'/**/

'/*hardlim=硬極限(為一函式，會產生1及0兩種分類，分類方式為輸入的值>=0即是1，<0即是0。)*/
Public Function hardlim(ByRef num As Double) As Integer
    If num >= 0 Then
        hardlim = 1
    Else
        hardlim = 0
    End If
End Function
'/**/
 
 
 
'/*建立感知器模型*/
Public Function Perceptron(ByRef HL As ANN_Module, ByVal P_W_Count As Long, ByVal Bias As Double, ByVal Lumbda As Double) As Boolean
    '/*初始化一些相關的資料，分別為資料筆數、樣本數、加權值數、加權修正值數、偏壓值、學習因子*/
    If Not HL.ANN_init Then
        HL.ANN_p_w_count = P_W_Count
    
        ReDim HL.ANN_p(HL.ANN_p_w_count)
        ReDim HL.ANN_width(HL.ANN_p_w_count)
        ReDim HL.ANN_width_edit(HL.ANN_p_w_count)
        HL.ANN_Bias = Bias
        
        HL.ANN_init = True
    End If
    HL.ANN_Lumbda = Lumbda
    '/**/
    
    
    Perceptron = True
End Function
'/**/


'/*建立單層前饋網路模型*/
Public Function Single_Layer_Feedforward_Networks(ByRef HL As ANN_Module, ByVal P_W_Count As Long, ByVal Bias As Double, ByVal Lumbda As Double) As Boolean
    '/*初始化一些相關的資料，分別為資料筆數、樣本數、加權值數*/
    If Not HL.ANN_init Then
        HL.ANN_p_w_count = P_W_Count
    
        ReDim HL.ANN_p(HL.ANN_p_w_count)
        ReDim HL.ANN_width(HL.ANN_p_w_count, HL.ANN_p_w_count)
        
        HL.ANN_init = True
    End If
    '/**/
    
    
    Single_Layer_Feedforward_Networks = True
End Function
'/**/


'/*採用Hebbian法則做訓練(只適用感知器模型)*/
Public Function Hebbian_Learning(ByRef HL As ANN_Module, ByRef p() As Double, ByVal Answer As Boolean) As Boolean
    Dim i As Long
    Dim temp As Double
    
    
    '/*將樣本p放進來*/
    For i = 0 To HL.ANN_p_w_count - 1
        HL.ANN_p(i) = p(i)
    Next
    '/**/
    
    
    '/*用Hebbian法則，將神經元及加權值累加，由類神經網路計算其認為的答案*/
    temp = 0
    For i = 0 To HL.ANN_p_w_count - 1
        temp = temp + (HL.ANN_p(i) * HL.ANN_width(i))
    Next
    HL.ANN_a = hardlims(temp + HL.ANN_Bias)
    '/**/
    

    '/*取得使用者為這個樣本，所指定的答案*/
    If Answer Then
         HL.ANN_t = 1
    Else
         HL.ANN_t = -1
    End If
    '/**/
        
        
    '/*讓使用者跟類神經網路比對答案，比對結果e=-2或e=2都是代表答錯，e=0才代表答對*/
    HL.ANN_e = HL.ANN_t - HL.ANN_a
    If HL.ANN_e = -2 Or HL.ANN_e = 2 Then
        '/*類神經網路答錯了，讓其依學習因子修正加權值、偏壓值等*/
        For i = 0 To HL.ANN_p_w_count - 1
            HL.ANN_width_edit(i) = HL.ANN_Lumbda * HL.ANN_e * HL.ANN_p(i)
            HL.ANN_width(i) = HL.ANN_width(i) + HL.ANN_width_edit(i)
        Next

        HL.ANN_Bias_edit = HL.ANN_Lumbda * HL.ANN_e
        HL.ANN_Bias = HL.ANN_Bias + HL.ANN_Bias_edit
        '/**/

        Hebbian_Learning = False
    Else
        Hebbian_Learning = True
    End If
    '/**/
End Function
'/**/


'/*採用Hebbian法則做回想(只適用感知器模型)*/
Public Function Hebbian_Recalling(ByRef HL As ANN_Module, ByRef p() As Double) As Boolean
    Dim i As Long
    Dim temp As Double
    
    
    '/*將樣本p放進來*/
    For i = 0 To HL.ANN_p_w_count - 1
        HL.ANN_p(i) = p(i)
    Next
    '/**/
    
    
    '/*用Hebbian法則，將神經元及加權值累加，由類神經網路計算結果*/
    temp = 0
    For i = 0 To HL.ANN_p_w_count - 1
        temp = temp + (HL.ANN_p(i) * HL.ANN_width(i))
    Next
    HL.ANN_a = hardlims(temp + HL.ANN_Bias)
    '/**/
    
    
    '/*傳回類神經用Hebbian法則算出來，其所認為答案是對或錯*/
    If HL.ANN_a = 1 Then
        Hebbian_Recalling = True
    Else
        Hebbian_Recalling = False
    End If
    '/**/
End Function
'/**/




'/*採用Hopfield法則做訓練(只適用前饋網路模型)*/
Public Function Hopfield_Learning(ByRef HL As ANN_Module, ByRef p() As Double) As Boolean
    Dim i As Long
    Dim j As Long
    
    
    '/*將樣本p放進來*/
    For i = 0 To HL.ANN_p_w_count - 1
        HL.ANN_p(i) = p(i)
    Next
    '/**/
    
    '/*建構Hopfield加權矩陣*/
    For i = 0 To HL.ANN_p_w_count - 1
        For j = 0 To HL.ANN_p_w_count - 1
            HL.ANN_width(i, j) = HL.ANN_width(i, j) + (HL.ANN_p(i) * HL.ANN_p(j))
        Next
    Next
    '/**/
    
    '/*將矩陣中的對角線元素值改為0，用以在後來聯想時，處理雜訊用途*/
    For i = 0 To HL.ANN_p_w_count - 1
        HL.ANN_width(i, i) = 0
    Next
    '/**/
    
    Hopfield_Learning = True
End Function
'/**/


'/*採用Hopfield法則做回想(只適用前饋網路模型)*/
Public Function Hopfield_Recalling(ByRef HL As ANN_Module, ByRef p() As Double, ByVal NCycle As Long) As Boolean
    Dim i As Long
    Dim j As Long
    Dim ICycle As Long
    Dim temp As Double
    Dim flag As Boolean
    
    
    '/*將樣本p放進來*/
    For i = 0 To HL.ANN_p_w_count - 1
        HL.ANN_p(i) = p(i)
    Next
    '/**/
    
    
    '/*限定其一定要在幾次的修正中收斂，不然會沒完沒了*/
    For ICycle = 0 To NCycle - 1
        '/*讓加權值乘樣本後再累計，再將聯想結果加以分類出新的樣本*/
        For j = 0 To HL.ANN_p_w_count - 1
            temp = 0
            For i = 0 To HL.ANN_p_w_count - 1
                temp = temp + (HL.ANN_width(i, j) * HL.ANN_p(i))
            Next
            
            Select Case temp
            Case Is < 0
                p(j) = -1
            Case Is = 0
                p(j) = HL.ANN_p(j)
            Case Is > 0
                p(j) = 1
            End Select
        Next
        '/**/
        
        
        '/*檢查類神經網路的新舊樣本是否相同，以判定是否收斂*/
        flag = True
        For i = 0 To HL.ANN_p_w_count - 1
            If HL.ANN_p(i) <> p(i) Then
                flag = False
                Exit For
            End If
        Next
        If flag Then
            Exit For
        Else
            For i = 0 To HL.ANN_p_w_count - 1
                HL.ANN_p(i) = p(i)
            Next
        End If
        '/**/
    Next
    '/**/
    
    
    '/*若flag還是未得到收斂的結果，即會傳回假，否則傳回真*/
    Hopfield_Recalling = flag
    '/**/
End Function
'/**/

