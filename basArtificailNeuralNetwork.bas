Attribute VB_Name = "basArtificailNeuralNetwork"
'/******************************************************************/
'/*�����G�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*��m�����g���������禡�C                                        */
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
'/*                                      Last Edit Date 2010/02/04 */
'/******************************************************************/
Option Explicit


'/*�Ω������g�����ϥΪ����c*/
Public Type ANN_Module
    ANN_p() As Double '�ϥΪ̵������D��
    ANN_t As Integer '�ϥΪ̵���������
    ANN_width() As Double '�[�v
    ANN_Bias As Double '����
    ANN_a As Integer '�g�L�ǲ߫�p��X������
    ANN_e As Integer '�p��ϥΪ̵��������׸�{���ǲ᪺߫���׬O�_�ۦP(�p�G�ϥ�hardlims�A-2���񵪮��٤p�A2���񵪮��٤j�A0���N���T���סC�p�G�ϥ�hardlim�A-���񵪮��٤p�A1���񵪮��٤j�A0���N���T���סC)
    ANN_width_edit() As Double 'width_edit=�ץ��[�v
    ANN_Bias_edit As Double '�ץ�����
    ANN_Lumbda As Double '�]�w���{�����ϥΪ��ǲߦ]�l����
    
    ANN_p_w_count As Long '�x�s��J�ȤΥ[�v�Ȧ��h��
    ANN_init As Boolean '�i�D�ϥΪ̳o�s�ܼƬO�_��l�ƹL
End Type
'/**/

'/*�Ω������g�����ϥΪ��ܼ�*/
Public ANN_HL As ANN_Module
'/**/


'/*hardlims=����w����(���@�禡�A�|����1��-1��ؤ����A�����覡����J����>=0�Y�O1�A<0�Y�O-1�C)*/
Public Function hardlims(ByRef num As Double) As Integer
    If num >= 0 Then
        hardlims = 1
    Else
        hardlims = -1
    End If
End Function
'/**/

'/*hardlim=�w����(���@�禡�A�|����1��0��ؤ����A�����覡����J����>=0�Y�O1�A<0�Y�O0�C)*/
Public Function hardlim(ByRef num As Double) As Integer
    If num >= 0 Then
        hardlim = 1
    Else
        hardlim = 0
    End If
End Function
'/**/
 
 
 
'/*�إ߷P�����ҫ�*/
Public Function Perceptron(ByRef HL As ANN_Module, ByVal P_W_Count As Long, ByVal Bias As Double, ByVal Lumbda As Double) As Boolean
    '/*��l�Ƥ@�Ǭ�������ơA���O����Ƶ��ơB�˥��ơB�[�v�ȼơB�[�v�ץ��ȼơB�����ȡB�ǲߦ]�l*/
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


'/*�إ߳�h�e�X�����ҫ�*/
Public Function Single_Layer_Feedforward_Networks(ByRef HL As ANN_Module, ByVal P_W_Count As Long, ByVal Bias As Double, ByVal Lumbda As Double) As Boolean
    '/*��l�Ƥ@�Ǭ�������ơA���O����Ƶ��ơB�˥��ơB�[�v�ȼ�*/
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


'/*�ĥ�Hebbian�k�h���V�m(�u�A�ηP�����ҫ�)*/
Public Function Hebbian_Learning(ByRef HL As ANN_Module, ByRef p() As Double, ByVal Answer As Boolean) As Boolean
    Dim i As Long
    Dim temp As Double
    
    
    '/*�N�˥�p��i��*/
    For i = 0 To HL.ANN_p_w_count - 1
        HL.ANN_p(i) = p(i)
    Next
    '/**/
    
    
    '/*��Hebbian�k�h�A�N���g���Υ[�v�Ȳ֥[�A�������g�����p���{��������*/
    temp = 0
    For i = 0 To HL.ANN_p_w_count - 1
        temp = temp + (HL.ANN_p(i) * HL.ANN_width(i))
    Next
    HL.ANN_a = hardlims(temp + HL.ANN_Bias)
    '/**/
    

    '/*���o�ϥΪ̬��o�Ӽ˥��A�ҫ��w������*/
    If Answer Then
         HL.ANN_t = 1
    Else
         HL.ANN_t = -1
    End If
    '/**/
        
        
    '/*���ϥΪ̸������g������ﵪ�סA��ﵲ�Ge=-2��e=2���O�N�����Ae=0�~�N����*/
    HL.ANN_e = HL.ANN_t - HL.ANN_a
    If HL.ANN_e = -2 Or HL.ANN_e = 2 Then
        '/*�����g���������F�A����̾ǲߦ]�l�ץ��[�v�ȡB�����ȵ�*/
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


'/*�ĥ�Hebbian�k�h���^�Q(�u�A�ηP�����ҫ�)*/
Public Function Hebbian_Recalling(ByRef HL As ANN_Module, ByRef p() As Double) As Boolean
    Dim i As Long
    Dim temp As Double
    
    
    '/*�N�˥�p��i��*/
    For i = 0 To HL.ANN_p_w_count - 1
        HL.ANN_p(i) = p(i)
    Next
    '/**/
    
    
    '/*��Hebbian�k�h�A�N���g���Υ[�v�Ȳ֥[�A�������g�����p�⵲�G*/
    temp = 0
    For i = 0 To HL.ANN_p_w_count - 1
        temp = temp + (HL.ANN_p(i) * HL.ANN_width(i))
    Next
    HL.ANN_a = hardlims(temp + HL.ANN_Bias)
    '/**/
    
    
    '/*�Ǧ^�����g��Hebbian�k�h��X�ӡA��һ{�����׬O��ο�*/
    If HL.ANN_a = 1 Then
        Hebbian_Recalling = True
    Else
        Hebbian_Recalling = False
    End If
    '/**/
End Function
'/**/




'/*�ĥ�Hopfield�k�h���V�m(�u�A�Ϋe�X�����ҫ�)*/
Public Function Hopfield_Learning(ByRef HL As ANN_Module, ByRef p() As Double) As Boolean
    Dim i As Long
    Dim j As Long
    
    
    '/*�N�˥�p��i��*/
    For i = 0 To HL.ANN_p_w_count - 1
        HL.ANN_p(i) = p(i)
    Next
    '/**/
    
    '/*�غcHopfield�[�v�x�}*/
    For i = 0 To HL.ANN_p_w_count - 1
        For j = 0 To HL.ANN_p_w_count - 1
            HL.ANN_width(i, j) = HL.ANN_width(i, j) + (HL.ANN_p(i) * HL.ANN_p(j))
        Next
    Next
    '/**/
    
    '/*�N�x�}�����﨤�u�����ȧאּ0�A�ΥH�b����p�Q�ɡA�B�z���T�γ~*/
    For i = 0 To HL.ANN_p_w_count - 1
        HL.ANN_width(i, i) = 0
    Next
    '/**/
    
    Hopfield_Learning = True
End Function
'/**/


'/*�ĥ�Hopfield�k�h���^�Q(�u�A�Ϋe�X�����ҫ�)*/
Public Function Hopfield_Recalling(ByRef HL As ANN_Module, ByRef p() As Double, ByVal NCycle As Long) As Boolean
    Dim i As Long
    Dim j As Long
    Dim ICycle As Long
    Dim temp As Double
    Dim flag As Boolean
    
    
    '/*�N�˥�p��i��*/
    For i = 0 To HL.ANN_p_w_count - 1
        HL.ANN_p(i) = p(i)
    Next
    '/**/
    
    
    '/*���w��@�w�n�b�X�����ץ������ġA���M�|�S���S�F*/
    For ICycle = 0 To NCycle - 1
        '/*���[�v�ȭ��˥���A�֭p�A�A�N�p�Q���G�[�H�����X�s���˥�*/
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
        
        
        '/*�ˬd�����g�������s�¼˥��O�_�ۦP�A�H�P�w�O�_����*/
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
    
    
    '/*�Yflag�٬O���o�즬�Ī����G�A�Y�|�Ǧ^���A�_�h�Ǧ^�u*/
    Hopfield_Recalling = flag
    '/**/
End Function
'/**/

