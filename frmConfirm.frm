VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmConfirm 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "Report Confirm"
   ClientHeight    =   3105
   ClientLeft      =   6315
   ClientTop       =   7785
   ClientWidth     =   4680
   Icon            =   "frmConfirm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4680
   Begin VB.TextBox txtCurrentDate 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   4
      Top             =   840
      Width           =   1650
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&N ���@��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   2
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton cmdConfirm 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&O �T�@�w"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   1
      Top             =   2520
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker dtpCurrentDate 
      Height          =   360
      Left            =   1800
      TabIndex        =   5
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy/MM/dd"
      Format          =   103612419
      CurrentDate     =   37058
   End
   Begin VB.Label Label1 
      Alignment       =   2  '�m�����
      BackColor       =   &H00808080&
      BackStyle       =   0  '�z��
      Caption         =   "�����(�`�b)�H��w�����ѳ����X�C�g�B��B�~����(�`�b)�H��w�����Ѫ���g�B��B�~�����X�C"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   855
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   4215
   End
   Begin VB.Label lblEntry 
      Alignment       =   1  '�a�k���
      BorderStyle     =   1  '��u�T�w
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   720
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  '�m�����
      BackColor       =   &H00808080&
      BackStyle       =   0  '�z��
      Caption         =   "����C�L"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '�m�����
      BackColor       =   &H00808080&
      BorderStyle     =   1  '��u�T�w
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Call Form_Unload(0)
End Sub

Sub DayReport(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String, OldPID As String, CurrentCount As Integer
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    
    
    SQL = "select * from product order by PID;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    

    
    
    Open TargetPath For Output As #1
        Print #1, "<table>"
    
        'show report datetime
        Body = "<tr><td>���</td><td colspan=10>" & txtCurrentDate.Text & "</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>���~"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        Body = Body & "</td></tr>"
        Print #1, Body
        
        'show every custom order per product
        product_rec.MoveFirst
        Do Until custom_rec.EOF
            'mark custom name
            CID = custom_rec.Fields.Item("CID")
            CName = custom_rec.Fields.Item("CName")
            Body = "<tr>"
            Body = Body & "<td>" & CName & "</td>"


            'show every custom order per product
            For i = 0 To Count - 1
                'query a new custom order
                selectFields = "CurrentCount,WinningCount"
                SQL = "select " & selectFields & " from [order] where [order].PID='" & PIDArray(i) & "' and CID='" & CID & "' and CurrentDate='" & txtCurrentDate.Text & "' order by [order].PID;"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
                
                'enum custom every product order count
                CurrentCount = 0
                Do Until rec1.EOF
                    If OldPID = "" Then
                        OldPID = PIDArray(i)
                        CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                    ElseIf OldPID = PIDArray(i) Then
                        CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                    Else
                        OldPID = PIDArray(i)
                        Body = Body & "<td>" & CurrentCount & "</td>"
                        CurrentCount = Val(rec1.Fields.Item("CurrentCount"))
                    End If
                    
                    rec1.MoveNext
                Loop
                Body = Body & "<td>" & CurrentCount & "</td>"
            Next
            Body = Body & "</tr>"
            Print #1, Body
            
            custom_rec.MoveNext
        Loop
        
        Print #1, "</table>"
    Close #1
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Sub WeekReport(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String, OldPID As String, CurrentCount As Integer
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    
    
    SQL = "select * from product order by PID;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    

    
    
    Open TargetPath For Output As #1
        Print #1, "<table>"
    
        'show report datetime
        Body = "<tr><td>���</td><td colspan=10>" & DateTime.DateAdd("d", -7, txtCurrentDate.Text) & "��" & txtCurrentDate.Text & "</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>���~"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        Body = Body & "</td></tr>"
        Print #1, Body
        
        'show every custom order per product
        product_rec.MoveFirst
        Do Until custom_rec.EOF
            'mark custom name
            CID = custom_rec.Fields.Item("CID")
            CName = custom_rec.Fields.Item("CName")
            Body = "<tr>"
            Body = Body & "<td>" & CName & "</td>"


            'show every custom order per product
            For i = 0 To Count - 1
                'query a new custom order
                selectFields = "CurrentCount,WinningCount"
                SQL = "select " & selectFields & " from [order] where [order].PID='" & PIDArray(i) & "' and CID='" & CID & "' and (CurrentDate>='" & Format(DateTime.DateAdd("d", -7, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd") & "' and CurrentDate<='" & txtCurrentDate.Text & "') order by [order].PID;"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
                
                'enum custom every product order count
                CurrentCount = 0
                Do Until rec1.EOF
                    If OldPID = "" Then
                        OldPID = PIDArray(i)
                        CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                    ElseIf OldPID = PIDArray(i) Then
                        CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                    Else
                        OldPID = PIDArray(i)
                        Body = Body & "<td>" & CurrentCount & "</td>"
                        CurrentCount = Val(rec1.Fields.Item("CurrentCount"))
                    End If
                    
                    rec1.MoveNext
                Loop
                Body = Body & "<td>" & CurrentCount & "</td>"
            Next
            Body = Body & "</tr>"
            Print #1, Body
            
            custom_rec.MoveNext
        Loop
        
        Print #1, "</table>"
    Close #1
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Sub MonthReport(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String, OldPID As String, CurrentCount As Integer
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    
    
    SQL = "select * from product order by PID;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    

    
    
    Open TargetPath For Output As #1
        Print #1, "<table>"
    
        'show report datetime
        Body = "<tr><td>���</td><td colspan=10>" & Format(txtCurrentDate.Text, "yyyy/MM") & "��</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>���~"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        Body = Body & "</td></tr>"
        Print #1, Body
        
        'show every custom order per product
        product_rec.MoveFirst
        Do Until custom_rec.EOF
            'mark custom name
            CID = custom_rec.Fields.Item("CID")
            CName = custom_rec.Fields.Item("CName")
            Body = "<tr>"
            Body = Body & "<td>" & CName & "</td>"


            'show every custom order per product
            For i = 0 To Count - 1
                'query a new custom order
                selectFields = "CurrentCount,WinningCount"
                SQL = "select " & selectFields & " from [order] where [order].PID='" & PIDArray(i) & "' and CID='" & CID & "' and (CurrentDate>='" & Format(txtCurrentDate.Text, "yyyy/MM/") & "01' and CurrentDate<='" & Format(txtCurrentDate.Text, "yyyy/MM/") & Date_Is_28_29_30_31(txtCurrentDate.Text) & "') order by [order].PID;"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
                
                'enum custom every product order count
                CurrentCount = 0
                Do Until rec1.EOF
                    If OldPID = "" Then
                        OldPID = PIDArray(i)
                        CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                    ElseIf OldPID = PIDArray(i) Then
                        CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                    Else
                        OldPID = PIDArray(i)
                        Body = Body & "<td>" & CurrentCount & "</td>"
                        CurrentCount = Val(rec1.Fields.Item("CurrentCount"))
                    End If
                    
                    rec1.MoveNext
                Loop
                Body = Body & "<td>" & CurrentCount & "</td>"
            Next
            Body = Body & "</tr>"
            Print #1, Body
            
            custom_rec.MoveNext
        Loop
        
        Print #1, "</table>"
    Close #1
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Sub YearReport(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String, OldPID As String, CurrentCount As Integer
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    
    
    SQL = "select * from product order by PID;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    

    
    
    Open TargetPath For Output As #1
        Print #1, "<table>"
    
        'show report datetime
        Body = "<tr><td>���</td><td colspan=10>" & Format(txtCurrentDate.Text, "yyyy") & "�~</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>���~"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        Body = Body & "</td></tr>"
        Print #1, Body
        
        'show every custom order per product
        product_rec.MoveFirst
        Do Until custom_rec.EOF
            'mark custom name
            CID = custom_rec.Fields.Item("CID")
            CName = custom_rec.Fields.Item("CName")
            Body = "<tr>"
            Body = Body & "<td>" & CName & "</td>"


            'show every custom order per product
            For i = 0 To Count - 1
                'query a new custom order
                selectFields = "CurrentCount,WinningCount"
                SQL = "select " & selectFields & " from [order] where [order].PID='" & PIDArray(i) & "' and CID='" & CID & "' and (CurrentDate>='" & Format(txtCurrentDate.Text, "yyyy/") & "01/01' and CurrentDate<='" & Format(txtCurrentDate.Text, "yyyy/") & "12/31') order by [order].PID;"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
                
                'enum custom every product order count
                CurrentCount = 0
                Do Until rec1.EOF
                    If OldPID = "" Then
                        OldPID = PIDArray(i)
                        CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                    ElseIf OldPID = PIDArray(i) Then
                        CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                    Else
                        OldPID = PIDArray(i)
                        Body = Body & "<td>" & CurrentCount & "</td>"
                        CurrentCount = Val(rec1.Fields.Item("CurrentCount"))
                    End If
                    
                    rec1.MoveNext
                Loop
                Body = Body & "<td>" & CurrentCount & "</td>"
            Next
            Body = Body & "</tr>"
            Print #1, Body
            
            custom_rec.MoveNext
        Loop
        
        Print #1, "</table>"
    Close #1
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Sub DayAccount(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String, OldPID As String, CurrentCount As Integer, WinningCount As Integer
    Dim CurrentCountAll(1024) As Long, WinningCountAll(1024) As Long
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    
    
    SQL = "select * from product order by PID;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    

    
    
    Open TargetPath For Output As #1
        Print #1, "<table>"
    
        'show report datetime
        Body = "<tr><td colspan=2>" & txtCurrentDate.Text & "</td><td>���`�p</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>���~"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "��</td>"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        Body = Body & "</td></tr>"
        Print #1, Body
        
        'show every custom order per product
        product_rec.MoveFirst
        Do Until custom_rec.EOF
            'mark custom name
            CID = custom_rec.Fields.Item("CID")
            CName = custom_rec.Fields.Item("CName")
            Body = "<tr>"
            Body = Body & "<td>" & CName & "</td>"


            'show every custom order per product
            For i = 0 To Count - 1
                'query a new custom order
                selectFields = "CurrentCount,WinningCount"
                SQL = "select " & selectFields & " from [order] where [order].PID='" & PIDArray(i) & "' and CID='" & CID & "' and CurrentDate='" & txtCurrentDate.Text & "' order by [order].PID;"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
                
                'enum custom every product order count
                CurrentCount = 0
                WinningCount = 0
                Do Until rec1.EOF
                    If OldPID = "" Then
                        OldPID = PIDArray(i)
                        CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                        WinningCount = WinningCount + Val(rec1.Fields.Item("WinningCount"))
                    ElseIf OldPID = PIDArray(i) Then
                        CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                        WinningCount = WinningCount + Val(rec1.Fields.Item("WinningCount"))
                    Else
                        OldPID = PIDArray(i)
                        Body = Body & "<td>" & CurrentCount & "</td>"
                        Body = Body & "<td>" & WinningCount & "</td>"
                        CurrentCount = Val(rec1.Fields.Item("CurrentCount"))
                        WinningCount = Val(rec1.Fields.Item("WinningCount"))
                    End If
                    
                    rec1.MoveNext
                Loop
                Body = Body & "<td>" & CurrentCount & "</td>"
                Body = Body & "<td>" & WinningCount & "</td>"
                
                'a new variable to account every product
                CurrentCountAll(i) = CurrentCountAll(i) + CurrentCount
                WinningCountAll(i) = WinningCountAll(i) + WinningCount
            Next
            Body = Body & "</tr>"
            Print #1, Body
            
            custom_rec.MoveNext
        Loop
        
        
        Body = "<tr><td>�`�p</td>"
        For i = 0 To Count - 1
            Body = Body & "<td>" & CurrentCountAll(i) & "</td>"
            Body = Body & "<td>" & WinningCountAll(i) & "</td>"
        Next
        Body = Body & "</tr>"
        Print #1, Body
        
        
        Print #1, "</table>"
    Close #1
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Sub WeekAccount(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String, OldPID As String, CurrentCount As Integer, WinningCount As Integer
    Dim CurrentCountAll(1024) As Long, WinningCountAll(1024) As Long
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    
    
    SQL = "select * from product order by PID;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    

    
    
    Open TargetPath For Output As #1
        Print #1, "<table>"
    
        'show report datetime
        Body = "<tr><td colspan=2>" & DateTime.DateAdd("d", -7, txtCurrentDate.Text) & "��" & txtCurrentDate.Text & "</td><td>�g�`�p</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>���~"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "��</td>"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        Body = Body & "</td></tr>"
        Print #1, Body
        
        'show every custom order per product
        product_rec.MoveFirst
        Do Until custom_rec.EOF
            'mark custom name
            CID = custom_rec.Fields.Item("CID")
            CName = custom_rec.Fields.Item("CName")
            Body = "<tr>"
            Body = Body & "<td>" & CName & "</td>"


            'show every custom order per product
            For i = 0 To Count - 1
                'query a new custom order
                selectFields = "CurrentCount,WinningCount"
                SQL = "select " & selectFields & " from [order] where [order].PID='" & PIDArray(i) & "' and CID='" & CID & "' and (CurrentDate>='" & Format(DateTime.DateAdd("d", -7, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd") & "' and CurrentDate<='" & txtCurrentDate.Text & "') order by [order].PID;"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
                
                'enum custom every product order count
                CurrentCount = 0
                WinningCount = 0
                Do Until rec1.EOF
                    If OldPID = "" Then
                        OldPID = PIDArray(i)
                        CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                        WinningCount = WinningCount + Val(rec1.Fields.Item("WinningCount"))
                    ElseIf OldPID = PIDArray(i) Then
                        CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                        WinningCount = WinningCount + Val(rec1.Fields.Item("WinningCount"))
                    Else
                        OldPID = PIDArray(i)
                        Body = Body & "<td>" & CurrentCount & "</td>"
                        Body = Body & "<td>" & WinningCount & "</td>"
                        CurrentCount = Val(rec1.Fields.Item("CurrentCount"))
                        WinningCount = Val(rec1.Fields.Item("WinningCount"))
                    End If
                    
                    rec1.MoveNext
                Loop
                Body = Body & "<td>" & CurrentCount & "</td>"
                Body = Body & "<td>" & WinningCount & "</td>"
                
                'a new variable to account every product
                CurrentCountAll(i) = CurrentCountAll(i) + CurrentCount
                WinningCountAll(i) = WinningCountAll(i) + WinningCount
            Next
            Body = Body & "</tr>"
            Print #1, Body
            
            custom_rec.MoveNext
        Loop
        
        
        Body = "<tr><td>�`�p</td>"
        For i = 0 To Count - 1
            Body = Body & "<td>" & CurrentCountAll(i) & "</td>"
            Body = Body & "<td>" & WinningCountAll(i) & "</td>"
        Next
        Body = Body & "</tr>"
        Print #1, Body
        
        
        Print #1, "</table>"
    Close #1
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Sub MonthAccount(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String, OldPID As String, CurrentCount As Integer, WinningCount As Integer
    Dim CurrentCountAll(1024) As Long, WinningCountAll(1024) As Long
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    
    
    SQL = "select * from product order by PID;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    

    
    
    Open TargetPath For Output As #1
        Print #1, "<table>"
    
        'show report datetime
        Body = "<tr><td colspan=2>" & Format(txtCurrentDate.Text, "yyyy/MM") & "</td><td>���`�p</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>���~"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "��</td>"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        Body = Body & "</td></tr>"
        Print #1, Body
        
        'show every custom order per product
        product_rec.MoveFirst
        Do Until custom_rec.EOF
            'mark custom name
            CID = custom_rec.Fields.Item("CID")
            CName = custom_rec.Fields.Item("CName")
            Body = "<tr>"
            Body = Body & "<td>" & CName & "</td>"


            'show every custom order per product
            For i = 0 To Count - 1
                'query a new custom order
                selectFields = "CurrentCount,WinningCount"
                SQL = "select " & selectFields & " from [order] where [order].PID='" & PIDArray(i) & "' and CID='" & CID & "' and (CurrentDate>='" & Format(txtCurrentDate.Text, "yyyy/MM/") & "01' and CurrentDate<='" & Format(txtCurrentDate.Text, "yyyy/MM/") & Date_Is_28_29_30_31(txtCurrentDate.Text) & "') order by [order].PID;"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
                
                'enum custom every product order count
                CurrentCount = 0
                WinningCount = 0
                Do Until rec1.EOF
                    If OldPID = "" Then
                        OldPID = PIDArray(i)
                        CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                        WinningCount = WinningCount + Val(rec1.Fields.Item("WinningCount"))
                    ElseIf OldPID = PIDArray(i) Then
                        CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                        WinningCount = WinningCount + Val(rec1.Fields.Item("WinningCount"))
                    Else
                        OldPID = PIDArray(i)
                        Body = Body & "<td>" & CurrentCount & "</td>"
                        Body = Body & "<td>" & WinningCount & "</td>"
                        CurrentCount = Val(rec1.Fields.Item("CurrentCount"))
                        WinningCount = Val(rec1.Fields.Item("WinningCount"))
                    End If
                    
                    rec1.MoveNext
                Loop
                Body = Body & "<td>" & CurrentCount & "</td>"
                Body = Body & "<td>" & WinningCount & "</td>"
                
                'a new variable to account every product
                CurrentCountAll(i) = CurrentCountAll(i) + CurrentCount
                WinningCountAll(i) = WinningCountAll(i) + WinningCount
            Next
            Body = Body & "</tr>"
            Print #1, Body
            
            custom_rec.MoveNext
        Loop
        
        
        Body = "<tr><td>�`�p</td>"
        For i = 0 To Count - 1
            Body = Body & "<td>" & CurrentCountAll(i) & "</td>"
            Body = Body & "<td>" & WinningCountAll(i) & "</td>"
        Next
        Body = Body & "</tr>"
        Print #1, Body
        
        
        Print #1, "</table>"
    Close #1
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Sub YearAccount(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String, OldPID As String, CurrentCount As Integer, WinningCount As Integer
    Dim CurrentCountAll(1024) As Long, WinningCountAll(1024) As Long
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    
    
    SQL = "select * from product order by PID;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    

    
    
    Open TargetPath For Output As #1
        Print #1, "<table>"
    
        'show report datetime
        Body = "<tr><td colspan=2>" & Format(txtCurrentDate.Text, "yyyy") & "</td><td>�~�`�p</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>���~"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "��</td>"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        Body = Body & "</td></tr>"
        Print #1, Body
        
        'show every custom order per product
        product_rec.MoveFirst
        Do Until custom_rec.EOF
            'mark custom name
            CID = custom_rec.Fields.Item("CID")
            CName = custom_rec.Fields.Item("CName")
            Body = "<tr>"
            Body = Body & "<td>" & CName & "</td>"


            'show every custom order per product
            For i = 0 To Count - 1
                'query a new custom order
                selectFields = "CurrentCount,WinningCount"
                SQL = "select " & selectFields & " from [order] where [order].PID='" & PIDArray(i) & "' and CID='" & CID & "' and (CurrentDate>='" & Format(txtCurrentDate.Text, "yyyy/") & "01/01' and CurrentDate<='" & Format(txtCurrentDate.Text, "yyyy/") & "12/31') order by [order].PID;"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
                
                'enum custom every product order count
                CurrentCount = 0
                WinningCount = 0
                Do Until rec1.EOF
                    If OldPID = "" Then
                        OldPID = PIDArray(i)
                        CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                        WinningCount = WinningCount + Val(rec1.Fields.Item("WinningCount"))
                    ElseIf OldPID = PIDArray(i) Then
                        CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                        WinningCount = WinningCount + Val(rec1.Fields.Item("WinningCount"))
                    Else
                        OldPID = PIDArray(i)
                        Body = Body & "<td>" & CurrentCount & "</td>"
                        Body = Body & "<td>" & WinningCount & "</td>"
                        CurrentCount = Val(rec1.Fields.Item("CurrentCount"))
                        WinningCount = Val(rec1.Fields.Item("WinningCount"))
                    End If
                    
                    rec1.MoveNext
                Loop
                Body = Body & "<td>" & CurrentCount & "</td>"
                Body = Body & "<td>" & WinningCount & "</td>"
                
                'a new variable to account every product
                CurrentCountAll(i) = CurrentCountAll(i) + CurrentCount
                WinningCountAll(i) = WinningCountAll(i) + WinningCount
            Next
            Body = Body & "</tr>"
            Print #1, Body
            
            custom_rec.MoveNext
        Loop
        
        
        Body = "<tr><td>�`�p</td>"
        For i = 0 To Count - 1
            Body = Body & "<td>" & CurrentCountAll(i) & "</td>"
            Body = Body & "<td>" & WinningCountAll(i) & "</td>"
        Next
        Body = Body & "</tr>"
        Print #1, Body
        
        
        Print #1, "</table>"
    Close #1
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub


Sub FourKDayReport(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String, OldPID As String, CurrentCount As Integer
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    
    
    SQL = "select * from product where PName like '%4K%' order by PID;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    

    
    
    Open TargetPath For Output As #1
        Print #1, "<table>"
    
        'show report datetime
        Body = "<tr><td>���</td><td colspan=10>" & txtCurrentDate.Text & "</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>���~"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        Body = Body & "</td></tr>"
        Print #1, Body
        
        'show every custom order per product
        product_rec.MoveFirst
        Do Until custom_rec.EOF
            'mark custom name
            CID = custom_rec.Fields.Item("CID")
            CName = custom_rec.Fields.Item("CName")
            Body = "<tr>"
            Body = Body & "<td>" & CName & "</td>"


            'show every custom order per product
            For i = 0 To Count - 1
                'query a new custom order
                selectFields = "CurrentCount,WinningCount"
                SQL = "select " & selectFields & " from [order] where [order].PID='" & PIDArray(i) & "' and CID='" & CID & "' and CurrentDate='" & txtCurrentDate.Text & "' order by [order].PID;"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
                
                'enum custom every product order count
                CurrentCount = 0
                Do Until rec1.EOF
                    If OldPID = "" Then
                        OldPID = PIDArray(i)
                        CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                    ElseIf OldPID = PIDArray(i) Then
                        CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                    Else
                        OldPID = PIDArray(i)
                        Body = Body & "<td>" & CurrentCount & "</td>"
                        CurrentCount = Val(rec1.Fields.Item("CurrentCount"))
                    End If
                    
                    rec1.MoveNext
                Loop
                Body = Body & "<td>" & CurrentCount & "</td>"
            Next
            Body = Body & "</tr>"
            Print #1, Body
            
            custom_rec.MoveNext
        Loop
        
        Print #1, "</table>"
    Close #1
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Sub FourKWeekReport(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String, OldPID As String, CurrentCount As Integer
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    
    
    SQL = "select * from product where PName like '%4K%' order by PID;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    

    
    
    Open TargetPath For Output As #1
        Print #1, "<table>"
    
        'show report datetime
        Body = "<tr><td>���</td><td colspan=10>" & DateTime.DateAdd("d", -7, txtCurrentDate.Text) & "��" & txtCurrentDate.Text & "</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>���~"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        Body = Body & "</td></tr>"
        Print #1, Body
        
        'show every custom order per product
        product_rec.MoveFirst
        Do Until custom_rec.EOF
            'mark custom name
            CID = custom_rec.Fields.Item("CID")
            CName = custom_rec.Fields.Item("CName")
            Body = "<tr>"
            Body = Body & "<td>" & CName & "</td>"


            'show every custom order per product
            For i = 0 To Count - 1
                'query a new custom order
                selectFields = "CurrentCount,WinningCount"
                SQL = "select " & selectFields & " from [order] where [order].PID='" & PIDArray(i) & "' and CID='" & CID & "' and (CurrentDate>='" & Format(DateTime.DateAdd("d", -7, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd") & "' and CurrentDate<='" & txtCurrentDate.Text & "') order by [order].PID;"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
                
                'enum custom every product order count
                CurrentCount = 0
                Do Until rec1.EOF
                    If OldPID = "" Then
                        OldPID = PIDArray(i)
                        CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                    ElseIf OldPID = PIDArray(i) Then
                        CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                    Else
                        OldPID = PIDArray(i)
                        Body = Body & "<td>" & CurrentCount & "</td>"
                        CurrentCount = Val(rec1.Fields.Item("CurrentCount"))
                    End If
                    
                    rec1.MoveNext
                Loop
                Body = Body & "<td>" & CurrentCount & "</td>"
            Next
            Body = Body & "</tr>"
            Print #1, Body
            
            custom_rec.MoveNext
        Loop
        
        Print #1, "</table>"
    Close #1
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Sub FourKMonthReport(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String, OldPID As String, CurrentCount As Integer
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    
    
    SQL = "select * from product where PName like '%4K%' order by PID;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    

    
    
    Open TargetPath For Output As #1
        Print #1, "<table>"
    
        'show report datetime
        Body = "<tr><td>���</td><td colspan=10>" & Format(txtCurrentDate.Text, "yyyy/MM") & "��</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>���~"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        Body = Body & "</td></tr>"
        Print #1, Body
        
        'show every custom order per product
        product_rec.MoveFirst
        Do Until custom_rec.EOF
            'mark custom name
            CID = custom_rec.Fields.Item("CID")
            CName = custom_rec.Fields.Item("CName")
            Body = "<tr>"
            Body = Body & "<td>" & CName & "</td>"


            'show every custom order per product
            For i = 0 To Count - 1
                'query a new custom order
                selectFields = "CurrentCount,WinningCount"
                SQL = "select " & selectFields & " from [order] where [order].PID='" & PIDArray(i) & "' and CID='" & CID & "' and (CurrentDate>='" & Format(txtCurrentDate.Text, "yyyy/MM/") & "01' and CurrentDate<='" & Format(txtCurrentDate.Text, "yyyy/MM/") & Date_Is_28_29_30_31(txtCurrentDate.Text) & "') order by [order].PID;"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
                
                'enum custom every product order count
                CurrentCount = 0
                Do Until rec1.EOF
                    If OldPID = "" Then
                        OldPID = PIDArray(i)
                        CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                    ElseIf OldPID = PIDArray(i) Then
                        CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                    Else
                        OldPID = PIDArray(i)
                        Body = Body & "<td>" & CurrentCount & "</td>"
                        CurrentCount = Val(rec1.Fields.Item("CurrentCount"))
                    End If
                    
                    rec1.MoveNext
                Loop
                Body = Body & "<td>" & CurrentCount & "</td>"
            Next
            Body = Body & "</tr>"
            Print #1, Body
            
            custom_rec.MoveNext
        Loop
        
        Print #1, "</table>"
    Close #1
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Sub FourKYearReport(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String, OldPID As String, CurrentCount As Integer
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    
    
    SQL = "select * from product where PName like '%4K%'order by PID;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    

    
    
    Open TargetPath For Output As #1
        Print #1, "<table>"
    
        'show report datetime
        Body = "<tr><td>���</td><td colspan=10>" & Format(txtCurrentDate.Text, "yyyy") & "�~</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>���~"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        Body = Body & "</td></tr>"
        Print #1, Body
        
        'show every custom order per product
        product_rec.MoveFirst
        Do Until custom_rec.EOF
            'mark custom name
            CID = custom_rec.Fields.Item("CID")
            CName = custom_rec.Fields.Item("CName")
            Body = "<tr>"
            Body = Body & "<td>" & CName & "</td>"


            'show every custom order per product
            For i = 0 To Count - 1
                'query a new custom order
                selectFields = "CurrentCount,WinningCount"
                SQL = "select " & selectFields & " from [order] where [order].PID='" & PIDArray(i) & "' and CID='" & CID & "' and (CurrentDate>='" & Format(txtCurrentDate.Text, "yyyy/") & "01/01' and CurrentDate<='" & Format(txtCurrentDate.Text, "yyyy/") & "12/31') order by [order].PID;"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
                
                'enum custom every product order count
                CurrentCount = 0
                Do Until rec1.EOF
                    If OldPID = "" Then
                        OldPID = PIDArray(i)
                        CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                    ElseIf OldPID = PIDArray(i) Then
                        CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                    Else
                        OldPID = PIDArray(i)
                        Body = Body & "<td>" & CurrentCount & "</td>"
                        CurrentCount = Val(rec1.Fields.Item("CurrentCount"))
                    End If
                    
                    rec1.MoveNext
                Loop
                Body = Body & "<td>" & CurrentCount & "</td>"
            Next
            Body = Body & "</tr>"
            Print #1, Body
            
            custom_rec.MoveNext
        Loop
        
        Print #1, "</table>"
    Close #1
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Sub FourKDayAccount(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String, OldPID As String, CurrentCount As Integer, WinningCount As Integer
    Dim CurrentCountAll(1024) As Long, WinningCountAll(1024) As Long
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    
    
    SQL = "select * from product where PName like '%4K%' order by PID;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    

    
    
    Open TargetPath For Output As #1
        Print #1, "<table>"
    
        'show report datetime
        Body = "<tr><td colspan=2>" & txtCurrentDate.Text & "</td><td>���`�p</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>���~"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "��</td>"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        Body = Body & "</td></tr>"
        Print #1, Body
        
        'show every custom order per product
        product_rec.MoveFirst
        Do Until custom_rec.EOF
            'mark custom name
            CID = custom_rec.Fields.Item("CID")
            CName = custom_rec.Fields.Item("CName")
            Body = "<tr>"
            Body = Body & "<td>" & CName & "</td>"


            'show every custom order per product
            For i = 0 To Count - 1
                'query a new custom order
                selectFields = "CurrentCount,WinningCount"
                SQL = "select " & selectFields & " from [order] where [order].PID='" & PIDArray(i) & "' and CID='" & CID & "' and CurrentDate='" & txtCurrentDate.Text & "' order by [order].PID;"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
                
                'enum custom every product order count
                CurrentCount = 0
                WinningCount = 0
                Do Until rec1.EOF
                    If OldPID = "" Then
                        OldPID = PIDArray(i)
                        CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                        WinningCount = WinningCount + Val(rec1.Fields.Item("WinningCount"))
                    ElseIf OldPID = PIDArray(i) Then
                        CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                        WinningCount = WinningCount + Val(rec1.Fields.Item("WinningCount"))
                    Else
                        OldPID = PIDArray(i)
                        Body = Body & "<td>" & CurrentCount & "</td>"
                        Body = Body & "<td>" & WinningCount & "</td>"
                        CurrentCount = Val(rec1.Fields.Item("CurrentCount"))
                        WinningCount = Val(rec1.Fields.Item("WinningCount"))
                    End If
                    
                    rec1.MoveNext
                Loop
                Body = Body & "<td>" & CurrentCount & "</td>"
                Body = Body & "<td>" & WinningCount & "</td>"
                
                'a new variable to account every product
                CurrentCountAll(i) = CurrentCountAll(i) + CurrentCount
                WinningCountAll(i) = WinningCountAll(i) + WinningCount
            Next
            Body = Body & "</tr>"
            Print #1, Body
            
            custom_rec.MoveNext
        Loop
        
        
        Body = "<tr><td>�`�p</td>"
        For i = 0 To Count - 1
            Body = Body & "<td>" & CurrentCountAll(i) & "</td>"
            Body = Body & "<td>" & WinningCountAll(i) & "</td>"
        Next
        Body = Body & "</tr>"
        Print #1, Body
        
        
        Print #1, "</table>"
    Close #1
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Sub FourKWeekAccount(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String, OldPID As String, CurrentCount As Integer, WinningCount As Integer
    Dim CurrentCountAll(1024) As Long, WinningCountAll(1024) As Long
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    
    
    SQL = "select * from product where PName like '%4K%' order by PID;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    

    
    
    Open TargetPath For Output As #1
        Print #1, "<table>"
    
        'show report datetime
        Body = "<tr><td colspan=2>" & DateTime.DateAdd("d", -7, txtCurrentDate.Text) & "��" & txtCurrentDate.Text & "</td><td>�g�`�p</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>���~"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "��</td>"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        Body = Body & "</td></tr>"
        Print #1, Body
        
        'show every custom order per product
        product_rec.MoveFirst
        Do Until custom_rec.EOF
            'mark custom name
            CID = custom_rec.Fields.Item("CID")
            CName = custom_rec.Fields.Item("CName")
            Body = "<tr>"
            Body = Body & "<td>" & CName & "</td>"


            'show every custom order per product
            For i = 0 To Count - 1
                'query a new custom order
                selectFields = "CurrentCount,WinningCount"
                SQL = "select " & selectFields & " from [order] where [order].PID='" & PIDArray(i) & "' and CID='" & CID & "' and (CurrentDate>='" & Format(DateTime.DateAdd("d", -7, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd") & "' and CurrentDate<='" & txtCurrentDate.Text & "') order by [order].PID;"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
                
                'enum custom every product order count
                CurrentCount = 0
                WinningCount = 0
                Do Until rec1.EOF
                    If OldPID = "" Then
                        OldPID = PIDArray(i)
                        CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                        WinningCount = WinningCount + Val(rec1.Fields.Item("WinningCount"))
                    ElseIf OldPID = PIDArray(i) Then
                        CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                        WinningCount = WinningCount + Val(rec1.Fields.Item("WinningCount"))
                    Else
                        OldPID = PIDArray(i)
                        Body = Body & "<td>" & CurrentCount & "</td>"
                        Body = Body & "<td>" & WinningCount & "</td>"
                        CurrentCount = Val(rec1.Fields.Item("CurrentCount"))
                        WinningCount = Val(rec1.Fields.Item("WinningCount"))
                    End If
                    
                    rec1.MoveNext
                Loop
                Body = Body & "<td>" & CurrentCount & "</td>"
                Body = Body & "<td>" & WinningCount & "</td>"
                
                'a new variable to account every product
                CurrentCountAll(i) = CurrentCountAll(i) + CurrentCount
                WinningCountAll(i) = WinningCountAll(i) + WinningCount
            Next
            Body = Body & "</tr>"
            Print #1, Body
            
            custom_rec.MoveNext
        Loop
        
        
        Body = "<tr><td>�`�p</td>"
        For i = 0 To Count - 1
            Body = Body & "<td>" & CurrentCountAll(i) & "</td>"
            Body = Body & "<td>" & WinningCountAll(i) & "</td>"
        Next
        Body = Body & "</tr>"
        Print #1, Body
        
        
        Print #1, "</table>"
    Close #1
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Sub FourKMonthAccount(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String, OldPID As String, CurrentCount As Integer, WinningCount As Integer
    Dim CurrentCountAll(1024) As Long, WinningCountAll(1024) As Long
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    
    
    SQL = "select * from product where PName like '%4K%' order by PID;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    

    
    
    Open TargetPath For Output As #1
        Print #1, "<table>"
    
        'show report datetime
        Body = "<tr><td colspan=2>" & Format(txtCurrentDate.Text, "yyyy/MM") & "</td><td>���`�p</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>���~"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "��</td>"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        Body = Body & "</td></tr>"
        Print #1, Body
        
        'show every custom order per product
        product_rec.MoveFirst
        Do Until custom_rec.EOF
            'mark custom name
            CID = custom_rec.Fields.Item("CID")
            CName = custom_rec.Fields.Item("CName")
            Body = "<tr>"
            Body = Body & "<td>" & CName & "</td>"


            'show every custom order per product
            For i = 0 To Count - 1
                'query a new custom order
                selectFields = "CurrentCount,WinningCount"
                SQL = "select " & selectFields & " from [order] where [order].PID='" & PIDArray(i) & "' and CID='" & CID & "' and (CurrentDate>='" & Format(txtCurrentDate.Text, "yyyy/MM/") & "01' and CurrentDate<='" & Format(txtCurrentDate.Text, "yyyy/MM/") & Date_Is_28_29_30_31(txtCurrentDate.Text) & "') order by [order].PID;"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
                
                'enum custom every product order count
                CurrentCount = 0
                WinningCount = 0
                Do Until rec1.EOF
                    If OldPID = "" Then
                        OldPID = PIDArray(i)
                        CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                        WinningCount = WinningCount + Val(rec1.Fields.Item("WinningCount"))
                    ElseIf OldPID = PIDArray(i) Then
                        CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                        WinningCount = WinningCount + Val(rec1.Fields.Item("WinningCount"))
                    Else
                        OldPID = PIDArray(i)
                        Body = Body & "<td>" & CurrentCount & "</td>"
                        Body = Body & "<td>" & WinningCount & "</td>"
                        CurrentCount = Val(rec1.Fields.Item("CurrentCount"))
                        WinningCount = Val(rec1.Fields.Item("WinningCount"))
                    End If
                    
                    rec1.MoveNext
                Loop
                Body = Body & "<td>" & CurrentCount & "</td>"
                Body = Body & "<td>" & WinningCount & "</td>"
                
                'a new variable to account every product
                CurrentCountAll(i) = CurrentCountAll(i) + CurrentCount
                WinningCountAll(i) = WinningCountAll(i) + WinningCount
            Next
            Body = Body & "</tr>"
            Print #1, Body
            
            custom_rec.MoveNext
        Loop
        
        
        Body = "<tr><td>�`�p</td>"
        For i = 0 To Count - 1
            Body = Body & "<td>" & CurrentCountAll(i) & "</td>"
            Body = Body & "<td>" & WinningCountAll(i) & "</td>"
        Next
        Body = Body & "</tr>"
        Print #1, Body
        
        
        Print #1, "</table>"
    Close #1
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Sub FourKYearAccount(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String, OldPID As String, CurrentCount As Integer, WinningCount As Integer
    Dim CurrentCountAll(1024) As Long, WinningCountAll(1024) As Long
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    
    
    SQL = "select * from product where PName like '%4K%'order by PID;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    

    
    
    Open TargetPath For Output As #1
        Print #1, "<table>"
    
        'show report datetime
        Body = "<tr><td colspan=2>" & Format(txtCurrentDate.Text, "yyyy") & "</td><td>�~�`�p</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>���~"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "��</td>"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        Body = Body & "</td></tr>"
        Print #1, Body
        
        'show every custom order per product
        product_rec.MoveFirst
        Do Until custom_rec.EOF
            'mark custom name
            CID = custom_rec.Fields.Item("CID")
            CName = custom_rec.Fields.Item("CName")
            Body = "<tr>"
            Body = Body & "<td>" & CName & "</td>"


            'show every custom order per product
            For i = 0 To Count - 1
                'query a new custom order
                selectFields = "CurrentCount,WinningCount"
                SQL = "select " & selectFields & " from [order] where [order].PID='" & PIDArray(i) & "' and CID='" & CID & "' and (CurrentDate>='" & Format(txtCurrentDate.Text, "yyyy/") & "01/01' and CurrentDate<='" & Format(txtCurrentDate.Text, "yyyy/") & "12/31') order by [order].PID;"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
                
                'enum custom every product order count
                CurrentCount = 0
                WinningCount = 0
                Do Until rec1.EOF
                    If OldPID = "" Then
                        OldPID = PIDArray(i)
                        CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                        WinningCount = WinningCount + Val(rec1.Fields.Item("WinningCount"))
                    ElseIf OldPID = PIDArray(i) Then
                        CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                        WinningCount = WinningCount + Val(rec1.Fields.Item("WinningCount"))
                    Else
                        OldPID = PIDArray(i)
                        Body = Body & "<td>" & CurrentCount & "</td>"
                        Body = Body & "<td>" & WinningCount & "</td>"
                        CurrentCount = Val(rec1.Fields.Item("CurrentCount"))
                        WinningCount = Val(rec1.Fields.Item("WinningCount"))
                    End If
                    
                    rec1.MoveNext
                Loop
                Body = Body & "<td>" & CurrentCount & "</td>"
                Body = Body & "<td>" & WinningCount & "</td>"
                
                'a new variable to account every product
                CurrentCountAll(i) = CurrentCountAll(i) + CurrentCount
                WinningCountAll(i) = WinningCountAll(i) + WinningCount
            Next
            Body = Body & "</tr>"
            Print #1, Body
            
            custom_rec.MoveNext
        Loop
        
        
        Body = "<tr><td>�`�p</td>"
        For i = 0 To Count - 1
            Body = Body & "<td>" & CurrentCountAll(i) & "</td>"
            Body = Body & "<td>" & WinningCountAll(i) & "</td>"
        Next
        Body = Body & "</tr>"
        Print #1, Body
        
        
        Print #1, "</table>"
    Close #1
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Private Sub cmdConfirm_Click()
    If txtCurrentDate.Text = "" Then
        MsgBox "�Х���ܭn�C�L���ɶ��I"
    Else
        Dim TargetPath As String
        
        TargetPath = App.Path
        If Right(TargetPath, 1) <> "\" Then
            TargetPath = TargetPath & "\report\" & Format(txtCurrentDate.Text, "yyyy") & "\" & Format(txtCurrentDate.Text, "mm") & "\"
        Else
            TargetPath = TargetPath & "report\" & Format(txtCurrentDate.Text, "yyyy") & "\" & Format(txtCurrentDate.Text, "mm") & "\"
        End If
        Call CreatePath(TargetPath)
        
        Select Case basVariable.Parameter
        Case "DayReport"
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyyMMdd") & "_�����.xls"
            Call DayReport(TargetPath)
        Case "WeekReport"
            TargetPath = TargetPath & Format(DateTime.DateAdd("d", -7, txtCurrentDate.Text), "yyyyMMdd") & "��" & Format(txtCurrentDate.Text, "yyyyMMdd") & "_�g����.xls"
            Call WeekReport(TargetPath)
        Case "MonthReport"
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyyMM") & "_�����.xls"
            Call MonthReport(TargetPath)
        Case "YearReport"
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyy") & "_�~����.xls"
            Call YearReport(TargetPath)
        Case "DayAccount"
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyyMMdd") & "_���`�b.xls"
            Call DayAccount(TargetPath)
        Case "WeekAccount"
            TargetPath = TargetPath & Format(DateTime.DateAdd("d", -7, txtCurrentDate.Text), "yyyyMMdd") & "��" & Format(txtCurrentDate.Text, "yyyyMMdd") & "_�g�`�b.xls"
            Call WeekAccount(TargetPath)
        Case "MonthAccount"
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyyMM") & "_���`�b.xls"
            Call MonthAccount(TargetPath)
        Case "YearAccount"
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyy") & "_�~�`�b.xls"
            Call YearAccount(TargetPath)
        Case "FourKDayReport"
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyyMMdd") & "_4K�����.xls"
            Call FourKDayReport(TargetPath)
        Case "FourKWeekReport"
            TargetPath = TargetPath & Format(DateTime.DateAdd("d", -7, txtCurrentDate.Text), "yyyyMMdd") & "��" & Format(txtCurrentDate.Text, "yyyyMMdd") & "_4K�g����.xls"
            Call FourKWeekReport(TargetPath)
        Case "FourKMonthReport"
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyyMM") & "_4K�����.xls"
            Call FourKMonthReport(TargetPath)
        Case "FourKYearReport"
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyy") & "_4K�~����.xls"
            Call FourKYearReport(TargetPath)
        Case "FourKDayAccount"
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyyMMdd") & "_4K���`�b.xls"
            Call FourKDayAccount(TargetPath)
        Case "FourKWeekAccount"
            TargetPath = TargetPath & Format(DateTime.DateAdd("d", -7, txtCurrentDate.Text), "yyyyMMdd") & "��" & Format(txtCurrentDate.Text, "yyyyMMdd") & "_4K�g�`�b.xls"
            Call FourKWeekAccount(TargetPath)
        Case "FourKMonthAccount"
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyyMM") & "_4K���`�b.xls"
            Call FourKMonthAccount(TargetPath)
        Case "FourKYearAccount"
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyy") & "_4K�~�`�b.xls"
            Call FourKYearAccount(TargetPath)
        End Select
    End If
    
    MsgBox "�w��X�����" & TargetPath & "�I"
End Sub

Private Sub dtpCurrentDate_CloseUp()
    txtCurrentDate.Text = Format(dtpCurrentDate.Value, "yyyy/MM/dd")
End Sub

Private Sub Form_Load()
    Select Case basVariable.Parameter
    Case "DayReport"
        Label1(0).Caption = "�����C�L"
    Case "WeekReport"
        Label1(0).Caption = "�g����C�L"
    Case "MonthReport"
        Label1(0).Caption = "�����C�L"
    Case "YearReport"
        Label1(0).Caption = "�~����C�L"
    Case "DayAccount"
        Label1(0).Caption = "���`�b�C�L"
    Case "WeekAccount"
        Label1(0).Caption = "�g�`�b�C�L"
    Case "MonthAccount"
        Label1(0).Caption = "���`�b�C�L"
    Case "YearAccount"
        Label1(0).Caption = "�~�`�b�C�L"
    Case "FourKDayReport"
        Label1(0).Caption = "4K�����C�L"
    Case "FourKWeekReport"
        Label1(0).Caption = "4K�g����C�L"
    Case "FourKMonthReport"
        Label1(0).Caption = "4K�����C�L"
    Case "FourKYearReport"
        Label1(0).Caption = "4K�~����C�L"
    Case "FourKDayAccount"
        Label1(0).Caption = "4K���`�b�C�L"
    Case "FourKWeekAccount"
        Label1(0).Caption = "4K�g�`�b�C�L"
    Case "FourKMonthAccount"
        Label1(0).Caption = "4K���`�b�C�L"
    Case "FourKYearAccount"
        Label1(0).Caption = "4K�~�`�b�C�L"
    End Select
    
    
    dtpCurrentDate.Value = Format(DateTime.Now, "yyyy/MM/dd")
    txtCurrentDate.Text = Format(DateTime.Now, "yyyy/MM/dd")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmProve.Show
    Unload Me
End Sub
