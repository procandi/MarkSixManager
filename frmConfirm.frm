VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmConfirm 
   BorderStyle     =   1  '單線固定
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
         Name            =   "新細明體"
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
      Caption         =   "&N 取　消"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton cmdConfirm 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&O 確　定"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  '圖片外觀
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
         Name            =   "新細明體"
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
      Alignment       =   2  '置中對齊
      BackColor       =   &H00808080&
      BackStyle       =   0  '透明
      Caption         =   "日報表(總帳)以選定的那天報表輸出。週、月、年報表(總帳)以選定的那天的當週、月、年報表輸出。"
      BeginProperty Font 
         Name            =   "新細明體"
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
      Alignment       =   1  '靠右對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "交易日期"
      BeginProperty Font 
         Name            =   "新細明體"
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
      Alignment       =   2  '置中對齊
      BackColor       =   &H00808080&
      BackStyle       =   0  '透明
      Caption         =   "報表列印"
      BeginProperty Font 
         Name            =   "新細明體"
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
      Alignment       =   2  '置中對齊
      BackColor       =   &H00808080&
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
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
        Body = "<tr><td>日期</td><td colspan=10>" & txtCurrentDate.Text & "</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>產品"
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
        Body = "<tr><td>日期</td><td colspan=10>" & DateTime.DateAdd("d", -7, txtCurrentDate.Text) & "至" & txtCurrentDate.Text & "</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>產品"
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
        Body = "<tr><td>日期</td><td colspan=10>" & Format(txtCurrentDate.Text, "yyyy/MM") & "月</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>產品"
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
        Body = "<tr><td>日期</td><td colspan=10>" & Format(txtCurrentDate.Text, "yyyy") & "年</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>產品"
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
        Body = "<tr><td colspan=2>" & txtCurrentDate.Text & "</td><td>日總計</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>產品"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "中</td>"
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
        
        
        Body = "<tr><td>總計</td>"
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
        Body = "<tr><td colspan=2>" & DateTime.DateAdd("d", -7, txtCurrentDate.Text) & "至" & txtCurrentDate.Text & "</td><td>週總計</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>產品"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "中</td>"
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
        
        
        Body = "<tr><td>總計</td>"
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
        Body = "<tr><td colspan=2>" & Format(txtCurrentDate.Text, "yyyy/MM") & "</td><td>月總計</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>產品"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "中</td>"
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
        
        
        Body = "<tr><td>總計</td>"
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
        Body = "<tr><td colspan=2>" & Format(txtCurrentDate.Text, "yyyy") & "</td><td>年總計</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>產品"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "中</td>"
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
        
        
        Body = "<tr><td>總計</td>"
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
        Body = "<tr><td>日期</td><td colspan=10>" & txtCurrentDate.Text & "</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>產品"
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
        Body = "<tr><td>日期</td><td colspan=10>" & DateTime.DateAdd("d", -7, txtCurrentDate.Text) & "至" & txtCurrentDate.Text & "</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>產品"
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
        Body = "<tr><td>日期</td><td colspan=10>" & Format(txtCurrentDate.Text, "yyyy/MM") & "月</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>產品"
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
        Body = "<tr><td>日期</td><td colspan=10>" & Format(txtCurrentDate.Text, "yyyy") & "年</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>產品"
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
        Body = "<tr><td colspan=2>" & txtCurrentDate.Text & "</td><td>日總計</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>產品"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "中</td>"
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
        
        
        Body = "<tr><td>總計</td>"
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
        Body = "<tr><td colspan=2>" & DateTime.DateAdd("d", -7, txtCurrentDate.Text) & "至" & txtCurrentDate.Text & "</td><td>週總計</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>產品"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "中</td>"
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
        
        
        Body = "<tr><td>總計</td>"
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
        Body = "<tr><td colspan=2>" & Format(txtCurrentDate.Text, "yyyy/MM") & "</td><td>月總計</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>產品"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "中</td>"
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
        
        
        Body = "<tr><td>總計</td>"
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
        Body = "<tr><td colspan=2>" & Format(txtCurrentDate.Text, "yyyy") & "</td><td>年總計</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>產品"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "中</td>"
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
        
        
        Body = "<tr><td>總計</td>"
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
        MsgBox "請先選擇要列印的時間！"
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
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyyMMdd") & "_日報表.xls"
            Call DayReport(TargetPath)
        Case "WeekReport"
            TargetPath = TargetPath & Format(DateTime.DateAdd("d", -7, txtCurrentDate.Text), "yyyyMMdd") & "至" & Format(txtCurrentDate.Text, "yyyyMMdd") & "_週報表.xls"
            Call WeekReport(TargetPath)
        Case "MonthReport"
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyyMM") & "_月報表.xls"
            Call MonthReport(TargetPath)
        Case "YearReport"
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyy") & "_年報表.xls"
            Call YearReport(TargetPath)
        Case "DayAccount"
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyyMMdd") & "_日總帳.xls"
            Call DayAccount(TargetPath)
        Case "WeekAccount"
            TargetPath = TargetPath & Format(DateTime.DateAdd("d", -7, txtCurrentDate.Text), "yyyyMMdd") & "至" & Format(txtCurrentDate.Text, "yyyyMMdd") & "_週總帳.xls"
            Call WeekAccount(TargetPath)
        Case "MonthAccount"
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyyMM") & "_月總帳.xls"
            Call MonthAccount(TargetPath)
        Case "YearAccount"
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyy") & "_年總帳.xls"
            Call YearAccount(TargetPath)
        Case "FourKDayReport"
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyyMMdd") & "_4K日報表.xls"
            Call FourKDayReport(TargetPath)
        Case "FourKWeekReport"
            TargetPath = TargetPath & Format(DateTime.DateAdd("d", -7, txtCurrentDate.Text), "yyyyMMdd") & "至" & Format(txtCurrentDate.Text, "yyyyMMdd") & "_4K週報表.xls"
            Call FourKWeekReport(TargetPath)
        Case "FourKMonthReport"
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyyMM") & "_4K月報表.xls"
            Call FourKMonthReport(TargetPath)
        Case "FourKYearReport"
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyy") & "_4K年報表.xls"
            Call FourKYearReport(TargetPath)
        Case "FourKDayAccount"
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyyMMdd") & "_4K日總帳.xls"
            Call FourKDayAccount(TargetPath)
        Case "FourKWeekAccount"
            TargetPath = TargetPath & Format(DateTime.DateAdd("d", -7, txtCurrentDate.Text), "yyyyMMdd") & "至" & Format(txtCurrentDate.Text, "yyyyMMdd") & "_4K週總帳.xls"
            Call FourKWeekAccount(TargetPath)
        Case "FourKMonthAccount"
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyyMM") & "_4K月總帳.xls"
            Call FourKMonthAccount(TargetPath)
        Case "FourKYearAccount"
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyy") & "_4K年總帳.xls"
            Call FourKYearAccount(TargetPath)
        End Select
    End If
    
    MsgBox "已輸出報表至" & TargetPath & "！"
End Sub

Private Sub dtpCurrentDate_CloseUp()
    txtCurrentDate.Text = Format(dtpCurrentDate.Value, "yyyy/MM/dd")
End Sub

Private Sub Form_Load()
    Select Case basVariable.Parameter
    Case "DayReport"
        Label1(0).Caption = "日報表列印"
    Case "WeekReport"
        Label1(0).Caption = "週報表列印"
    Case "MonthReport"
        Label1(0).Caption = "月報表列印"
    Case "YearReport"
        Label1(0).Caption = "年報表列印"
    Case "DayAccount"
        Label1(0).Caption = "日總帳列印"
    Case "WeekAccount"
        Label1(0).Caption = "週總帳列印"
    Case "MonthAccount"
        Label1(0).Caption = "月總帳列印"
    Case "YearAccount"
        Label1(0).Caption = "年總帳列印"
    Case "FourKDayReport"
        Label1(0).Caption = "4K日報表列印"
    Case "FourKWeekReport"
        Label1(0).Caption = "4K週報表列印"
    Case "FourKMonthReport"
        Label1(0).Caption = "4K月報表列印"
    Case "FourKYearReport"
        Label1(0).Caption = "4K年報表列印"
    Case "FourKDayAccount"
        Label1(0).Caption = "4K日總帳列印"
    Case "FourKWeekAccount"
        Label1(0).Caption = "4K週總帳列印"
    Case "FourKMonthAccount"
        Label1(0).Caption = "4K月總帳列印"
    Case "FourKYearAccount"
        Label1(0).Caption = "4K年總帳列印"
    End Select
    
    
    dtpCurrentDate.Value = Format(DateTime.Now, "yyyy/MM/dd")
    txtCurrentDate.Text = Format(DateTime.Now, "yyyy/MM/dd")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmProve.Show
    Unload Me
End Sub
