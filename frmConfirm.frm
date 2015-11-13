VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmConfirm 
   BorderStyle     =   1  '單線固定
   Caption         =   "Report Confirm"
   ClientHeight    =   3990
   ClientLeft      =   6315
   ClientTop       =   7785
   ClientWidth     =   4680
   Icon            =   "frmConfirm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   4680
   Begin VB.ComboBox cmbPName 
      Height          =   300
      Left            =   1800
      TabIndex        =   11
      Top             =   1800
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ComboBox cmbCName 
      Height          =   300
      Left            =   1800
      TabIndex        =   10
      Top             =   840
      Visible         =   0   'False
      Width           =   1935
   End
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
      Top             =   1320
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
      Top             =   3360
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
      Left            =   0
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   3360
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker dtpCurrentDate 
      Height          =   360
      Left            =   1800
      TabIndex        =   5
      Top             =   1320
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
      Format          =   35454979
      CurrentDate     =   37058
   End
   Begin VB.Label lblEntry 
      Alignment       =   1  '靠右對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "產品名稱"
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
      Index           =   2
      Left            =   720
      TabIndex        =   9
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblEntry 
      Alignment       =   1  '靠右對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "客戶名稱"
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
      Index           =   0
      Left            =   720
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
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
      Top             =   2280
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
      Top             =   1320
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
      Height          =   3135
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
Option Explicit

Private Sub cmdCancel_Click()
    Call Form_Unload(0)
End Sub

Sub DayReport(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String
    Dim CurrentCount As Integer, CurrentPrice As Double, WinningCount As Integer, WinningPrice As Double
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset, rec2 As New adoDB.Recordset
    Dim BuyCurrentCount As Integer, SellCurrentCount As Integer, BuyCurrentPrice As Double, SellCurrentPrice As Double
    Dim BuyWinningCount As Integer, SellWinningCount As Integer, BuyWinningPrice As Double, SellWinningPrice As Double
    Dim OrderDate As String
    Dim Proportion As Integer, BonusTarget As String, BonusMoney As Double, BonusProduct As String
    Dim CID As String, CName As String
    
    
    BuyCurrentCount = 0
    BuyWinningCount = 0
    BuyCurrentPrice = 0
    BuyWinningPrice = 0
    SellCurrentCount = 0
    SellWinningCount = 0
    SellCurrentPrice = 0
    SellWinningPrice = 0
    
    
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
        Body = "<tr><td>產品</td>"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "交易金額</td>"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        Body = Body & "</tr>"
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
                selectFields = "CurrentCount,WinningCount,CurrentDate"
                SQL = "select " & selectFields & " from [order] where PID='" & PIDArray(i) & "' and CID='" & CID & "' and CurrentDate='" & txtCurrentDate.Text & "' order by PID;"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
                
                'enum custom every product order count
                CurrentCount = 0
                CurrentPrice = 0
                WinningCount = 0
                WinningPrice = 0
                Do Until rec1.EOF
                    'search price, and addition to variable
                    OrderDate = rec1.Fields.Item("CurrentDate")
                    selectFields = "CurrentPrice,WinningPrice,Upset"
                    SQL = "select * from price where PID='" & PIDArray(i) & "' and CID='" & CID & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
                    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
                
                    CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                    WinningCount = WinningCount + Val(rec1.Fields.Item("WinningCount"))
                    If Not price_rec.EOF Then
                        CurrentPrice = CurrentPrice + (CurrentCount * Val(price_rec.Fields.Item("CurrentPrice")))
                        WinningPrice = WinningPrice + (WinningCount * Val(price_rec.Fields.Item("WinningPrice")))
                    End If
                    
                    rec1.MoveNext
                Loop
                Body = Body & "<td>" & CurrentCount & "</td>"
                Body = Body & "<td>" & CurrentPrice & "</td>"
                
                'a new variable to account buy and sell
                If CurrentCount >= 0 Then
                    BuyCurrentCount = BuyCurrentCount + CurrentCount
                    BuyCurrentPrice = BuyCurrentPrice + CurrentPrice
                    BuyWinningCount = BuyWinningCount + WinningCount
                    BuyWinningPrice = BuyWinningPrice + WinningPrice
                Else
                    SellCurrentCount = SellCurrentCount + CurrentCount
                    SellCurrentPrice = SellCurrentPrice + CurrentPrice
                    SellWinningCount = SellWinningCount + WinningCount
                    SellWinningPrice = SellWinningPrice + WinningPrice
                End If
            Next
            Body = Body & "</tr>"
            Print #1, Body
            
            custom_rec.MoveNext
        Loop
        
        
        Body = "<tr></tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計買</td>"
        Body = Body & "<td>" & BuyCurrentCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & BuyCurrentPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計買中</td>"
        Body = Body & "<td>" & BuyWinningCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & BuyWinningPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計賣</td>"
        Body = Body & "<td>" & SellCurrentCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & SellCurrentPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計賣中</td>"
        Body = Body & "<td>" & SellWinningCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & SellWinningPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        
        
        Print #1, "</table>"
    Close #1
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Sub WeekReport(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String
    Dim CurrentCount As Integer, CurrentPrice As Double, WinningCount As Integer, WinningPrice As Double
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset, rec2 As New adoDB.Recordset
    Dim BuyCurrentCount As Integer, SellCurrentCount As Integer, BuyCurrentPrice As Double, SellCurrentPrice As Double
    Dim BuyWinningCount As Integer, SellWinningCount As Integer, BuyWinningPrice As Double, SellWinningPrice As Double
    Dim OrderDate As String
    Dim Proportion As Integer, BonusTarget As String, BonusMoney As Double, BonusProduct As String
    Dim CID As String, CName As String
    
    
    BuyCurrentCount = 0
    BuyWinningCount = 0
    BuyCurrentPrice = 0
    BuyWinningPrice = 0
    SellCurrentCount = 0
    SellWinningCount = 0
    SellCurrentPrice = 0
    SellWinningPrice = 0
    
    
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
        Body = "<tr><td>產品</td>"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "交易金額</td>"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        Body = Body & "<td>成數</td>"
        Body = Body & "<td>前帳</td>"
        Body = Body & "</tr>"
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
                selectFields = "CurrentCount,WinningCount,CurrentDate"
                SQL = "select " & selectFields & " from [order] where [order].PID='" & PIDArray(i) & "' and CID='" & CID & "' and (CurrentDate>='" & Format(DateTime.DateAdd("d", -7, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd") & "' and CurrentDate<='" & txtCurrentDate.Text & "') order by [order].PID;"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
                
                'enum custom every product order count
                CurrentCount = 0
                CurrentPrice = 0
                WinningCount = 0
                WinningPrice = 0
                Do Until rec1.EOF
                    'search price, and addition to variable
                    OrderDate = rec1.Fields.Item("CurrentDate")
                    selectFields = "CurrentPrice,WinningPrice,Upset"
                    SQL = "select * from price where PID='" & PIDArray(i) & "' and CID='" & CID & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
                    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
                
                    CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                    WinningCount = WinningCount + Val(rec1.Fields.Item("WinningCount"))
                    If Not price_rec.EOF Then
                        CurrentPrice = CurrentPrice + (CurrentCount * Val(price_rec.Fields.Item("CurrentPrice")))
                        WinningPrice = WinningPrice + (WinningCount * Val(price_rec.Fields.Item("WinningPrice")))
                    End If
                    
                    rec1.MoveNext
                Loop
                Body = Body & "<td>" & CurrentCount & "</td>"
                Body = Body & "<td>" & CurrentPrice & "</td>"
                
                'a new variable to account buy and sell
                If CurrentCount >= 0 Then
                    BuyCurrentCount = BuyCurrentCount + CurrentCount
                    BuyCurrentPrice = BuyCurrentPrice + CurrentPrice
                    BuyWinningCount = BuyWinningCount + WinningCount
                    BuyWinningPrice = BuyWinningPrice + WinningPrice
                Else
                    SellCurrentCount = SellCurrentCount + CurrentCount
                    SellCurrentPrice = SellCurrentPrice + CurrentPrice
                    SellWinningCount = SellWinningCount + WinningCount
                    SellWinningPrice = SellWinningPrice + WinningPrice
                End If
            Next
            
            
            'add proportion
            SQL = "select * from custom where BonusTarget='" & CID & " " & CName & "';"
            Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
            Do Until rec1.EOF
                BonusTarget = rec1.Fields.Item("CID")
                Proportion = rec1.Fields.Item("Proportion")
                
                'search the bonus fromer order count.
                BonusMoney = 0
                SQL = "select * from [order] where CID='" & BonusTarget & "' and (CurrentDate>='" & Format(DateTime.DateAdd("d", -7, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd") & "' and CurrentDate<='" & txtCurrentDate.Text & "');"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec2)
                Do Until rec2.EOF
                    'search price, and addition to variable
                    OrderDate = rec2.Fields.Item("CurrentDate")
                    BonusProduct = rec2.Fields.Item("PID")
                    selectFields = "CurrentPrice,WinningPrice,Upset"
                    SQL = "select * from price where PID='" & BonusProduct & "' and CID='" & BonusTarget & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
                    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
                    
                    If Not price_rec.EOF Then
                        BonusMoney = BonusMoney + (Val(rec2.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")))
                    End If
                    rec2.MoveNext
                Loop
                
                Body = Body & "<td>來自" & rec1.Fields.Item("CName") & "成數" & Proportion & "%共" & BonusMoney * (Proportion / 100) & "元。</td>"
                rec1.MoveNext
            Loop
            
            Body = Body & "</tr>"
            Print #1, Body
            
            custom_rec.MoveNext
        Loop
        
        
        Body = "<tr></tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計買</td>"
        Body = Body & "<td>" & BuyCurrentCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & BuyCurrentPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計買中</td>"
        Body = Body & "<td>" & BuyWinningCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & BuyWinningPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計賣</td>"
        Body = Body & "<td>" & SellCurrentCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & SellCurrentPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計賣中</td>"
        Body = Body & "<td>" & SellWinningCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & SellWinningPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        
        
        Print #1, "</table>"
    Close #1
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Sub MonthReport(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String
    Dim CurrentCount As Integer, CurrentPrice As Double, WinningCount As Integer, WinningPrice As Double
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset, rec2 As New adoDB.Recordset
    Dim BuyCurrentCount As Integer, SellCurrentCount As Integer, BuyCurrentPrice As Double, SellCurrentPrice As Double
    Dim BuyWinningCount As Integer, SellWinningCount As Integer, BuyWinningPrice As Double, SellWinningPrice As Double
    Dim OrderDate As String
    Dim Proportion As Integer, BonusTarget As String, BonusMoney As Double, BonusProduct As String
    Dim CID As String, CName As String
    
    
    BuyCurrentCount = 0
    BuyWinningCount = 0
    BuyCurrentPrice = 0
    BuyWinningPrice = 0
    SellCurrentCount = 0
    SellWinningCount = 0
    SellCurrentPrice = 0
    SellWinningPrice = 0
    
    
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
        Body = "<tr><td>產品</td>"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "交易金額</td>"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        Body = Body & "<td>成數</td>"
        Body = Body & "</tr>"
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
                selectFields = "CurrentCount,WinningCount,CurrentDate"
                SQL = "select " & selectFields & " from [order] where [order].PID='" & PIDArray(i) & "' and CID='" & CID & "' and (CurrentDate>='" & Format(txtCurrentDate.Text, "yyyy/MM/") & "01' and CurrentDate<='" & Format(txtCurrentDate.Text, "yyyy/MM/") & Date_Is_28_29_30_31(txtCurrentDate.Text) & "') order by [order].PID;"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
                
                'enum custom every product order count
                CurrentCount = 0
                CurrentPrice = 0
                WinningCount = 0
                WinningPrice = 0
                Do Until rec1.EOF
                    'search price, and addition to variable
                    OrderDate = rec1.Fields.Item("CurrentDate")
                    selectFields = "CurrentPrice,WinningPrice,Upset"
                    SQL = "select * from price where PID='" & PIDArray(i) & "' and CID='" & CID & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
                    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
                
                    CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                    WinningCount = WinningCount + Val(rec1.Fields.Item("WinningCount"))
                    If Not price_rec.EOF Then
                        CurrentPrice = CurrentPrice + (CurrentCount * Val(price_rec.Fields.Item("CurrentPrice")))
                        WinningPrice = WinningPrice + (WinningCount * Val(price_rec.Fields.Item("WinningPrice")))
                    End If
                    
                    rec1.MoveNext
                Loop
                Body = Body & "<td>" & CurrentCount & "</td>"
                Body = Body & "<td>" & CurrentPrice & "</td>"
                
                'a new variable to account buy and sell
                If CurrentCount >= 0 Then
                    BuyCurrentCount = BuyCurrentCount + CurrentCount
                    BuyCurrentPrice = BuyCurrentPrice + CurrentPrice
                    BuyWinningCount = BuyWinningCount + WinningCount
                    BuyWinningPrice = BuyWinningPrice + WinningPrice
                Else
                    SellCurrentCount = SellCurrentCount + CurrentCount
                    SellCurrentPrice = SellCurrentPrice + CurrentPrice
                    SellWinningCount = SellWinningCount + WinningCount
                    SellWinningPrice = SellWinningPrice + WinningPrice
                End If
            Next
            
            'add proportion
            SQL = "select * from custom where BonusTarget='" & CID & " " & CName & "';"
            Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
            Do Until rec1.EOF
                BonusTarget = rec1.Fields.Item("CID")
                Proportion = rec1.Fields.Item("Proportion")
                
                'search the bonus fromer order count.
                BonusMoney = 0
                SQL = "select * from [order] where CID='" & BonusTarget & "' and (CurrentDate>='" & Format(txtCurrentDate.Text, "yyyy/MM/") & "01' and CurrentDate<='" & Format(txtCurrentDate.Text, "yyyy/MM/") & Date_Is_28_29_30_31(txtCurrentDate.Text) & "');"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec2)
                Do Until rec2.EOF
                    'search price, and addition to variable
                    OrderDate = rec2.Fields.Item("CurrentDate")
                    BonusProduct = rec2.Fields.Item("PID")
                    selectFields = "CurrentPrice,WinningPrice,Upset"
                    SQL = "select * from price where PID='" & BonusProduct & "' and CID='" & BonusTarget & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
                    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
                    
                    If Not price_rec.EOF Then
                        BonusMoney = BonusMoney + (Val(rec2.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")))
                    End If
                    rec2.MoveNext
                Loop
                
                Body = Body & "<td>來自" & rec1.Fields.Item("CName") & "成數" & Proportion & "%共" & BonusMoney * (Proportion / 100) & "元。</td>"
                rec1.MoveNext
            Loop
            
            Body = Body & "</tr>"
            Print #1, Body
            
            custom_rec.MoveNext
        Loop
        
        
        Body = "<tr></tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計買</td>"
        Body = Body & "<td>" & BuyCurrentCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & BuyCurrentPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計買中</td>"
        Body = Body & "<td>" & BuyWinningCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & BuyWinningPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計賣</td>"
        Body = Body & "<td>" & SellCurrentCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & SellCurrentPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計賣中</td>"
        Body = Body & "<td>" & SellWinningCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & SellWinningPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        
        
        Print #1, "</table>"
    Close #1
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Sub YearReport(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String
    Dim CurrentCount As Integer, CurrentPrice As Double, WinningCount As Integer, WinningPrice As Double
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset, rec2 As New adoDB.Recordset
    Dim BuyCurrentCount As Integer, SellCurrentCount As Integer, BuyCurrentPrice As Double, SellCurrentPrice As Double
    Dim BuyWinningCount As Integer, SellWinningCount As Integer, BuyWinningPrice As Double, SellWinningPrice As Double
    Dim OrderDate As String
    Dim Proportion As Integer, BonusTarget As String, BonusMoney As Double, BonusProduct As String
    Dim CID As String, CName As String
    
    
    BuyCurrentCount = 0
    BuyWinningCount = 0
    BuyCurrentPrice = 0
    BuyWinningPrice = 0
    SellCurrentCount = 0
    SellWinningCount = 0
    SellCurrentPrice = 0
    SellWinningPrice = 0
    
    
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
        Body = "<tr><td>產品</td>"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "交易金額</td>"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        Body = Body & "<td>成數</td>"
        Body = Body & "</tr>"
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
                selectFields = "CurrentCount,WinningCount,CurrentDate"
                SQL = "select " & selectFields & " from [order] where [order].PID='" & PIDArray(i) & "' and CID='" & CID & "' and (CurrentDate>='" & Format(txtCurrentDate.Text, "yyyy/") & "01/01' and CurrentDate<='" & Format(txtCurrentDate.Text, "yyyy/") & "12/31') order by [order].PID;"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
                
                'search price
                selectFields = "CurrentPrice,WinningPrice,Upset"
                SQL = "select * from price where PID='" & PIDArray(i) & "' and CID='" & CID & "' and CurrentDate<='" & txtCurrentDate.Text & "' order by CurrentDate desc;"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
                
                'enum custom every product order count
                CurrentCount = 0
                CurrentPrice = 0
                WinningCount = 0
                WinningPrice = 0
                Do Until rec1.EOF
                    'search price, and addition to variable
                    OrderDate = rec1.Fields.Item("CurrentDate")
                    selectFields = "CurrentPrice,WinningPrice,Upset"
                    SQL = "select * from price where PID='" & PIDArray(i) & "' and CID='" & CID & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
                    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
                
                    CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                    WinningCount = WinningCount + Val(rec1.Fields.Item("WinningCount"))
                    If Not price_rec.EOF Then
                        CurrentPrice = CurrentPrice + (CurrentCount * Val(price_rec.Fields.Item("CurrentPrice")))
                        WinningPrice = WinningPrice + (WinningCount * Val(price_rec.Fields.Item("WinningPrice")))
                    End If
                    
                    rec1.MoveNext
                Loop
                Body = Body & "<td>" & CurrentCount & "</td>"
                Body = Body & "<td>" & CurrentPrice & "</td>"
                
                'a new variable to account buy and sell
                If CurrentCount >= 0 Then
                    BuyCurrentCount = BuyCurrentCount + CurrentCount
                    BuyCurrentPrice = BuyCurrentPrice + CurrentPrice
                    BuyWinningCount = BuyWinningCount + WinningCount
                    BuyWinningPrice = BuyWinningPrice + WinningPrice
                Else
                    SellCurrentCount = SellCurrentCount + CurrentCount
                    SellCurrentPrice = SellCurrentPrice + CurrentPrice
                    SellWinningCount = SellWinningCount + WinningCount
                    SellWinningPrice = SellWinningPrice + WinningPrice
                End If
            Next
            
            'add proportion
            SQL = "select * from custom where BonusTarget='" & CID & " " & CName & "';"
            Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
            Do Until rec1.EOF
                BonusTarget = rec1.Fields.Item("CID")
                Proportion = rec1.Fields.Item("Proportion")
                
                'search the bonus fromer order count.
                BonusMoney = 0
                SQL = "select * from [order] where CID='" & BonusTarget & "' and (CurrentDate>='" & Format(txtCurrentDate.Text, "yyyy/") & "01/01' and CurrentDate<='" & Format(txtCurrentDate.Text, "yyyy/") & "12/31');"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec2)
                Do Until rec2.EOF
                    'search price, and addition to variable
                    OrderDate = rec2.Fields.Item("CurrentDate")
                    BonusProduct = rec2.Fields.Item("PID")
                    selectFields = "CurrentPrice,WinningPrice,Upset"
                    SQL = "select * from price where PID='" & BonusProduct & "' and CID='" & BonusTarget & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
                    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
                    
                    If Not price_rec.EOF Then
                        BonusMoney = BonusMoney + (Val(rec2.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")))
                    End If
                    rec2.MoveNext
                Loop
                
                Body = Body & "<td>來自" & rec1.Fields.Item("CName") & "成數" & Proportion & "%共" & BonusMoney * (Proportion / 100) & "元。</td>"
                rec1.MoveNext
            Loop
            
            Body = Body & "</tr>"
            Print #1, Body
            
            custom_rec.MoveNext
        Loop
        
        
        Body = "<tr></tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計買</td>"
        Body = Body & "<td>" & BuyCurrentCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & BuyCurrentPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計買中</td>"
        Body = Body & "<td>" & BuyWinningCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & BuyWinningPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計賣</td>"
        Body = Body & "<td>" & SellCurrentCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & SellCurrentPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計賣中</td>"
        Body = Body & "<td>" & SellWinningCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & SellWinningPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        
        
        Print #1, "</table>"
    Close #1
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Sub DayAccount(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String
    Dim CurrentCount As Integer, CurrentPrice As Double, WinningCount As Integer, WinningPrice As Double
    Dim CurrentCountAll(1024) As Long, WinningCountAll(1024) As Long
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset, rec2 As New adoDB.Recordset
    Dim BuyCurrentCount As Integer, SellCurrentCount As Integer, BuyCurrentPrice As Double, SellCurrentPrice As Double
    Dim BuyWinningCount As Integer, SellWinningCount As Integer, BuyWinningPrice As Double, SellWinningPrice As Double
    Dim OrderDate As String
    Dim Proportion As Integer, BonusTarget As String, BonusMoney As Double, BonusProduct As String
    Dim CID As String, CName As String
    
    
    BuyCurrentCount = 0
    BuyWinningCount = 0
    BuyCurrentPrice = 0
    BuyWinningPrice = 0
    SellCurrentCount = 0
    SellWinningCount = 0
    SellCurrentPrice = 0
    SellWinningPrice = 0
    
    
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
        Body = "<tr><td>產品</td>"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "交易金額</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "中</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "中獎金額</td>"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        Body = Body & "</tr>"
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
                selectFields = "CurrentCount,WinningCount,CurrentDate"
                SQL = "select " & selectFields & " from [order] where [order].PID='" & PIDArray(i) & "' and CID='" & CID & "' and CurrentDate='" & txtCurrentDate.Text & "' order by [order].PID;"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
                
                'enum custom every product order count
                CurrentCount = 0
                CurrentPrice = 0
                WinningCount = 0
                WinningPrice = 0
                Do Until rec1.EOF
                    'search price, and addition to variable
                    OrderDate = rec1.Fields.Item("CurrentDate")
                    selectFields = "CurrentPrice,WinningPrice,Upset"
                    SQL = "select * from price where PID='" & PIDArray(i) & "' and CID='" & CID & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
                    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
                
                    CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                    WinningCount = WinningCount + Val(rec1.Fields.Item("WinningCount"))
                    If Not price_rec.EOF Then
                        CurrentPrice = CurrentPrice + (CurrentCount * Val(price_rec.Fields.Item("CurrentPrice")))
                        WinningPrice = WinningPrice + (WinningCount * Val(price_rec.Fields.Item("WinningPrice")))
                    End If
                    
                    rec1.MoveNext
                Loop
                Body = Body & "<td>" & CurrentCount & "</td>"
                Body = Body & "<td>" & CurrentPrice & "</td>"
                Body = Body & "<td>" & WinningCount & "</td>"
                Body = Body & "<td>" & WinningPrice & "</td>"
                
                'a new variable to account buy and sell
                If CurrentCount >= 0 Then
                    BuyCurrentCount = BuyCurrentCount + CurrentCount
                    BuyCurrentPrice = BuyCurrentPrice + CurrentPrice
                    BuyWinningCount = BuyWinningCount + WinningCount
                    BuyWinningPrice = BuyWinningPrice + WinningPrice
                Else
                    SellCurrentCount = SellCurrentCount + CurrentCount
                    SellCurrentPrice = SellCurrentPrice + CurrentPrice
                    SellWinningCount = SellWinningCount + WinningCount
                    SellWinningPrice = SellWinningPrice + WinningPrice
                End If
                
                'a new variable to account every product
                CurrentCountAll(i) = CurrentCountAll(i) + CurrentCount
                WinningCountAll(i) = WinningCountAll(i) + WinningCount
            Next
            Body = Body & "</tr>"
            Print #1, Body
            
            custom_rec.MoveNext
        Loop
        
        
        Body = "<tr><td>數量總計</td>"
        For i = 0 To Count - 1
            Body = Body & "<td>" & CurrentCountAll(i) & "</td><td></td>"
            Body = Body & "<td>" & WinningCountAll(i) & "</td><td></td>"
        Next
        Body = Body & "</tr>"
        Print #1, Body
        
        
        Body = "<tr></tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計買</td>"
        Body = Body & "<td>" & BuyCurrentCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & BuyCurrentPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計買中</td>"
        Body = Body & "<td>" & BuyWinningCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & BuyWinningPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計賣</td>"
        Body = Body & "<td>" & SellCurrentCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & SellCurrentPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計賣中</td>"
        Body = Body & "<td>" & SellWinningCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & SellWinningPrice & "</td><td></td>"
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
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String
    Dim CurrentCount As Integer, CurrentPrice As Double, WinningCount As Integer, WinningPrice As Double
    Dim CurrentCountAll(1024) As Long, WinningCountAll(1024) As Long
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset, rec2 As New adoDB.Recordset
    Dim BuyCurrentCount As Integer, SellCurrentCount As Integer, BuyCurrentPrice As Double, SellCurrentPrice As Double
    Dim BuyWinningCount As Integer, SellWinningCount As Integer, BuyWinningPrice As Double, SellWinningPrice As Double
    Dim OrderDate As String
    Dim Proportion As Integer, BonusTarget As String, BonusMoney As Double, BonusProduct As String
    Dim CID As String, CName As String
    
    
    BuyCurrentCount = 0
    BuyWinningCount = 0
    BuyCurrentPrice = 0
    BuyWinningPrice = 0
    SellCurrentCount = 0
    SellWinningCount = 0
    SellCurrentPrice = 0
    SellWinningPrice = 0
    
    
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
        Body = "<tr><td>產品</td>"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "交易金額</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "中</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "中獎金額</td>"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        Body = Body & "<td>成數</td>"
        Body = Body & "<td>前帳</td>"
        Body = Body & "</tr>"
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
                selectFields = "CurrentCount,WinningCount,CurrentDate"
                SQL = "select " & selectFields & " from [order] where [order].PID='" & PIDArray(i) & "' and CID='" & CID & "' and (CurrentDate>='" & Format(DateTime.DateAdd("d", -7, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd") & "' and CurrentDate<='" & txtCurrentDate.Text & "') order by [order].PID;"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
                
                'enum custom every product order count
                CurrentCount = 0
                CurrentPrice = 0
                WinningCount = 0
                WinningPrice = 0
                Do Until rec1.EOF
                    'search price, and addition to variable
                    OrderDate = rec1.Fields.Item("CurrentDate")
                    selectFields = "CurrentPrice,WinningPrice,Upset"
                    SQL = "select * from price where PID='" & PIDArray(i) & "' and CID='" & CID & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
                    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
                
                    CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                    WinningCount = WinningCount + Val(rec1.Fields.Item("WinningCount"))
                    If Not price_rec.EOF Then
                        CurrentPrice = CurrentPrice + (CurrentCount * Val(price_rec.Fields.Item("CurrentPrice")))
                        WinningPrice = WinningPrice + (WinningCount * Val(price_rec.Fields.Item("WinningPrice")))
                    End If
                    
                    rec1.MoveNext
                Loop
                Body = Body & "<td>" & CurrentCount & "</td>"
                Body = Body & "<td>" & CurrentPrice & "</td>"
                Body = Body & "<td>" & WinningCount & "</td>"
                Body = Body & "<td>" & WinningPrice & "</td>"
                
                'a new variable to account buy and sell
                If CurrentCount >= 0 Then
                    BuyCurrentCount = BuyCurrentCount + CurrentCount
                    BuyCurrentPrice = BuyCurrentPrice + CurrentPrice
                    BuyWinningCount = BuyWinningCount + WinningCount
                    BuyWinningPrice = BuyWinningPrice + WinningPrice
                Else
                    SellCurrentCount = SellCurrentCount + CurrentCount
                    SellCurrentPrice = SellCurrentPrice + CurrentPrice
                    SellWinningCount = SellWinningCount + WinningCount
                    SellWinningPrice = SellWinningPrice + WinningPrice
                End If
                
                'a new variable to account every product
                CurrentCountAll(i) = CurrentCountAll(i) + CurrentCount
                WinningCountAll(i) = WinningCountAll(i) + WinningCount
            Next
            
            'add proportion
            SQL = "select * from custom where BonusTarget='" & CID & " " & CName & "';"
            Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
            Do Until rec1.EOF
                BonusTarget = rec1.Fields.Item("CID")
                Proportion = rec1.Fields.Item("Proportion")
                
                'search the bonus fromer order count.
                BonusMoney = 0
                SQL = "select * from [order] where CID='" & BonusTarget & "' and (CurrentDate>='" & Format(DateTime.DateAdd("d", -7, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd") & "' and CurrentDate<='" & txtCurrentDate.Text & "');"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec2)
                Do Until rec2.EOF
                    'search price, and addition to variable
                    OrderDate = rec2.Fields.Item("CurrentDate")
                    BonusProduct = rec2.Fields.Item("PID")
                    selectFields = "CurrentPrice,WinningPrice,Upset"
                    SQL = "select * from price where PID='" & BonusProduct & "' and CID='" & BonusTarget & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
                    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
                    
                    If Not price_rec.EOF Then
                        BonusMoney = BonusMoney + (Val(rec2.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")))
                    End If
                    rec2.MoveNext
                Loop
                
                Body = Body & "<td>來自" & rec1.Fields.Item("CName") & "成數" & Proportion & "%共" & BonusMoney * (Proportion / 100) & "元。</td>"
                rec1.MoveNext
            Loop
            
            Body = Body & "</tr>"
            Print #1, Body
            
            custom_rec.MoveNext
        Loop
        
        
        Body = "<tr><td>數量總計</td>"
        For i = 0 To Count - 1
            Body = Body & "<td>" & CurrentCountAll(i) & "</td><td></td>"
            Body = Body & "<td>" & WinningCountAll(i) & "</td><td></td>"
        Next
        Body = Body & "</tr>"
        Print #1, Body
        
        
        Body = "<tr></tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計買</td>"
        Body = Body & "<td>" & BuyCurrentCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & BuyCurrentPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計買中</td>"
        Body = Body & "<td>" & BuyWinningCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & BuyWinningPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計賣</td>"
        Body = Body & "<td>" & SellCurrentCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & SellCurrentPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計賣中</td>"
        Body = Body & "<td>" & SellWinningCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & SellWinningPrice & "</td><td></td>"
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
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String
    Dim CurrentCount As Integer, CurrentPrice As Double, WinningCount As Integer, WinningPrice As Double
    Dim CurrentCountAll(1024) As Long, WinningCountAll(1024) As Long
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset, rec2 As New adoDB.Recordset
    Dim BuyCurrentCount As Integer, SellCurrentCount As Integer, BuyCurrentPrice As Double, SellCurrentPrice As Double
    Dim BuyWinningCount As Integer, SellWinningCount As Integer, BuyWinningPrice As Double, SellWinningPrice As Double
    Dim OrderDate As String
    Dim Proportion As Integer, BonusTarget As String, BonusMoney As Double, BonusProduct As String
    Dim CID As String, CName As String
    
    
    BuyCurrentCount = 0
    BuyWinningCount = 0
    BuyCurrentPrice = 0
    BuyWinningPrice = 0
    SellCurrentCount = 0
    SellWinningCount = 0
    SellCurrentPrice = 0
    SellWinningPrice = 0
    
    
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
        Body = "<tr><td>產品</td>"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "交易金額</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "中</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "中獎金額</td>"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        Body = Body & "<td>成數</td>"
        Body = Body & "</tr>"
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
                selectFields = "CurrentCount,WinningCount,CurrentDate"
                SQL = "select " & selectFields & " from [order] where [order].PID='" & PIDArray(i) & "' and CID='" & CID & "' and (CurrentDate>='" & Format(txtCurrentDate.Text, "yyyy/MM/") & "01' and CurrentDate<='" & Format(txtCurrentDate.Text, "yyyy/MM/") & Date_Is_28_29_30_31(txtCurrentDate.Text) & "') order by [order].PID;"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)

                'enum custom every product order count
                CurrentCount = 0
                CurrentPrice = 0
                WinningCount = 0
                WinningPrice = 0
                Do Until rec1.EOF
                    'search price, and addition to variable
                    OrderDate = rec1.Fields.Item("CurrentDate")
                    selectFields = "CurrentPrice,WinningPrice,Upset"
                    SQL = "select * from price where PID='" & PIDArray(i) & "' and CID='" & CID & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
                    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
                
                    CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                    WinningCount = WinningCount + Val(rec1.Fields.Item("WinningCount"))
                    If Not price_rec.EOF Then
                        CurrentPrice = CurrentPrice + (CurrentCount * Val(price_rec.Fields.Item("CurrentPrice")))
                        WinningPrice = WinningPrice + (WinningCount * Val(price_rec.Fields.Item("WinningPrice")))
                    End If
                    
                    rec1.MoveNext
                Loop
                Body = Body & "<td>" & CurrentCount & "</td>"
                Body = Body & "<td>" & CurrentPrice & "</td>"
                Body = Body & "<td>" & WinningCount & "</td>"
                Body = Body & "<td>" & WinningPrice & "</td>"
                
                'a new variable to account buy and sell
                If CurrentCount >= 0 Then
                    BuyCurrentCount = BuyCurrentCount + CurrentCount
                    BuyCurrentPrice = BuyCurrentPrice + CurrentPrice
                    BuyWinningCount = BuyWinningCount + WinningCount
                    BuyWinningPrice = BuyWinningPrice + WinningPrice
                Else
                    SellCurrentCount = SellCurrentCount + CurrentCount
                    SellCurrentPrice = SellCurrentPrice + CurrentPrice
                    SellWinningCount = SellWinningCount + WinningCount
                    SellWinningPrice = SellWinningPrice + WinningPrice
                End If
                
                'a new variable to account every product
                CurrentCountAll(i) = CurrentCountAll(i) + CurrentCount
                WinningCountAll(i) = WinningCountAll(i) + WinningCount
            Next
            
            'add proportion
            SQL = "select * from custom where BonusTarget='" & CID & " " & CName & "';"
            Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
            Do Until rec1.EOF
                BonusTarget = rec1.Fields.Item("CID")
                Proportion = rec1.Fields.Item("Proportion")
                
                'search the bonus fromer order count.
                BonusMoney = 0
                SQL = "select * from [order] where CID='" & BonusTarget & "' and (CurrentDate>='" & Format(txtCurrentDate.Text, "yyyy/MM/") & "01' and CurrentDate<='" & Format(txtCurrentDate.Text, "yyyy/MM/") & Date_Is_28_29_30_31(txtCurrentDate.Text) & "');"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec2)
                Do Until rec2.EOF
                    'search price, and addition to variable
                    OrderDate = rec2.Fields.Item("CurrentDate")
                    BonusProduct = rec2.Fields.Item("PID")
                    selectFields = "CurrentPrice,WinningPrice,Upset"
                    SQL = "select * from price where PID='" & BonusProduct & "' and CID='" & BonusTarget & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
                    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
                    
                    If Not price_rec.EOF Then
                        BonusMoney = BonusMoney + (Val(rec2.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")))
                    End If
                    rec2.MoveNext
                Loop
                
                Body = Body & "<td>來自" & rec1.Fields.Item("CName") & "成數" & Proportion & "%共" & BonusMoney * (Proportion / 100) & "元。</td>"
                rec1.MoveNext
            Loop
            
            Body = Body & "</tr>"
            Print #1, Body
            
            custom_rec.MoveNext
        Loop
        
        
        Body = "<tr><td>數量總計</td>"
        For i = 0 To Count - 1
            Body = Body & "<td>" & CurrentCountAll(i) & "</td><td></td>"
            Body = Body & "<td>" & WinningCountAll(i) & "</td><td></td>"
        Next
        Body = Body & "</tr>"
        Print #1, Body
        
        
        Body = "<tr></tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計買</td>"
        Body = Body & "<td>" & BuyCurrentCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & BuyCurrentPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計買中</td>"
        Body = Body & "<td>" & BuyWinningCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & BuyWinningPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計賣</td>"
        Body = Body & "<td>" & SellCurrentCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & SellCurrentPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計賣中</td>"
        Body = Body & "<td>" & SellWinningCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & SellWinningPrice & "</td><td></td>"
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
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String
    Dim CurrentCount As Integer, CurrentPrice As Double, WinningCount As Integer, WinningPrice As Double
    Dim CurrentCountAll(1024) As Long, WinningCountAll(1024) As Long
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset, rec2 As New adoDB.Recordset
    Dim BuyCurrentCount As Integer, SellCurrentCount As Integer, BuyCurrentPrice As Double, SellCurrentPrice As Double
    Dim BuyWinningCount As Integer, SellWinningCount As Integer, BuyWinningPrice As Double, SellWinningPrice As Double
    Dim OrderDate As String
    Dim Proportion As Integer, BonusTarget As String, BonusMoney As Double, BonusProduct As String
    Dim CID As String, CName As String
    
    
    BuyCurrentCount = 0
    BuyWinningCount = 0
    BuyCurrentPrice = 0
    BuyWinningPrice = 0
    SellCurrentCount = 0
    SellWinningCount = 0
    SellCurrentPrice = 0
    SellWinningPrice = 0
    
    
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
        Body = "<tr><td>產品</td>"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "交易金額</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "中</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "中獎金額</td>"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        Body = Body & "<td>成數</td>"
        Body = Body & "</tr>"
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
                selectFields = "CurrentCount,WinningCount,CurrentDate"
                SQL = "select " & selectFields & " from [order] where [order].PID='" & PIDArray(i) & "' and CID='" & CID & "' and (CurrentDate>='" & Format(txtCurrentDate.Text, "yyyy/") & "01/01' and CurrentDate<='" & Format(txtCurrentDate.Text, "yyyy/") & "12/31') order by [order].PID;"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
                
                'enum custom every product order count
                CurrentCount = 0
                CurrentPrice = 0
                WinningCount = 0
                WinningPrice = 0
                Do Until rec1.EOF
                    'search price, and addition to variable
                    OrderDate = rec1.Fields.Item("CurrentDate")
                    selectFields = "CurrentPrice,WinningPrice,Upset"
                    SQL = "select * from price where PID='" & PIDArray(i) & "' and CID='" & CID & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
                    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
                
                    CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                    WinningCount = WinningCount + Val(rec1.Fields.Item("WinningCount"))
                    If Not price_rec.EOF Then
                        CurrentPrice = CurrentPrice + (CurrentCount * Val(price_rec.Fields.Item("CurrentPrice")))
                        WinningPrice = WinningPrice + (WinningCount * Val(price_rec.Fields.Item("WinningPrice")))
                    End If
                    
                    rec1.MoveNext
                Loop
                Body = Body & "<td>" & CurrentCount & "</td>"
                Body = Body & "<td>" & CurrentPrice & "</td>"
                Body = Body & "<td>" & WinningCount & "</td>"
                Body = Body & "<td>" & WinningPrice & "</td>"
                
                'a new variable to account buy and sell
                If CurrentCount >= 0 Then
                    BuyCurrentCount = BuyCurrentCount + CurrentCount
                    BuyCurrentPrice = BuyCurrentPrice + CurrentPrice
                    BuyWinningCount = BuyWinningCount + WinningCount
                    BuyWinningPrice = BuyWinningPrice + WinningPrice
                Else
                    SellCurrentCount = SellCurrentCount + CurrentCount
                    SellCurrentPrice = SellCurrentPrice + CurrentPrice
                    SellWinningCount = SellWinningCount + WinningCount
                    SellWinningPrice = SellWinningPrice + WinningPrice
                End If
                
                'a new variable to account every product
                CurrentCountAll(i) = CurrentCountAll(i) + CurrentCount
                WinningCountAll(i) = WinningCountAll(i) + WinningCount
            Next
            
            'add proportion
            SQL = "select * from custom where BonusTarget='" & CID & " " & CName & "';"
            Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
            Do Until rec1.EOF
                BonusTarget = rec1.Fields.Item("CID")
                Proportion = rec1.Fields.Item("Proportion")
                
                'search the bonus fromer order count.
                BonusMoney = 0
                SQL = "select * from [order] where CID='" & BonusTarget & "' and (CurrentDate>='" & Format(txtCurrentDate.Text, "yyyy/") & "01/01' and CurrentDate<='" & Format(txtCurrentDate.Text, "yyyy/") & "12/31');"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec2)
                Do Until rec2.EOF
                    'search price, and addition to variable
                    OrderDate = rec2.Fields.Item("CurrentDate")
                    BonusProduct = rec2.Fields.Item("PID")
                    selectFields = "CurrentPrice,WinningPrice,Upset"
                    SQL = "select * from price where PID='" & BonusProduct & "' and CID='" & BonusTarget & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
                    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
                    
                    If Not price_rec.EOF Then
                        BonusMoney = BonusMoney + (Val(rec2.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")))
                    End If
                    rec2.MoveNext
                Loop
                
                Body = Body & "<td>來自" & rec1.Fields.Item("CName") & "成數" & Proportion & "%共" & BonusMoney * (Proportion / 100) & "元。</td>"
                rec1.MoveNext
            Loop
            
            Body = Body & "</tr>"
            Print #1, Body
            
            custom_rec.MoveNext
        Loop
        
        
        Body = "<tr><td>數量總計</td>"
        For i = 0 To Count - 1
            Body = Body & "<td>" & CurrentCountAll(i) & "</td><td></td>"
            Body = Body & "<td>" & WinningCountAll(i) & "</td><td></td>"
        Next
        Body = Body & "</tr>"
        Print #1, Body
        
        
        Body = "<tr></tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計買</td>"
        Body = Body & "<td>" & BuyCurrentCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & BuyCurrentPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計買中</td>"
        Body = Body & "<td>" & BuyWinningCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & BuyWinningPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計賣</td>"
        Body = Body & "<td>" & SellCurrentCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & SellCurrentPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計賣中</td>"
        Body = Body & "<td>" & SellWinningCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & SellWinningPrice & "</td><td></td>"
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
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String
    Dim CurrentCount As Integer, CurrentPrice As Double, WinningCount As Integer, WinningPrice As Double
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset, rec2 As New adoDB.Recordset
    Dim BuyCurrentCount As Integer, SellCurrentCount As Integer, BuyCurrentPrice As Double, SellCurrentPrice As Double
    Dim BuyWinningCount As Integer, SellWinningCount As Integer, BuyWinningPrice As Double, SellWinningPrice As Double
    Dim OrderDate As String
    Dim Proportion As Integer, BonusTarget As String, BonusMoney As Double, BonusProduct As String
    Dim CID As String, CName As String
    
    
    BuyCurrentCount = 0
    BuyWinningCount = 0
    BuyCurrentPrice = 0
    BuyWinningPrice = 0
    SellCurrentCount = 0
    SellWinningCount = 0
    SellCurrentPrice = 0
    SellWinningPrice = 0
    
    
    SQL = "select * from product where PName like '539_4K' order by PID;" '只要顯示539的4K 'SQL = "select * from product where PName like '%4K%' order by PID;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    
    
    Open TargetPath For Output As #1
        Print #1, "<table>"
    
        'show report datetime
        Body = "<tr><td>日期</td><td colspan=10>" & txtCurrentDate.Text & "</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>產品</td>"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "交易金額</td>"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        Body = Body & "</tr>"
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
                selectFields = "CurrentCount,WinningCount,CurrentDate"
                SQL = "select " & selectFields & " from [order] where [order].PID='" & PIDArray(i) & "' and CID='" & CID & "' and CurrentDate='" & txtCurrentDate.Text & "' order by [order].PID;"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)

                'enum custom every product order count
                CurrentCount = 0
                CurrentPrice = 0
                WinningCount = 0
                WinningPrice = 0
                Do Until rec1.EOF
                    'search price, and addition to variable
                    OrderDate = rec1.Fields.Item("CurrentDate")
                    selectFields = "CurrentPrice,WinningPrice,Upset"
                    SQL = "select * from price where PID='" & PIDArray(i) & "' and CID='" & CID & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
                    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
                
                    CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                    WinningCount = WinningCount + Val(rec1.Fields.Item("WinningCount"))
                    If Not price_rec.EOF Then
                        CurrentPrice = CurrentPrice + (CurrentCount * Val(price_rec.Fields.Item("CurrentPrice")))
                        WinningPrice = WinningPrice + (WinningCount * Val(price_rec.Fields.Item("WinningPrice")))
                    End If
                    
                    rec1.MoveNext
                Loop
                Body = Body & "<td>" & CurrentCount & "</td>"
                Body = Body & "<td>" & CurrentPrice & "</td>"
                
                'a new variable to account buy and sell
                If CurrentCount >= 0 Then
                    BuyCurrentCount = BuyCurrentCount + CurrentCount
                    BuyCurrentPrice = BuyCurrentPrice + CurrentPrice
                    BuyWinningCount = BuyWinningCount + WinningCount
                    BuyWinningPrice = BuyWinningPrice + WinningPrice
                Else
                    SellCurrentCount = SellCurrentCount + CurrentCount
                    SellCurrentPrice = SellCurrentPrice + CurrentPrice
                    SellWinningCount = SellWinningCount + WinningCount
                    SellWinningPrice = SellWinningPrice + WinningPrice
                End If
            Next
            Body = Body & "</tr>"
            Print #1, Body
            
            custom_rec.MoveNext
        Loop
        
        
        Body = "<tr></tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計買</td>"
        Body = Body & "<td>" & BuyCurrentCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & BuyCurrentPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計買中</td>"
        Body = Body & "<td>" & BuyWinningCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & BuyWinningPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計賣</td>"
        Body = Body & "<td>" & SellCurrentCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & SellCurrentPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計賣中</td>"
        Body = Body & "<td>" & SellWinningCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & SellWinningPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        
        
        Print #1, "</table>"
    Close #1
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Sub FourKWeekReport(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String
    Dim CurrentCount As Integer, CurrentPrice As Double, WinningCount As Integer, WinningPrice As Double
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset, rec2 As New adoDB.Recordset
    Dim BuyCurrentCount As Integer, SellCurrentCount As Integer, BuyCurrentPrice As Double, SellCurrentPrice As Double
    Dim BuyWinningCount As Integer, SellWinningCount As Integer, BuyWinningPrice As Double, SellWinningPrice As Double
    Dim OrderDate As String
    Dim Proportion As Integer, BonusTarget As String, BonusMoney As Double, BonusProduct As String
    Dim CID As String, CName As String
    
    
    BuyCurrentCount = 0
    BuyWinningCount = 0
    BuyCurrentPrice = 0
    BuyWinningPrice = 0
    SellCurrentCount = 0
    SellWinningCount = 0
    SellCurrentPrice = 0
    SellWinningPrice = 0
    
    
    SQL = "select * from product where PName like '539_4K' order by PID;" '只要顯示539的4K 'SQL = "select * from product where PName like '%4K%' order by PID;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    
    
    Open TargetPath For Output As #1
        Print #1, "<table>"
    
        'show report datetime
        Body = "<tr><td>日期</td><td colspan=10>" & DateTime.DateAdd("d", -7, txtCurrentDate.Text) & "至" & txtCurrentDate.Text & "</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>產品</td>"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "交易金額</td>"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        Body = Body & "<td>成數</td>"
        Body = Body & "<td>前帳</td>"
        Body = Body & "</tr>"
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
                selectFields = "CurrentCount,WinningCount,CurrentDate"
                SQL = "select " & selectFields & " from [order] where [order].PID='" & PIDArray(i) & "' and CID='" & CID & "' and (CurrentDate>='" & Format(DateTime.DateAdd("d", -7, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd") & "' and CurrentDate<='" & txtCurrentDate.Text & "') order by [order].PID;"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
                
                'enum custom every product order count
                CurrentCount = 0
                CurrentPrice = 0
                WinningCount = 0
                WinningPrice = 0
                Do Until rec1.EOF
                    'search price, and addition to variable
                    OrderDate = rec1.Fields.Item("CurrentDate")
                    selectFields = "CurrentPrice,WinningPrice,Upset"
                    SQL = "select * from price where PID='" & PIDArray(i) & "' and CID='" & CID & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
                    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
                
                    CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                    WinningCount = WinningCount + Val(rec1.Fields.Item("WinningCount"))
                    If Not price_rec.EOF Then
                        CurrentPrice = CurrentPrice + (CurrentCount * Val(price_rec.Fields.Item("CurrentPrice")))
                        WinningPrice = WinningPrice + (WinningCount * Val(price_rec.Fields.Item("WinningPrice")))
                    End If
                    
                    rec1.MoveNext
                Loop
                Body = Body & "<td>" & CurrentCount & "</td>"
                Body = Body & "<td>" & CurrentPrice & "</td>"
                
                'a new variable to account buy and sell
                If CurrentCount >= 0 Then
                    BuyCurrentCount = BuyCurrentCount + CurrentCount
                    BuyCurrentPrice = BuyCurrentPrice + CurrentPrice
                    BuyWinningCount = BuyWinningCount + WinningCount
                    BuyWinningPrice = BuyWinningPrice + WinningPrice
                Else
                    SellCurrentCount = SellCurrentCount + CurrentCount
                    SellCurrentPrice = SellCurrentPrice + CurrentPrice
                    SellWinningCount = SellWinningCount + WinningCount
                    SellWinningPrice = SellWinningPrice + WinningPrice
                End If
            Next
            
            'add proportion
            SQL = "select * from custom where BonusTarget='" & CID & " " & CName & "';"
            Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
            Do Until rec1.EOF
                BonusTarget = rec1.Fields.Item("CID")
                Proportion = rec1.Fields.Item("Proportion")
                
                'search the bonus fromer order count.
                BonusMoney = 0
                SQL = "select * from [order] where CID='" & BonusTarget & "' and (CurrentDate>='" & Format(DateTime.DateAdd("d", -7, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd") & "' and CurrentDate<='" & txtCurrentDate.Text & "');"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec2)
                Do Until rec2.EOF
                    'search price, and addition to variable
                    OrderDate = rec2.Fields.Item("CurrentDate")
                    BonusProduct = rec2.Fields.Item("PID")
                    selectFields = "CurrentPrice,WinningPrice,Upset"
                    SQL = "select * from price where PID='" & BonusProduct & "' and CID='" & BonusTarget & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
                    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
                    
                    If Not price_rec.EOF Then
                        BonusMoney = BonusMoney + (Val(rec2.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")))
                    End If
                    rec2.MoveNext
                Loop
                
                Body = Body & "<td>來自" & rec1.Fields.Item("CName") & "成數" & Proportion & "%共" & BonusMoney * (Proportion / 100) & "元。</td>"
                rec1.MoveNext
            Loop
            
            Body = Body & "</tr>"
            Print #1, Body
            
            custom_rec.MoveNext
        Loop

        
        Body = "<tr></tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計買</td>"
        Body = Body & "<td>" & BuyCurrentCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & BuyCurrentPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計買中</td>"
        Body = Body & "<td>" & BuyWinningCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & BuyWinningPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計賣</td>"
        Body = Body & "<td>" & SellCurrentCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & SellCurrentPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計賣中</td>"
        Body = Body & "<td>" & SellWinningCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & SellWinningPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        
        
        Print #1, "</table>"
    Close #1
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Sub FourKMonthReport(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String
    Dim CurrentCount As Integer, CurrentPrice As Double, WinningCount As Integer, WinningPrice As Double
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset, rec2 As New adoDB.Recordset
    Dim BuyCurrentCount As Integer, SellCurrentCount As Integer, BuyCurrentPrice As Double, SellCurrentPrice As Double
    Dim BuyWinningCount As Integer, SellWinningCount As Integer, BuyWinningPrice As Double, SellWinningPrice As Double
    Dim OrderDate As String
    Dim Proportion As Integer, BonusTarget As String, BonusMoney As Double, BonusProduct As String
    Dim CID As String, CName As String
    
    
    BuyCurrentCount = 0
    BuyWinningCount = 0
    BuyCurrentPrice = 0
    BuyWinningPrice = 0
    SellCurrentCount = 0
    SellWinningCount = 0
    SellCurrentPrice = 0
    SellWinningPrice = 0
    
    
    SQL = "select * from product where PName like '539_4K' order by PID;" '只要顯示539的4K 'SQL = "select * from product where PName like '%4K%' order by PID;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    

    
    
    Open TargetPath For Output As #1
        Print #1, "<table>"
    
        'show report datetime
        Body = "<tr><td>日期</td><td colspan=10>" & Format(txtCurrentDate.Text, "yyyy/MM") & "月</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>產品</td>"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "交易金額</td>"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        Body = Body & "<td>成數</td>"
        Body = Body & "</tr>"
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
                selectFields = "CurrentCount,WinningCount,CurrentDate"
                SQL = "select " & selectFields & " from [order] where [order].PID='" & PIDArray(i) & "' and CID='" & CID & "' and (CurrentDate>='" & Format(txtCurrentDate.Text, "yyyy/MM/") & "01' and CurrentDate<='" & Format(txtCurrentDate.Text, "yyyy/MM/") & Date_Is_28_29_30_31(txtCurrentDate.Text) & "') order by [order].PID;"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)

                'enum custom every product order count
                CurrentCount = 0
                CurrentPrice = 0
                WinningCount = 0
                WinningPrice = 0
                Do Until rec1.EOF
                    'search price, and addition to variable
                    OrderDate = rec1.Fields.Item("CurrentDate")
                    selectFields = "CurrentPrice,WinningPrice,Upset"
                    SQL = "select * from price where PID='" & PIDArray(i) & "' and CID='" & CID & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
                    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
                
                    CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                    WinningCount = WinningCount + Val(rec1.Fields.Item("WinningCount"))
                    If Not price_rec.EOF Then
                        CurrentPrice = CurrentPrice + (CurrentCount * Val(price_rec.Fields.Item("CurrentPrice")))
                        WinningPrice = WinningPrice + (WinningCount * Val(price_rec.Fields.Item("WinningPrice")))
                    End If
                    
                    rec1.MoveNext
                Loop
                Body = Body & "<td>" & CurrentCount & "</td>"
                Body = Body & "<td>" & CurrentPrice & "</td>"
                
                'a new variable to account buy and sell
                If CurrentCount >= 0 Then
                    BuyCurrentCount = BuyCurrentCount + CurrentCount
                    BuyCurrentPrice = BuyCurrentPrice + CurrentPrice
                    BuyWinningCount = BuyWinningCount + WinningCount
                    BuyWinningPrice = BuyWinningPrice + WinningPrice
                Else
                    SellCurrentCount = SellCurrentCount + CurrentCount
                    SellCurrentPrice = SellCurrentPrice + CurrentPrice
                    SellWinningCount = SellWinningCount + WinningCount
                    SellWinningPrice = SellWinningPrice + WinningPrice
                End If
            Next
            
            'add proportion
            SQL = "select * from custom where BonusTarget='" & CID & " " & CName & "';"
            Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
            Do Until rec1.EOF
                BonusTarget = rec1.Fields.Item("CID")
                Proportion = rec1.Fields.Item("Proportion")
                
                'search the bonus fromer order count.
                BonusMoney = 0
                SQL = "select * from [order] where CID='" & BonusTarget & "' and (CurrentDate>='" & Format(txtCurrentDate.Text, "yyyy/MM/") & "01' and CurrentDate<='" & Format(txtCurrentDate.Text, "yyyy/MM/") & Date_Is_28_29_30_31(txtCurrentDate.Text) & "');"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec2)
                Do Until rec2.EOF
                    'search price, and addition to variable
                    OrderDate = rec2.Fields.Item("CurrentDate")
                    BonusProduct = rec2.Fields.Item("PID")
                    selectFields = "CurrentPrice,WinningPrice,Upset"
                    SQL = "select * from price where PID='" & BonusProduct & "' and CID='" & BonusTarget & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
                    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
                    
                    If Not price_rec.EOF Then
                        BonusMoney = BonusMoney + (Val(rec2.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")))
                    End If
                    rec2.MoveNext
                Loop
                
                Body = Body & "<td>來自" & rec1.Fields.Item("CName") & "成數" & Proportion & "%共" & BonusMoney * (Proportion / 100) & "元。</td>"
                rec1.MoveNext
            Loop
            
            Body = Body & "</tr>"
            Print #1, Body
            
            custom_rec.MoveNext
        Loop
        
        
        Body = "<tr></tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計買</td>"
        Body = Body & "<td>" & BuyCurrentCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & BuyCurrentPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計買中</td>"
        Body = Body & "<td>" & BuyWinningCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & BuyWinningPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計賣</td>"
        Body = Body & "<td>" & SellCurrentCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & SellCurrentPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計賣中</td>"
        Body = Body & "<td>" & SellWinningCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & SellWinningPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        
        
        Print #1, "</table>"
    Close #1
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Sub FourKYearReport(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String
    Dim CurrentCount As Integer, CurrentPrice As Double, WinningCount As Integer, WinningPrice As Double
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset, rec2 As New adoDB.Recordset
    Dim BuyCurrentCount As Integer, SellCurrentCount As Integer, BuyCurrentPrice As Double, SellCurrentPrice As Double
    Dim BuyWinningCount As Integer, SellWinningCount As Integer, BuyWinningPrice As Double, SellWinningPrice As Double
    Dim OrderDate As String
    Dim Proportion As Integer, BonusTarget As String, BonusMoney As Double, BonusProduct As String
    Dim CID As String, CName As String
    
    
    BuyCurrentCount = 0
    BuyWinningCount = 0
    BuyCurrentPrice = 0
    BuyWinningPrice = 0
    SellCurrentCount = 0
    SellWinningCount = 0
    SellCurrentPrice = 0
    SellWinningPrice = 0
    
    
    SQL = "select * from product where PName like '539_4K' order by PID;" '只要顯示539的4K 'SQL = "select * from product where PName like '%4K%' order by PID;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    

    
    
    Open TargetPath For Output As #1
        Print #1, "<table>"
    
        'show report datetime
        Body = "<tr><td>日期</td><td colspan=10>" & Format(txtCurrentDate.Text, "yyyy") & "年</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>產品</td>"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "交易金額</td>"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        Body = Body & "<td>成數</td>"
        Body = Body & "</tr>"
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
                selectFields = "CurrentCount,WinningCount,CurrentDate"
                SQL = "select " & selectFields & " from [order] where [order].PID='" & PIDArray(i) & "' and CID='" & CID & "' and (CurrentDate>='" & Format(txtCurrentDate.Text, "yyyy/") & "01/01' and CurrentDate<='" & Format(txtCurrentDate.Text, "yyyy/") & "12/31') order by [order].PID;"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)

                'enum custom every product order count
                CurrentCount = 0
                CurrentPrice = 0
                WinningCount = 0
                WinningPrice = 0
                Do Until rec1.EOF
                    'search price, and addition to variable
                    OrderDate = rec1.Fields.Item("CurrentDate")
                    selectFields = "CurrentPrice,WinningPrice,Upset"
                    SQL = "select * from price where PID='" & PIDArray(i) & "' and CID='" & CID & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
                    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
                
                    CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                    WinningCount = WinningCount + Val(rec1.Fields.Item("WinningCount"))
                    If Not price_rec.EOF Then
                        CurrentPrice = CurrentPrice + (CurrentCount * Val(price_rec.Fields.Item("CurrentPrice")))
                        WinningPrice = WinningPrice + (WinningCount * Val(price_rec.Fields.Item("WinningPrice")))
                    End If
                    
                    rec1.MoveNext
                Loop
                Body = Body & "<td>" & CurrentCount & "</td>"
                Body = Body & "<td>" & CurrentPrice & "</td>"
                
                'a new variable to account buy and sell
                If CurrentCount >= 0 Then
                    BuyCurrentCount = BuyCurrentCount + CurrentCount
                    BuyCurrentPrice = BuyCurrentPrice + CurrentPrice
                    BuyWinningCount = BuyWinningCount + WinningCount
                    BuyWinningPrice = BuyWinningPrice + WinningPrice
                Else
                    SellCurrentCount = SellCurrentCount + CurrentCount
                    SellCurrentPrice = SellCurrentPrice + CurrentPrice
                    SellWinningCount = SellWinningCount + WinningCount
                    SellWinningPrice = SellWinningPrice + WinningPrice
                End If
            Next
            
            'add proportion
            SQL = "select * from custom where BonusTarget='" & CID & " " & CName & "';"
            Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
            Do Until rec1.EOF
                BonusTarget = rec1.Fields.Item("CID")
                Proportion = rec1.Fields.Item("Proportion")
                
                'search the bonus fromer order count.
                BonusMoney = 0
                SQL = "select * from [order] where CID='" & BonusTarget & "' and and (CurrentDate>='" & Format(txtCurrentDate.Text, "yyyy/") & "01/01' and CurrentDate<='" & Format(txtCurrentDate.Text, "yyyy/") & "12/31');"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec2)
                Do Until rec2.EOF
                    'search price, and addition to variable
                    OrderDate = rec2.Fields.Item("CurrentDate")
                    BonusProduct = rec2.Fields.Item("PID")
                    selectFields = "CurrentPrice,WinningPrice,Upset"
                    SQL = "select * from price where PID='" & BonusProduct & "' and CID='" & BonusTarget & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
                    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
                    
                    If Not price_rec.EOF Then
                        BonusMoney = BonusMoney + (Val(rec2.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")))
                    End If
                    rec2.MoveNext
                Loop
                
                Body = Body & "<td>來自" & rec1.Fields.Item("CName") & "成數" & Proportion & "%共" & BonusMoney * (Proportion / 100) & "元。</td>"
                rec1.MoveNext
            Loop
            
            Body = Body & "</tr>"
            Print #1, Body
            
            custom_rec.MoveNext
        Loop
        
        
        Body = "<tr></tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計買</td>"
        Body = Body & "<td>" & BuyCurrentCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & BuyCurrentPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計買中</td>"
        Body = Body & "<td>" & BuyWinningCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & BuyWinningPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計賣</td>"
        Body = Body & "<td>" & SellCurrentCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & SellCurrentPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計賣中</td>"
        Body = Body & "<td>" & SellWinningCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & SellWinningPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        
        
        Print #1, "</table>"
    Close #1
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Sub FourKDayAccount(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String
    Dim CurrentCount As Integer, CurrentPrice As Double, WinningCount As Integer, WinningPrice As Double
    Dim CurrentCountAll(1024) As Long, WinningCountAll(1024) As Long
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset, rec2 As New adoDB.Recordset
    Dim BuyCurrentCount As Integer, SellCurrentCount As Integer, BuyCurrentPrice As Double, SellCurrentPrice As Double
    Dim BuyWinningCount As Integer, SellWinningCount As Integer, BuyWinningPrice As Double, SellWinningPrice As Double
    Dim OrderDate As String
    Dim Proportion As Integer, BonusTarget As String, BonusMoney As Double, BonusProduct As String
    Dim CID As String, CName As String
    
    
    BuyCurrentCount = 0
    BuyWinningCount = 0
    BuyCurrentPrice = 0
    BuyWinningPrice = 0
    SellCurrentCount = 0
    SellWinningCount = 0
    SellCurrentPrice = 0
    SellWinningPrice = 0
    
    
    SQL = "select * from product where PName like '539_4K' order by PID;" '只要顯示539的4K 'SQL = "select * from product where PName like '%4K%' order by PID;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    
    
    Open TargetPath For Output As #1
        Print #1, "<table>"
    
        'show report datetime
        Body = "<tr><td colspan=2>" & txtCurrentDate.Text & "</td><td>日總計</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>產品</td>"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "交易金額</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "中</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "中獎金額</td>"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        Body = Body & "</tr>"
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
                selectFields = "CurrentCount,WinningCount,CurrentDate"
                SQL = "select " & selectFields & " from [order] where [order].PID='" & PIDArray(i) & "' and CID='" & CID & "' and CurrentDate='" & txtCurrentDate.Text & "' order by [order].PID;"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
                 
                'enum custom every product order count
                CurrentCount = 0
                CurrentPrice = 0
                WinningCount = 0
                WinningPrice = 0
                Do Until rec1.EOF
                    'search price, and addition to variable
                    OrderDate = rec1.Fields.Item("CurrentDate")
                    selectFields = "CurrentPrice,WinningPrice,Upset"
                    SQL = "select * from price where PID='" & PIDArray(i) & "' and CID='" & CID & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
                    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
                
                    CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                    WinningCount = WinningCount + Val(rec1.Fields.Item("WinningCount"))
                    If Not price_rec.EOF Then
                        CurrentPrice = CurrentPrice + (CurrentCount * Val(price_rec.Fields.Item("CurrentPrice")))
                        WinningPrice = WinningPrice + (WinningCount * Val(price_rec.Fields.Item("WinningPrice")))
                    End If
                    
                    rec1.MoveNext
                Loop
                Body = Body & "<td>" & CurrentCount & "</td>"
                Body = Body & "<td>" & CurrentPrice & "</td>"
                Body = Body & "<td>" & WinningCount & "</td>"
                Body = Body & "<td>" & WinningPrice & "</td>"
                
                'a new variable to account buy and sell
                If CurrentCount >= 0 Then
                    BuyCurrentCount = BuyCurrentCount + CurrentCount
                    BuyCurrentPrice = BuyCurrentPrice + CurrentPrice
                    BuyWinningCount = BuyWinningCount + WinningCount
                    BuyWinningPrice = BuyWinningPrice + WinningPrice
                Else
                    SellCurrentCount = SellCurrentCount + CurrentCount
                    SellCurrentPrice = SellCurrentPrice + CurrentPrice
                    SellWinningCount = SellWinningCount + WinningCount
                    SellWinningPrice = SellWinningPrice + WinningPrice
                End If
                
                'a new variable to account every product
                CurrentCountAll(i) = CurrentCountAll(i) + CurrentCount
                WinningCountAll(i) = WinningCountAll(i) + WinningCount
            Next
            Body = Body & "</tr>"
            Print #1, Body
            
            custom_rec.MoveNext
        Loop
        
        
        Body = "<tr><td>數量總計</td>"
        For i = 0 To Count - 1
            Body = Body & "<td>" & CurrentCountAll(i) & "</td><td></td>"
            Body = Body & "<td>" & WinningCountAll(i) & "</td><td></td>"
        Next
        Body = Body & "</tr>"
        Print #1, Body
        
        
        Body = "<tr></tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計買</td>"
        Body = Body & "<td>" & BuyCurrentCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & BuyCurrentPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計買中</td>"
        Body = Body & "<td>" & BuyWinningCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & BuyWinningPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計賣</td>"
        Body = Body & "<td>" & SellCurrentCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & SellCurrentPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計賣中</td>"
        Body = Body & "<td>" & SellWinningCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & SellWinningPrice & "</td><td></td>"
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
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String
    Dim CurrentCount As Integer, CurrentPrice As Double, WinningCount As Integer, WinningPrice As Double
    Dim CurrentCountAll(1024) As Long, WinningCountAll(1024) As Long
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset, rec2 As New adoDB.Recordset
    Dim BuyCurrentCount As Integer, SellCurrentCount As Integer, BuyCurrentPrice As Double, SellCurrentPrice As Double
    Dim BuyWinningCount As Integer, SellWinningCount As Integer, BuyWinningPrice As Double, SellWinningPrice As Double
    Dim OrderDate As String
    Dim Proportion As Integer, BonusTarget As String, BonusMoney As Double, BonusProduct As String
    Dim CID As String, CName As String
    
    
    BuyCurrentCount = 0
    BuyWinningCount = 0
    BuyCurrentPrice = 0
    BuyWinningPrice = 0
    SellCurrentCount = 0
    SellWinningCount = 0
    SellCurrentPrice = 0
    SellWinningPrice = 0
    
    
    SQL = "select * from product where PName like '539_4K' order by PID;" '只要顯示539的4K 'SQL = "select * from product where PName like '%4K%' order by PID;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    

    
    
    Open TargetPath For Output As #1
        Print #1, "<table>"
    
        'show report datetime
        Body = "<tr><td colspan=2>" & DateTime.DateAdd("d", -7, txtCurrentDate.Text) & "至" & txtCurrentDate.Text & "</td><td>週總計</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>產品</td>"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "交易金額</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "中</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "中獎金額</td>"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        Body = Body & "<td>成數</td>"
        Body = Body & "<td>前帳</td>"
        Body = Body & "</tr>"
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
                selectFields = "CurrentCount,WinningCount,CurrentDate"
                SQL = "select " & selectFields & " from [order] where [order].PID='" & PIDArray(i) & "' and CID='" & CID & "' and (CurrentDate>='" & Format(DateTime.DateAdd("d", -7, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd") & "' and CurrentDate<='" & txtCurrentDate.Text & "') order by [order].PID;"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)

                'enum custom every product order count
                CurrentCount = 0
                CurrentPrice = 0
                WinningCount = 0
                WinningPrice = 0
                Do Until rec1.EOF
                    'search price, and addition to variable
                    OrderDate = rec1.Fields.Item("CurrentDate")
                    selectFields = "CurrentPrice,WinningPrice,Upset"
                    SQL = "select * from price where PID='" & PIDArray(i) & "' and CID='" & CID & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
                    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
                
                    CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                    WinningCount = WinningCount + Val(rec1.Fields.Item("WinningCount"))
                    If Not price_rec.EOF Then
                        CurrentPrice = CurrentPrice + (CurrentCount * Val(price_rec.Fields.Item("CurrentPrice")))
                        WinningPrice = WinningPrice + (WinningCount * Val(price_rec.Fields.Item("WinningPrice")))
                    End If
                    
                    rec1.MoveNext
                Loop
                Body = Body & "<td>" & CurrentCount & "</td>"
                Body = Body & "<td>" & CurrentPrice & "</td>"
                Body = Body & "<td>" & WinningCount & "</td>"
                Body = Body & "<td>" & WinningPrice & "</td>"
                
                'a new variable to account buy and sell
                If CurrentCount >= 0 Then
                    BuyCurrentCount = BuyCurrentCount + CurrentCount
                    BuyCurrentPrice = BuyCurrentPrice + CurrentPrice
                    BuyWinningCount = BuyWinningCount + WinningCount
                    BuyWinningPrice = BuyWinningPrice + WinningPrice
                Else
                    SellCurrentCount = SellCurrentCount + CurrentCount
                    SellCurrentPrice = SellCurrentPrice + CurrentPrice
                    SellWinningCount = SellWinningCount + WinningCount
                    SellWinningPrice = SellWinningPrice + WinningPrice
                End If
                
                'a new variable to account every product
                CurrentCountAll(i) = CurrentCountAll(i) + CurrentCount
                WinningCountAll(i) = WinningCountAll(i) + WinningCount
            Next
            
            'add proportion
            SQL = "select * from custom where BonusTarget='" & CID & " " & CName & "';"
            Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
            Do Until rec1.EOF
                BonusTarget = rec1.Fields.Item("CID")
                Proportion = rec1.Fields.Item("Proportion")
                
                'search the bonus fromer order count.
                BonusMoney = 0
                SQL = "select * from [order] where CID='" & BonusTarget & "' and (CurrentDate>='" & Format(DateTime.DateAdd("d", -7, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd") & "' and CurrentDate<='" & txtCurrentDate.Text & "');"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec2)
                Do Until rec2.EOF
                    'search price, and addition to variable
                    OrderDate = rec2.Fields.Item("CurrentDate")
                    BonusProduct = rec2.Fields.Item("PID")
                    selectFields = "CurrentPrice,WinningPrice,Upset"
                    SQL = "select * from price where PID='" & BonusProduct & "' and CID='" & BonusTarget & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
                    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
                    
                    If Not price_rec.EOF Then
                        BonusMoney = BonusMoney + (Val(rec2.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")))
                    End If
                    rec2.MoveNext
                Loop
                
                Body = Body & "<td>來自" & rec1.Fields.Item("CName") & "成數" & Proportion & "%共" & BonusMoney * (Proportion / 100) & "元。</td>"
                rec1.MoveNext
            Loop
            
            Body = Body & "</tr>"
            Print #1, Body
            
            custom_rec.MoveNext
        Loop
        
        
        Body = "<tr><td>數量總計</td>"
        For i = 0 To Count - 1
            Body = Body & "<td>" & CurrentCountAll(i) & "</td><td></td>"
            Body = Body & "<td>" & WinningCountAll(i) & "</td><td></td>"
        Next
        Body = Body & "</tr>"
        Print #1, Body
        

        Body = "<tr></tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計買</td>"
        Body = Body & "<td>" & BuyCurrentCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & BuyCurrentPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計買中</td>"
        Body = Body & "<td>" & BuyWinningCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & BuyWinningPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計賣</td>"
        Body = Body & "<td>" & SellCurrentCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & SellCurrentPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計賣中</td>"
        Body = Body & "<td>" & SellWinningCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & SellWinningPrice & "</td><td></td>"
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
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String
    Dim CurrentCount As Integer, CurrentPrice As Double, WinningCount As Integer, WinningPrice As Double
    Dim CurrentCountAll(1024) As Long, WinningCountAll(1024) As Long
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset, rec2 As New adoDB.Recordset
    Dim BuyCurrentCount As Integer, SellCurrentCount As Integer, BuyCurrentPrice As Double, SellCurrentPrice As Double
    Dim BuyWinningCount As Integer, SellWinningCount As Integer, BuyWinningPrice As Double, SellWinningPrice As Double
    Dim OrderDate As String
    Dim Proportion As Integer, BonusTarget As String, BonusMoney As Double, BonusProduct As String
    Dim CID As String, CName As String
    
    
    BuyCurrentCount = 0
    BuyWinningCount = 0
    BuyCurrentPrice = 0
    BuyWinningPrice = 0
    SellCurrentCount = 0
    SellWinningCount = 0
    SellCurrentPrice = 0
    SellWinningPrice = 0
    
    
    SQL = "select * from product where PName like '539_4K' order by PID;" '只要顯示539的4K 'SQL = "select * from product where PName like '%4K%' order by PID;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    

    Open TargetPath For Output As #1
        Print #1, "<table>"
    
        'show report datetime
        Body = "<tr><td colspan=2>" & Format(txtCurrentDate.Text, "yyyy/MM") & "</td><td>月總計</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>產品</td>"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "交易金額</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "中</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "中獎金額</td>"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        Body = Body & "<td>成數</td>"
        Body = Body & "</tr>"
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
                selectFields = "CurrentCount,WinningCount,CurrentDate"
                SQL = "select " & selectFields & " from [order] where [order].PID='" & PIDArray(i) & "' and CID='" & CID & "' and (CurrentDate>='" & Format(txtCurrentDate.Text, "yyyy/MM/") & "01' and CurrentDate<='" & Format(txtCurrentDate.Text, "yyyy/MM/") & Date_Is_28_29_30_31(txtCurrentDate.Text) & "') order by [order].PID;"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
                
                'enum custom every product order count
                CurrentCount = 0
                CurrentPrice = 0
                WinningCount = 0
                WinningPrice = 0
                Do Until rec1.EOF
                    'search price, and addition to variable
                    OrderDate = rec1.Fields.Item("CurrentDate")
                    selectFields = "CurrentPrice,WinningPrice,Upset"
                    SQL = "select * from price where PID='" & PIDArray(i) & "' and CID='" & CID & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
                    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
                
                    CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                    WinningCount = WinningCount + Val(rec1.Fields.Item("WinningCount"))
                    If Not price_rec.EOF Then
                        CurrentPrice = CurrentPrice + (CurrentCount * Val(price_rec.Fields.Item("CurrentPrice")))
                        WinningPrice = WinningPrice + (WinningCount * Val(price_rec.Fields.Item("WinningPrice")))
                    End If
                    
                    rec1.MoveNext
                Loop
                Body = Body & "<td>" & CurrentCount & "</td>"
                Body = Body & "<td>" & CurrentPrice & "</td>"
                Body = Body & "<td>" & WinningCount & "</td>"
                Body = Body & "<td>" & WinningPrice & "</td>"
                
                'a new variable to account buy and sell
                If CurrentCount >= 0 Then
                    BuyCurrentCount = BuyCurrentCount + CurrentCount
                    BuyCurrentPrice = BuyCurrentPrice + CurrentPrice
                    BuyWinningCount = BuyWinningCount + WinningCount
                    BuyWinningPrice = BuyWinningPrice + WinningPrice
                Else
                    SellCurrentCount = SellCurrentCount + CurrentCount
                    SellCurrentPrice = SellCurrentPrice + CurrentPrice
                    SellWinningCount = SellWinningCount + WinningCount
                    SellWinningPrice = SellWinningPrice + WinningPrice
                End If
        
                'a new variable to account every product
                CurrentCountAll(i) = CurrentCountAll(i) + CurrentCount
                WinningCountAll(i) = WinningCountAll(i) + WinningCount
            Next
            
            'add proportion
            SQL = "select * from custom where BonusTarget='" & CID & " " & CName & "';"
            Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
            Do Until rec1.EOF
                BonusTarget = rec1.Fields.Item("CID")
                Proportion = rec1.Fields.Item("Proportion")
                
                'search the bonus fromer order count.
                BonusMoney = 0
                SQL = "select * from [order] where CID='" & BonusTarget & "' and (CurrentDate>='" & Format(txtCurrentDate.Text, "yyyy/MM/") & "01' and CurrentDate<='" & Format(txtCurrentDate.Text, "yyyy/MM/") & Date_Is_28_29_30_31(txtCurrentDate.Text) & "');"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec2)
                Do Until rec2.EOF
                    'search price, and addition to variable
                    OrderDate = rec2.Fields.Item("CurrentDate")
                    BonusProduct = rec2.Fields.Item("PID")
                    selectFields = "CurrentPrice,WinningPrice,Upset"
                    SQL = "select * from price where PID='" & BonusProduct & "' and CID='" & BonusTarget & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
                    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
                    
                    If Not price_rec.EOF Then
                        BonusMoney = BonusMoney + (Val(rec2.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")))
                    End If
                    rec2.MoveNext
                Loop
                
                Body = Body & "<td>來自" & rec1.Fields.Item("CName") & "成數" & Proportion & "%共" & BonusMoney * (Proportion / 100) & "元。</td>"
                rec1.MoveNext
            Loop
            
            Body = Body & "</tr>"
            Print #1, Body
            
            custom_rec.MoveNext
        Loop
        
        
        Body = "<tr><td>數量總計</td>"
        For i = 0 To Count - 1
            Body = Body & "<td>" & CurrentCountAll(i) & "</td><td></td>"
            Body = Body & "<td>" & WinningCountAll(i) & "</td><td></td>"
        Next
        Body = Body & "</tr>"
        Print #1, Body


        Body = "<tr></tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計買</td>"
        Body = Body & "<td>" & BuyCurrentCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & BuyCurrentPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計買中</td>"
        Body = Body & "<td>" & BuyWinningCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & BuyWinningPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計賣</td>"
        Body = Body & "<td>" & SellCurrentCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & SellCurrentPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計賣中</td>"
        Body = Body & "<td>" & SellWinningCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & SellWinningPrice & "</td><td></td>"
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
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String
    Dim CurrentCount As Integer, CurrentPrice As Double, WinningCount As Integer, WinningPrice As Double
    Dim CurrentCountAll(1024) As Long, WinningCountAll(1024) As Long
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset, rec2 As New adoDB.Recordset
    Dim BuyCurrentCount As Integer, SellCurrentCount As Integer, BuyCurrentPrice As Double, SellCurrentPrice As Double
    Dim BuyWinningCount As Integer, SellWinningCount As Integer, BuyWinningPrice As Double, SellWinningPrice As Double
    Dim OrderDate As String
    Dim Proportion As Integer, BonusTarget As String, BonusMoney As Double, BonusProduct As String
    Dim CID As String, CName As String
    
    
    BuyCurrentCount = 0
    BuyWinningCount = 0
    BuyCurrentPrice = 0
    BuyWinningPrice = 0
    SellCurrentCount = 0
    SellWinningCount = 0
    SellCurrentPrice = 0
    SellWinningPrice = 0
    
    
    SQL = "select * from product where PName like '539_4K' order by PID;" '只要顯示539的4K 'SQL = "select * from product where PName like '%4K%' order by PID;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    
    
    Open TargetPath For Output As #1
        Print #1, "<table>"
    
        'show report datetime
        Body = "<tr><td colspan=2>" & Format(txtCurrentDate.Text, "yyyy") & "</td><td>年總計</td></tr>"
        Print #1, Body
        
        'show product name
        Body = "<tr><td>產品</td>"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "交易金額</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "中</td>"
            Body = Body & "<td>" & product_rec.Fields.Item("PName") & "中獎金額</td>"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        Body = Body & "<td>成數</td>"
        Body = Body & "</tr>"
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
                selectFields = "CurrentCount,WinningCount,CurrentDate"
                SQL = "select " & selectFields & " from [order] where [order].PID='" & PIDArray(i) & "' and CID='" & CID & "' and (CurrentDate>='" & Format(txtCurrentDate.Text, "yyyy/") & "01/01' and CurrentDate<='" & Format(txtCurrentDate.Text, "yyyy/") & "12/31') order by [order].PID;"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
                
                'enum custom every product order count
                CurrentCount = 0
                CurrentPrice = 0
                WinningCount = 0
                WinningPrice = 0
                Do Until rec1.EOF
                    'search price, and addition to variable
                    OrderDate = rec1.Fields.Item("CurrentDate")
                    selectFields = "CurrentPrice,WinningPrice,Upset"
                    SQL = "select * from price where PID='" & PIDArray(i) & "' and CID='" & CID & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
                    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
                
                    CurrentCount = CurrentCount + Val(rec1.Fields.Item("CurrentCount"))
                    WinningCount = WinningCount + Val(rec1.Fields.Item("WinningCount"))
                    If Not price_rec.EOF Then
                        CurrentPrice = CurrentPrice + (CurrentCount * Val(price_rec.Fields.Item("CurrentPrice")))
                        WinningPrice = WinningPrice + (WinningCount * Val(price_rec.Fields.Item("WinningPrice")))
                    End If
                    
                    rec1.MoveNext
                Loop
                Body = Body & "<td>" & CurrentCount & "</td>"
                Body = Body & "<td>" & CurrentPrice & "</td>"
                Body = Body & "<td>" & WinningCount & "</td>"
                Body = Body & "<td>" & WinningPrice & "</td>"
                
                'a new variable to account buy and sell
                If CurrentCount >= 0 Then
                    BuyCurrentCount = BuyCurrentCount + CurrentCount
                    BuyCurrentPrice = BuyCurrentPrice + CurrentPrice
                    BuyWinningCount = BuyWinningCount + WinningCount
                    BuyWinningPrice = BuyWinningPrice + WinningPrice
                Else
                    SellCurrentCount = SellCurrentCount + CurrentCount
                    SellCurrentPrice = SellCurrentPrice + CurrentPrice
                    SellWinningCount = SellWinningCount + WinningCount
                    SellWinningPrice = SellWinningPrice + WinningPrice
                End If
                
                'a new variable to account every product
                CurrentCountAll(i) = CurrentCountAll(i) + CurrentCount
                WinningCountAll(i) = WinningCountAll(i) + WinningCount
            Next
            
            'add proportion
            SQL = "select * from custom where BonusTarget='" & CID & " " & CName & "';"
            Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
            Do Until rec1.EOF
                BonusTarget = rec1.Fields.Item("CID")
                Proportion = rec1.Fields.Item("Proportion")
                
                'search the bonus fromer order count.
                BonusMoney = 0
                SQL = "select * from [order] where CID='" & BonusTarget & "' and (CurrentDate>='" & Format(txtCurrentDate.Text, "yyyy/") & "01/01' and CurrentDate<='" & Format(txtCurrentDate.Text, "yyyy/") & "12/31');"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec2)
                Do Until rec2.EOF
                    'search price, and addition to variable
                    OrderDate = rec2.Fields.Item("CurrentDate")
                    BonusProduct = rec2.Fields.Item("PID")
                    selectFields = "CurrentPrice,WinningPrice,Upset"
                    SQL = "select * from price where PID='" & BonusProduct & "' and CID='" & BonusTarget & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
                    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
                    
                    If Not price_rec.EOF Then
                        BonusMoney = BonusMoney + (Val(rec2.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")))
                    End If
                    rec2.MoveNext
                Loop
                
                Body = Body & "<td>來自" & rec1.Fields.Item("CName") & "成數" & Proportion & "%共" & BonusMoney * (Proportion / 100) & "元。</td>"
                rec1.MoveNext
            Loop
            
            Body = Body & "</tr>"
            Print #1, Body
            
            custom_rec.MoveNext
        Loop
        
        
        Body = "<tr><td>數量總計</td>"
        For i = 0 To Count - 1
            Body = Body & "<td>" & CurrentCountAll(i) & "</td><td></td>"
            Body = Body & "<td>" & WinningCountAll(i) & "</td><td></td>"
        Next
        Body = Body & "</tr>"
        Print #1, Body
        

        Body = "<tr></tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計買</td>"
        Body = Body & "<td>" & BuyCurrentCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & BuyCurrentPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計買中</td>"
        Body = Body & "<td>" & BuyWinningCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & BuyWinningPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計賣</td>"
        Body = Body & "<td>" & SellCurrentCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & SellCurrentPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        Body = "<tr>"
        Body = Body & "<td>小計賣中</td>"
        Body = Body & "<td>" & SellWinningCount & "</td><td></td>"
        Body = Body & "<td>金額</td>"
        Body = Body & "<td>" & SellWinningPrice & "</td><td></td>"
        Body = Body & "</tr>"
        Print #1, Body
        
        
        Print #1, "</table>"
    Close #1
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub


Sub CustromProductDayReport(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer
    Dim CurrentCount As Integer, CurrentPrice As Double, WinningCount As Integer, WinningPrice As Double
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    Dim CData() As String
    Dim PData() As String
    Dim beginv As Integer
    Dim endv As Integer
    Dim OrderDate As String, ProductID As String, ProductName As String
    Dim PriceCount As Double
    
        
    'search order
    CData = Split(cmbCName.Text, " ")
    PData = Split(cmbPName.Text, " ")
    If Val(PData(0)) >= 100 Then
        beginv = Mid(PData(0), 2, 1)
        endv = beginv & "9"
        beginv = beginv & "0"
        SQL = "select * from [order] where cint(PID)>=" & beginv & " and cint(PID)<=" & endv & " and CID='" & CData(0) & "' and CurrentDate='" & txtCurrentDate.Text & "' order by CurrentDate;"
    Else
        SQL = "select * from [order] where PID='" & PData(0) & "' and CID='" & CData(0) & "' and CurrentDate='" & txtCurrentDate.Text & "' order by CurrentDate;"
    End If
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, order_rec)
       
    
    Open TargetPath For Output As #1
        Print #1, "<table>"
    
        'show report datetime
        Body = "<tr><td>日期</td><td colspan=10>" & txtCurrentDate.Text & "</td></tr>"
        Print #1, Body
        
        'show custom name and product name
        Body = "<tr><td>客戶名稱</td><td colspan=10>" & CData(1) & "</td></tr>"
        Print #1, Body
        Body = "<tr><td>產品名稱</td><td colspan=10>" & PData(1) & "</td></tr>"
        Print #1, Body
        Body = "<tr><td>交易流水號</td><td>交易日期</td><td>產品名稱</td><td>交易數量</td><td>交易金額</td><td>中獎數量</td><td>中獎金額</td><td>漲價</td><td>退水金額</td><td>備註</td></tr>"
        Print #1, Body
        
        'show every order with current date
        PriceCount = 0
        'order_rec.MoveFirst
        Do Until order_rec.EOF
            'search price
            OrderDate = order_rec.Fields.Item("CurrentDate")
            ProductID = order_rec.Fields.Item("PID")
            selectFields = "CurrentPrice,WinningPrice,Upset"
            SQL = "select * from price where PID='" & ProductID & "' and CID='" & CData(0) & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
            Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
            
            'search product
            SQL = "select * from product where PID='" & ProductID & "';"
            Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
            ProductName = product_rec.Fields.Item("PName")

            'mark custom name
            Body = "<tr>"
            Body = Body & "<td>" & order_rec.Fields.Item("SwiftCode") & "</td>"
            Body = Body & "<td>" & OrderDate & "</td>"
            Body = Body & "<td>" & ProductName & "</td>"
            
            Body = Body & "<td>" & order_rec.Fields.Item("CurrentCount") & "</td>"
            If price_rec.EOF Then
                Body = Body & "<td>0</td>"
            Else
                Body = Body & "<td>" & order_rec.Fields.Item("CurrentCount") * Val(price_rec.Fields.Item("CurrentPrice")) & "</td>"
                PriceCount = PriceCount + (order_rec.Fields.Item("CurrentCount") * Val(price_rec.Fields.Item("CurrentPrice")))
            End If
            
            Body = Body & "<td>" & order_rec.Fields.Item("WinningCount") & "</td>"
            If price_rec.EOF Then
                Body = Body & "<td>0</td>"
            Else
                Body = Body & "<td>" & order_rec.Fields.Item("WinningCount") * Val(price_rec.Fields.Item("WinningPrice")) & "</td>"
                PriceCount = PriceCount - (order_rec.Fields.Item("WinningCount") * Val(price_rec.Fields.Item("WinningPrice")))
            End If
            
            Body = Body & "<td>" & order_rec.Fields.Item("AddMoney") & "</td>"
            If Not IsNull(order_rec.Fields.Item("AddMoney")) Then PriceCount = PriceCount + Val(order_rec.Fields.Item("AddMoney"))
            
            Body = Body & "<td>" & order_rec.Fields.Item("BonusMoney") & "</td>"
            '不計退水金額 'If Not IsNull(order_rec.Fields.Item("BonusMoney")) Then PriceCount = PriceCount - order_rec.Fields.Item("BonusMoney")
            
            Body = Body & "<td>" & order_rec.Fields.Item("Note") & "</td>"
              

            Body = Body & "</tr>"
            Print #1, Body
            
            order_rec.MoveNext
        Loop
        
        'show price count
        Body = "<tr><td>應收</td><td>" & PriceCount & "</td></tr>"
        Print #1, Body
        
        Print #1, "</table>"
    Close #1
    
    order_rec.Close
End Sub

Sub CustromProductWeekReport(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer
    Dim CurrentCount As Integer, CurrentPrice As Double, WinningCount As Integer, WinningPrice As Double
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    Dim CData() As String
    Dim PData() As String
    Dim beginv As Integer
    Dim endv As Integer
    Dim OrderDate As String, ProductID As String, ProductName As String
    Dim PriceCount As Double
    
       
    'search order
    CData = Split(cmbCName.Text, " ")
    PData = Split(cmbPName.Text, " ")
    If Val(PData(0)) >= 100 Then
        beginv = Mid(PData(0), 2, 1)
        endv = beginv & "9"
        beginv = beginv & "0"
        SQL = "select * from [order] where cint(PID)>=" & beginv & " and cint(PID)<=" & endv & " and CID='" & CData(0) & "' and (CurrentDate>='" & Format(DateTime.DateAdd("d", -7, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd") & "' and CurrentDate<='" & txtCurrentDate.Text & "') order by CurrentDate;"
    Else
        SQL = "select * from [order] where PID='" & PData(0) & "' and CID='" & CData(0) & "' and (CurrentDate>='" & Format(DateTime.DateAdd("d", -7, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd") & "' and CurrentDate<='" & txtCurrentDate.Text & "') order by CurrentDate;"
    End If
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, order_rec)
  
    
    Open TargetPath For Output As #1
        Print #1, "<table>"
    
        'show report datetime
        Body = "<tr><td>日期</td><td colspan=10>" & DateTime.DateAdd("d", -7, txtCurrentDate.Text) & "至" & txtCurrentDate.Text & "</td></tr>"
        Print #1, Body
        
        'show custom name and product name
        Body = "<tr><td>客戶名稱</td><td colspan=10>" & CData(1) & "</td></tr>"
        Print #1, Body
        Body = "<tr><td>產品名稱</td><td colspan=10>" & PData(1) & "</td></tr>"
        Print #1, Body
        Body = "<tr><td>交易流水號</td><td>交易日期</td><td>產品名稱</td><td>交易數量</td><td>交易金額</td><td>中獎數量</td><td>中獎金額</td><td>漲價</td><td>退水金額</td><td>備註</td></tr>"
        Print #1, Body
        
        'show every order with current date
        PriceCount = 0
        'order_rec.MoveFirst
        Do Until order_rec.EOF
            'search price
            OrderDate = order_rec.Fields.Item("CurrentDate")
            ProductID = order_rec.Fields.Item("PID")
            selectFields = "CurrentPrice,WinningPrice,Upset"
            SQL = "select * from price where PID='" & ProductID & "' and CID='" & CData(0) & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
            Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
            
            'search product
            SQL = "select * from product where PID='" & ProductID & "';"
            Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
            ProductName = product_rec.Fields.Item("PName")
        
            'mark custom name
            Body = "<tr>"
            Body = Body & "<td>" & order_rec.Fields.Item("SwiftCode") & "</td>"
            Body = Body & "<td>" & OrderDate & "</td>"
            Body = Body & "<td>" & ProductName & "</td>"
            
            Body = Body & "<td>" & order_rec.Fields.Item("CurrentCount") & "</td>"
            If price_rec.EOF Then
                Body = Body & "<td>0</td>"
            Else
                Body = Body & "<td>" & order_rec.Fields.Item("CurrentCount") * Val(price_rec.Fields.Item("CurrentPrice")) & "</td>"
                PriceCount = PriceCount + (order_rec.Fields.Item("CurrentCount") * Val(price_rec.Fields.Item("CurrentPrice")))
            End If
            
            Body = Body & "<td>" & order_rec.Fields.Item("WinningCount") & "</td>"
            If price_rec.EOF Then
                Body = Body & "<td>0</td>"
            Else
                Body = Body & "<td>" & order_rec.Fields.Item("WinningCount") * Val(price_rec.Fields.Item("WinningPrice")) & "</td>"
                PriceCount = PriceCount - (order_rec.Fields.Item("WinningCount") * Val(price_rec.Fields.Item("WinningPrice")))
            End If
            
            Body = Body & "<td>" & order_rec.Fields.Item("AddMoney") & "</td>"
            If Not IsNull(order_rec.Fields.Item("AddMoney")) Then PriceCount = PriceCount + Val(order_rec.Fields.Item("AddMoney"))
            
            Body = Body & "<td>" & order_rec.Fields.Item("BonusMoney") & "</td>"
            '不計退水金額 'If Not IsNull(order_rec.Fields.Item("BonusMoney")) Then PriceCount = PriceCount - order_rec.Fields.Item("BonusMoney")
            
            Body = Body & "<td>" & order_rec.Fields.Item("Note") & "</td>"
              

            Body = Body & "</tr>"
            Print #1, Body
            
            order_rec.MoveNext
        Loop
        
        'show price count
        Body = "<tr><td>應收</td><td>" & PriceCount & "</td></tr>"
        Print #1, Body
        
        Print #1, "</table>"
    Close #1
    
    order_rec.Close
End Sub

Sub CustromProductWeekTransaction(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer
    Dim CurrentCount As Integer, CurrentPrice As Double, WinningCount As Integer, WinningPrice As Double
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    Dim CData() As String
    Dim PData() As String
    Dim beginv As Integer
    Dim endv As Integer
    Dim OrderDate As String, ProductID As String, ProductName As String
    Dim PriceCount As Double
    Dim OldOrderDate As String, CurrentPriceCount As Double, WinningPriceCount As Double, AddMoney As Double, BonusMoney As Double
    Dim DayDiff As Integer
       
       
    'search order
    CData = Split(cmbCName.Text, " ")
    PData = Split(cmbPName.Text, " ")
    If Val(PData(0)) >= 100 Then
        beginv = Mid(PData(0), 2, 1)
        endv = beginv & "9"
        beginv = beginv & "0"
        SQL = "select * from [order] where cint(PID)>=" & beginv & " and cint(PID)<=" & endv & " and CID='" & CData(0) & "' and (CurrentDate>='" & Format(DateTime.DateAdd("d", -7, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd") & "' and CurrentDate<='" & txtCurrentDate.Text & "') order by CurrentDate;"
    Else
        SQL = "select * from [order] where PID='" & PData(0) & "' and CID='" & CData(0) & "' and (CurrentDate>='" & Format(DateTime.DateAdd("d", -7, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd") & "' and CurrentDate<='" & txtCurrentDate.Text & "') order by CurrentDate;"
    End If
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, order_rec)
  
    
    Open TargetPath For Output As #1
        Print #1, "<table>"
    
        'show report datetime
        Body = "<tr><td>日期</td><td colspan=10>" & DateTime.DateAdd("d", -7, txtCurrentDate.Text) & "至" & txtCurrentDate.Text & "</td></tr>"
        Print #1, Body
        
        'show custom name and product name
        Body = "<tr><td>客戶名稱</td><td colspan=10>" & CData(1) & "</td></tr>"
        Print #1, Body
        Body = "<tr><td>產品名稱</td><td colspan=10>" & PData(1) & "</td></tr>"
        Print #1, Body
        Body = "<tr><td>交易日期</td><td>交易金額</td><td>中獎金額</td><td>漲價</td><td>退水金額</td><td>小計</td></tr>"
        Print #1, Body
        
        'show every order with current date
        PriceCount = 0
        CurrentPriceCount = 0
        WinningPriceCount = 0
        AddMoney = 0
        BonusMoney = 0
        'order_rec.MoveFirst
        Do Until order_rec.EOF
            'search price
            OrderDate = order_rec.Fields.Item("CurrentDate")
            ProductID = order_rec.Fields.Item("PID")
            selectFields = "CurrentPrice,WinningPrice,Upset"
            SQL = "select * from price where PID='" & ProductID & "' and CID='" & CData(0) & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
            Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
            
            
            If OldOrderDate = OrderDate Or OldOrderDate = "" Then
                'fill data as first
                If OldOrderDate = "" And OrderDate <> "" Then
                    DayDiff = DateDiff("d", Format(DateTime.DateAdd("d", -7, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd"), OrderDate)
                    For i = 0 To DayDiff - 1
                        Body = "<tr>"
                        Body = Body & "<td>" & DateTime.DateAdd("d", i, Format(DateTime.DateAdd("d", -7, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd")) & "</td>"
                        Body = Body & "<td>0</td>"
                        Body = Body & "<td>0</td>"
                        Body = Body & "<td>0</td>"
                        Body = Body & "<td>0</td>"
                        Body = Body & "<td>0</td>"
                        Body = Body & "</tr>"
                        Print #1, Body
                    Next
                End If
                
                'save data
                If Not price_rec.EOF Then
                    CurrentPriceCount = CurrentPriceCount + (order_rec.Fields.Item("CurrentCount") * Val(price_rec.Fields.Item("CurrentPrice")))
                    WinningPriceCount = WinningPriceCount + (order_rec.Fields.Item("WinningCount") * Val(price_rec.Fields.Item("WinningPrice")))
                End If
                If Not IsNull(order_rec.Fields.Item("AddMoney")) Then AddMoney = AddMoney + Val(order_rec.Fields.Item("AddMoney"))
                If Not IsNull(order_rec.Fields.Item("BonusMoney")) Then BonusMoney = BonusMoney + Val(order_rec.Fields.Item("BonusMoney"))
            Else
                'mark custom name
                Body = "<tr>"
                Body = Body & "<td>" & OldOrderDate & "</td>"
                Body = Body & "<td>" & CurrentPriceCount & "</td>"
                Body = Body & "<td>" & WinningPriceCount & "</td>"
                Body = Body & "<td>" & AddMoney & "</td>"
                Body = Body & "<td>" & BonusMoney & "</td>"
                Body = Body & "<td>" & CurrentPriceCount + WinningPriceCount + AddMoney + BonusMoney & "</td>"
                PriceCount = PriceCount + CurrentPriceCount + WinningPriceCount + AddMoney + BonusMoney
                Body = Body & "</tr>"
                Print #1, Body
                
                'clear variable data
                CurrentPriceCount = 0
                WinningPriceCount = 0
                AddMoney = 0
                BonusMoney = 0
                
                'fill data per day
                DayDiff = DateDiff("d", OldOrderDate, OrderDate)
                For i = 1 To DayDiff - 1
                    Body = "<tr>"
                    Body = Body & "<td>" & DateTime.DateAdd("d", i, OldOrderDate) & "</td>"
                    Body = Body & "<td>0</td>"
                    Body = Body & "<td>0</td>"
                    Body = Body & "<td>0</td>"
                    Body = Body & "<td>0</td>"
                    Body = Body & "<td>0</td>"
                    Body = Body & "</tr>"
                    Print #1, Body
                Next
            End If
            OldOrderDate = OrderDate
            
            
            order_rec.MoveNext
        Loop
        
        
        
        'show the last data and fill data to lastday as last
        'if everything is empty, then setting a startup date for full data
        If OrderDate = "" Then OrderDate = Format(DateTime.DateAdd("d", -7, Format(txtCurrentDate.Text, "yyyy/MM/dd")))
        If OldOrderDate = "" Then OldOrderDate = OrderDate
        
        'mark custom name
        Body = "<tr>"
        Body = Body & "<td>" & OldOrderDate & "</td>"
        Body = Body & "<td>" & CurrentPriceCount & "</td>"
        Body = Body & "<td>" & WinningPriceCount & "</td>"
        Body = Body & "<td>" & AddMoney & "</td>"
        Body = Body & "<td>" & BonusMoney & "</td>"
        Body = Body & "<td>" & CurrentPriceCount + WinningPriceCount + AddMoney + BonusMoney & "</td>"
        PriceCount = PriceCount + CurrentPriceCount + WinningPriceCount + AddMoney + BonusMoney
        Body = Body & "</tr>"
        Print #1, Body
        DayDiff = DateDiff("d", OrderDate, txtCurrentDate.Text)
        For i = 1 To DayDiff - 1
            Body = "<tr>"
            Body = Body & "<td>" & DateTime.DateAdd("d", i, OrderDate) & "</td>"
            Body = Body & "<td>0</td>"
            Body = Body & "<td>0</td>"
            Body = Body & "<td>0</td>"
            Body = Body & "<td>0</td>"
            Body = Body & "<td>0</td>"
            Body = Body & "</tr>"
            Print #1, Body
        Next
        
        
        'show price count
        Body = "<tr><td>總計</td><td>" & PriceCount & "</td></tr>"
        Print #1, Body
        
        Print #1, "</table>"
    Close #1
    
    order_rec.Close
End Sub

Sub CustromWeekReport(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer
    Dim CurrentCount As Integer, CurrentPrice As Double, WinningCount As Integer, WinningPrice As Double
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    Dim CData() As String
    Dim PData() As String
    Dim beginv As Integer
    Dim endv As Integer
    Dim OrderDate As String, ProductID As String, ProductName As String
    Dim PriceCount As Double
    
    
    'search order
    CData = Split(cmbCName.Text, " ")
    SQL = "select * from [order] where CID='" & CData(0) & "' and (CurrentDate>='" & Format(DateTime.DateAdd("d", -7, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd") & "' and CurrentDate<='" & txtCurrentDate.Text & "') order by CurrentDate;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, order_rec)
  
    
    Open TargetPath For Output As #1
        Print #1, "<table>"
    
        'show report datetime
        Body = "<tr><td>日期</td><td colspan=10>" & DateTime.DateAdd("d", -7, txtCurrentDate.Text) & "至" & txtCurrentDate.Text & "</td></tr>"
        Print #1, Body
        
        'show custom name and product name
        Body = "<tr><td>客戶名稱</td><td colspan=10>" & CData(1) & "</td></tr>"
        Print #1, Body
        Body = "<tr><td>交易流水號</td><td>交易日期</td><td>產品名稱</td><td>交易數量</td><td>交易金額</td><td>中獎數量</td><td>中獎金額</td><td>漲價</td><td>退水金額</td><td>備註</td></tr>"
        Print #1, Body
        
        'show every order with current date
        PriceCount = 0
        'order_rec.MoveFirst
        Do Until order_rec.EOF
            'search price
            OrderDate = order_rec.Fields.Item("CurrentDate")
            ProductID = order_rec.Fields.Item("PID")
            selectFields = "CurrentPrice,WinningPrice,Upset"
            SQL = "select * from price where PID='" & ProductID & "' and CID='" & CData(0) & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
            Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
            
            'search product
            SQL = "select * from product where PID='" & ProductID & "';"
            Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
            ProductName = product_rec.Fields.Item("PName")
        
            'mark custom name
            Body = "<tr>"
            Body = Body & "<td>" & order_rec.Fields.Item("SwiftCode") & "</td>"
            Body = Body & "<td>" & OrderDate & "</td>"
            Body = Body & "<td>" & ProductName & "</td>"
            
            Body = Body & "<td>" & order_rec.Fields.Item("CurrentCount") & "</td>"
            If price_rec.EOF Then
                Body = Body & "<td>0</td>"
            Else
                Body = Body & "<td>" & order_rec.Fields.Item("CurrentCount") * Val(price_rec.Fields.Item("CurrentPrice")) & "</td>"
                PriceCount = PriceCount + (order_rec.Fields.Item("CurrentCount") * Val(price_rec.Fields.Item("CurrentPrice")))
            End If
            
            Body = Body & "<td>" & order_rec.Fields.Item("WinningCount") & "</td>"
            If price_rec.EOF Then
                Body = Body & "<td>0</td>"
            Else
                Body = Body & "<td>" & order_rec.Fields.Item("WinningCount") * Val(price_rec.Fields.Item("WinningPrice")) & "</td>"
                PriceCount = PriceCount - (order_rec.Fields.Item("WinningCount") * Val(price_rec.Fields.Item("WinningPrice")))
            End If
            
            Body = Body & "<td>" & order_rec.Fields.Item("AddMoney") & "</td>"
            If Not IsNull(order_rec.Fields.Item("AddMoney")) Then PriceCount = PriceCount + Val(order_rec.Fields.Item("AddMoney"))
            
            Body = Body & "<td>" & order_rec.Fields.Item("BonusMoney") & "</td>"
            '不計退水金額 'If Not IsNull(order_rec.Fields.Item("BonusMoney")) Then PriceCount = PriceCount - order_rec.Fields.Item("BonusMoney")
            
            Body = Body & "<td>" & order_rec.Fields.Item("Note") & "</td>"
              

            Body = Body & "</tr>"
            Print #1, Body
            
            order_rec.MoveNext
        Loop
        
        'show price count
        Body = "<tr><td>應收</td><td>" & PriceCount & "</td></tr>"
        Print #1, Body
        
        Print #1, "</table>"
    Close #1
    
    order_rec.Close
End Sub

Sub CustromMonthReport(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer
    Dim CurrentCount As Integer, CurrentPrice As Double, WinningCount As Integer, WinningPrice As Double
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    Dim CData() As String
    Dim PData() As String
    Dim beginv As Integer
    Dim endv As Integer
    Dim OrderDate As String, ProductID As String, ProductName As String
    Dim PriceCount As Double
    
    
    'search order
    CData = Split(cmbCName.Text, " ")
    SQL = "select * from [order] where CID='" & CData(0) & "' and (CurrentDate>='" & Format(txtCurrentDate.Text, "yyyy/MM/") & "01' and CurrentDate<='" & Format(txtCurrentDate.Text, "yyyy/MM/") & Date_Is_28_29_30_31(txtCurrentDate.Text) & "') order by CurrentDate;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, order_rec)
  
    
    Open TargetPath For Output As #1
        Print #1, "<table>"
    
        'show report datetime
        Body = "<tr><td>日期</td><td colspan=10>" & Format(txtCurrentDate.Text, "yyyy/MM") & "</td></tr>"
        Print #1, Body
        
        'show custom name and product name
        Body = "<tr><td>客戶名稱</td><td colspan=10>" & CData(1) & "</td></tr>"
        Print #1, Body
        Body = "<tr><td>交易流水號</td><td>交易日期</td><td>產品名稱</td><td>交易數量</td><td>交易金額</td><td>中獎數量</td><td>中獎金額</td><td>漲價</td><td>退水金額</td><td>備註</td></tr>"
        Print #1, Body
        
        'show every order with current date
        PriceCount = 0
        'order_rec.MoveFirst
        Do Until order_rec.EOF
            'search price
            OrderDate = order_rec.Fields.Item("CurrentDate")
            ProductID = order_rec.Fields.Item("PID")
            selectFields = "CurrentPrice,WinningPrice,Upset"
            SQL = "select * from price where PID='" & ProductID & "' and CID='" & CData(0) & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
            Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
            
            'search product
            SQL = "select * from product where PID='" & ProductID & "';"
            Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
            ProductName = product_rec.Fields.Item("PName")
        
            'mark custom name
            Body = "<tr>"
            Body = Body & "<td>" & order_rec.Fields.Item("SwiftCode") & "</td>"
            Body = Body & "<td>" & OrderDate & "</td>"
            Body = Body & "<td>" & ProductName & "</td>"
            
            Body = Body & "<td>" & order_rec.Fields.Item("CurrentCount") & "</td>"
            If price_rec.EOF Then
                Body = Body & "<td>0</td>"
            Else
                Body = Body & "<td>" & order_rec.Fields.Item("CurrentCount") * Val(price_rec.Fields.Item("CurrentPrice")) & "</td>"
                PriceCount = PriceCount + (order_rec.Fields.Item("CurrentCount") * Val(price_rec.Fields.Item("CurrentPrice")))
            End If
            
            Body = Body & "<td>" & order_rec.Fields.Item("WinningCount") & "</td>"
            If price_rec.EOF Then
                Body = Body & "<td>0</td>"
            Else
                Body = Body & "<td>" & order_rec.Fields.Item("WinningCount") * Val(price_rec.Fields.Item("WinningPrice")) & "</td>"
                PriceCount = PriceCount - (order_rec.Fields.Item("WinningCount") * Val(price_rec.Fields.Item("WinningPrice")))
            End If
            
            Body = Body & "<td>" & order_rec.Fields.Item("AddMoney") & "</td>"
            If Not IsNull(order_rec.Fields.Item("AddMoney")) Then PriceCount = PriceCount + Val(order_rec.Fields.Item("AddMoney"))
            
            Body = Body & "<td>" & order_rec.Fields.Item("BonusMoney") & "</td>"
            '不計退水金額 'If Not IsNull(order_rec.Fields.Item("BonusMoney")) Then PriceCount = PriceCount - order_rec.Fields.Item("BonusMoney")
            
            Body = Body & "<td>" & order_rec.Fields.Item("Note") & "</td>"
              

            Body = Body & "</tr>"
            Print #1, Body
            
            order_rec.MoveNext
        Loop
        
        'show price count
        Body = "<tr><td>應收</td><td>" & PriceCount & "</td></tr>"
        Print #1, Body
        
        Print #1, "</table>"
    Close #1
    
    order_rec.Close
End Sub

Sub CustromYearReport(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer
    Dim CurrentCount As Integer, CurrentPrice As Double, WinningCount As Integer, WinningPrice As Double
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    Dim CData() As String
    Dim PData() As String
    Dim beginv As Integer
    Dim endv As Integer
    Dim OrderDate As String, ProductID As String, ProductName As String
    Dim PriceCount As Double
    
    
    'search order
    CData = Split(cmbCName.Text, " ")
    SQL = "select * from [order] where CID='" & CData(0) & "' and (CurrentDate>='" & Format(txtCurrentDate.Text, "yyyy/") & "01/01' and CurrentDate<='" & Format(txtCurrentDate.Text, "yyyy/") & "12/31') order by CurrentDate;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, order_rec)
  
    
    Open TargetPath For Output As #1
        Print #1, "<table>"
    
        'show report datetime
        Body = "<tr><td>日期</td><td colspan=10>" & Format(txtCurrentDate.Text, "yyyy") & "</td></tr>"
        Print #1, Body
        
        'show custom name and product name
        Body = "<tr><td>客戶名稱</td><td colspan=10>" & CData(1) & "</td></tr>"
        Print #1, Body
        Body = "<tr><td>交易流水號</td><td>交易日期</td><td>產品名稱</td><td>交易數量</td><td>交易金額</td><td>中獎數量</td><td>中獎金額</td><td>漲價</td><td>退水金額</td><td>備註</td></tr>"
        Print #1, Body
        
        'show every order with current date
        PriceCount = 0
        'order_rec.MoveFirst
        Do Until order_rec.EOF
            'search price
            OrderDate = order_rec.Fields.Item("CurrentDate")
            ProductID = order_rec.Fields.Item("PID")
            selectFields = "CurrentPrice,WinningPrice,Upset"
            SQL = "select * from price where PID='" & ProductID & "' and CID='" & CData(0) & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
            Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
            
            'search product
            SQL = "select * from product where PID='" & ProductID & "';"
            Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
            ProductName = product_rec.Fields.Item("PName")
        
            'mark custom name
            Body = "<tr>"
            Body = Body & "<td>" & order_rec.Fields.Item("SwiftCode") & "</td>"
            Body = Body & "<td>" & OrderDate & "</td>"
            Body = Body & "<td>" & ProductName & "</td>"
            
            Body = Body & "<td>" & order_rec.Fields.Item("CurrentCount") & "</td>"
            If price_rec.EOF Then
                Body = Body & "<td>0</td>"
            Else
                Body = Body & "<td>" & order_rec.Fields.Item("CurrentCount") * Val(price_rec.Fields.Item("CurrentPrice")) & "</td>"
                PriceCount = PriceCount + (order_rec.Fields.Item("CurrentCount") * Val(price_rec.Fields.Item("CurrentPrice")))
            End If
            
            Body = Body & "<td>" & order_rec.Fields.Item("WinningCount") & "</td>"
            If price_rec.EOF Then
                Body = Body & "<td>0</td>"
            Else
                Body = Body & "<td>" & order_rec.Fields.Item("WinningCount") * Val(price_rec.Fields.Item("WinningPrice")) & "</td>"
                PriceCount = PriceCount - (order_rec.Fields.Item("WinningCount") * Val(price_rec.Fields.Item("WinningPrice")))
            End If
            
            Body = Body & "<td>" & order_rec.Fields.Item("AddMoney") & "</td>"
            If Not IsNull(order_rec.Fields.Item("AddMoney")) Then PriceCount = PriceCount + Val(order_rec.Fields.Item("AddMoney"))
            
            Body = Body & "<td>" & order_rec.Fields.Item("BonusMoney") & "</td>"
            '不計退水金額 'If Not IsNull(order_rec.Fields.Item("BonusMoney")) Then PriceCount = PriceCount - order_rec.Fields.Item("BonusMoney")
            
            Body = Body & "<td>" & order_rec.Fields.Item("Note") & "</td>"
              

            Body = Body & "</tr>"
            Print #1, Body
            
            order_rec.MoveNext
        Loop
        
        'show price count
        Body = "<tr><td>應收</td><td>" & PriceCount & "</td></tr>"
        Print #1, Body
        
        Print #1, "</table>"
    Close #1
    
    order_rec.Close
End Sub

Sub CustromWeekTransaction(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer
    Dim CurrentCount As Integer, CurrentPrice As Double, WinningCount As Integer, WinningPrice As Double
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    Dim CData() As String
    Dim PData() As String
    Dim beginv As Integer
    Dim endv As Integer
    Dim OrderDate As String, ProductID As String, ProductName As String
    Dim PriceCount As Double
    Dim OldOrderDate As String, CurrentPriceCount As Double, WinningPriceCount As Double, AddMoney As Double, BonusMoney As Double
    Dim DayDiff As Integer
       
       
    'search order
    CData = Split(cmbCName.Text, " ")
    SQL = "select * from [order] where CID='" & CData(0) & "' and (CurrentDate>='" & Format(DateTime.DateAdd("d", -7, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd") & "' and CurrentDate<='" & txtCurrentDate.Text & "') order by CurrentDate;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, order_rec)
  
    
    Open TargetPath For Output As #1
        Print #1, "<table>"
    
        'show report datetime
        Body = "<tr><td>日期</td><td colspan=10>" & DateTime.DateAdd("d", -7, txtCurrentDate.Text) & "至" & txtCurrentDate.Text & "</td></tr>"
        Print #1, Body
        
        'show custom name and product name
        Body = "<tr><td>客戶名稱</td><td colspan=10>" & CData(1) & "</td></tr>"
        Print #1, Body
        Body = "<tr><td>交易日期</td><td>交易金額</td><td>中獎金額</td><td>漲價</td><td>退水金額</td><td>小計</td></tr>"
        Print #1, Body
        
        'show every order with current date
        PriceCount = 0
        CurrentPriceCount = 0
        WinningPriceCount = 0
        AddMoney = 0
        BonusMoney = 0
        'order_rec.MoveFirst
        Do Until order_rec.EOF
            'search price
            OrderDate = order_rec.Fields.Item("CurrentDate")
            ProductID = order_rec.Fields.Item("PID")
            selectFields = "CurrentPrice,WinningPrice,Upset"
            SQL = "select * from price where PID='" & ProductID & "' and CID='" & CData(0) & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
            Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
            
            
            If OldOrderDate = OrderDate Or OldOrderDate = "" Then
                'fill data as first
                If OldOrderDate = "" And OrderDate <> "" Then
                    DayDiff = DateDiff("d", Format(DateTime.DateAdd("d", -7, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd"), OrderDate)
                    For i = 0 To DayDiff - 1
                        Body = "<tr>"
                        Body = Body & "<td>" & DateTime.DateAdd("d", i, Format(DateTime.DateAdd("d", -7, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd")) & "</td>"
                        Body = Body & "<td>0</td>"
                        Body = Body & "<td>0</td>"
                        Body = Body & "<td>0</td>"
                        Body = Body & "<td>0</td>"
                        Body = Body & "<td>0</td>"
                        Body = Body & "</tr>"
                        Print #1, Body
                    Next
                End If
                
                'save data
                If Not price_rec.EOF Then
                    CurrentPriceCount = CurrentPriceCount + (order_rec.Fields.Item("CurrentCount") * Val(price_rec.Fields.Item("CurrentPrice")))
                    WinningPriceCount = WinningPriceCount + (order_rec.Fields.Item("WinningCount") * Val(price_rec.Fields.Item("WinningPrice")))
                End If
                If Not IsNull(order_rec.Fields.Item("AddMoney")) Then AddMoney = AddMoney + Val(order_rec.Fields.Item("AddMoney"))
                If Not IsNull(order_rec.Fields.Item("BonusMoney")) Then BonusMoney = BonusMoney + Val(order_rec.Fields.Item("BonusMoney"))
            Else
                'mark custom name
                Body = "<tr>"
                Body = Body & "<td>" & OldOrderDate & "</td>"
                Body = Body & "<td>" & CurrentPriceCount & "</td>"
                Body = Body & "<td>" & WinningPriceCount & "</td>"
                Body = Body & "<td>" & AddMoney & "</td>"
                Body = Body & "<td>" & BonusMoney & "</td>"
                Body = Body & "<td>" & CurrentPriceCount + WinningPriceCount + AddMoney + BonusMoney & "</td>"
                PriceCount = PriceCount + CurrentPriceCount + WinningPriceCount + AddMoney + BonusMoney
                Body = Body & "</tr>"
                Print #1, Body
                
                'clear variable data
                CurrentPriceCount = 0
                WinningPriceCount = 0
                AddMoney = 0
                BonusMoney = 0
                
                'fill data per day
                DayDiff = DateDiff("d", OldOrderDate, OrderDate)
                For i = 1 To DayDiff - 1
                    Body = "<tr>"
                    Body = Body & "<td>" & DateTime.DateAdd("d", i, OldOrderDate) & "</td>"
                    Body = Body & "<td>0</td>"
                    Body = Body & "<td>0</td>"
                    Body = Body & "<td>0</td>"
                    Body = Body & "<td>0</td>"
                    Body = Body & "<td>0</td>"
                    Body = Body & "</tr>"
                    Print #1, Body
                Next
            End If
            OldOrderDate = OrderDate
            
            
            order_rec.MoveNext
        Loop
        
        
        'show the last data and fill data to lastday as last
        'if everything is empty, then setting a startup date for full data
        If OrderDate = "" Then OrderDate = Format(DateTime.DateAdd("d", -7, Format(txtCurrentDate.Text, "yyyy/MM/dd")))
        If OldOrderDate = "" Then OldOrderDate = OrderDate
        
        'mark custom name
        Body = "<tr>"
        Body = Body & "<td>" & OldOrderDate & "</td>"
        Body = Body & "<td>" & CurrentPriceCount & "</td>"
        Body = Body & "<td>" & WinningPriceCount & "</td>"
        Body = Body & "<td>" & AddMoney & "</td>"
        Body = Body & "<td>" & BonusMoney & "</td>"
        Body = Body & "<td>" & CurrentPriceCount + WinningPriceCount + AddMoney + BonusMoney & "</td>"
        PriceCount = PriceCount + CurrentPriceCount + WinningPriceCount + AddMoney + BonusMoney
        Body = Body & "</tr>"
        Print #1, Body
        DayDiff = DateDiff("d", OrderDate, txtCurrentDate.Text)
        For i = 1 To DayDiff - 1
            Body = "<tr>"
            Body = Body & "<td>" & DateTime.DateAdd("d", i, OrderDate) & "</td>"
            Body = Body & "<td>0</td>"
            Body = Body & "<td>0</td>"
            Body = Body & "<td>0</td>"
            Body = Body & "<td>0</td>"
            Body = Body & "<td>0</td>"
            Body = Body & "</tr>"
            Print #1, Body
        Next
        
        
        'show price count
        Body = "<tr><td>總計</td><td>" & PriceCount & "</td></tr>"
        Print #1, Body
        
        Print #1, "</table>"
    Close #1
    
    order_rec.Close
End Sub


Sub CustromMonthTransaction(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer
    Dim CurrentCount As Integer, CurrentPrice As Double, WinningCount As Integer, WinningPrice As Double
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    Dim CData() As String
    Dim PData() As String
    Dim beginv As Integer
    Dim endv As Integer
    Dim OrderDate As String, ProductID As String, ProductName As String
    Dim PriceCount As Double
    Dim OldOrderDate As String, CurrentPriceCount As Double, WinningPriceCount As Double, AddMoney As Double, BonusMoney As Double
    Dim DayDiff As Integer
       
       
    'search order
    CData = Split(cmbCName.Text, " ")
    SQL = "select * from [order] where CID='" & CData(0) & "' and (CurrentDate>='" & Format(txtCurrentDate.Text, "yyyy/MM/") & "01' and CurrentDate<='" & Format(txtCurrentDate.Text, "yyyy/MM/") & Date_Is_28_29_30_31(txtCurrentDate.Text) & "') order by CurrentDate;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, order_rec)
  
    
    Open TargetPath For Output As #1
        Print #1, "<table>"
    
        'show report datetime
        Body = "<tr><td>日期</td><td colspan=10>" & Format(txtCurrentDate.Text, "yyyy/MM") & "</td></tr>"
        Print #1, Body
        
        'show custom name and product name
        Body = "<tr><td>客戶名稱</td><td colspan=10>" & CData(1) & "</td></tr>"
        Print #1, Body
        Body = "<tr><td>交易日期</td><td>交易金額</td><td>中獎金額</td><td>漲價</td><td>退水金額</td><td>小計</td></tr>"
        Print #1, Body
        
        'show every order with current date
        PriceCount = 0
        CurrentPriceCount = 0
        WinningPriceCount = 0
        AddMoney = 0
        BonusMoney = 0
        'order_rec.MoveFirst
        Do Until order_rec.EOF
            'search price
            OrderDate = order_rec.Fields.Item("CurrentDate")
            ProductID = order_rec.Fields.Item("PID")
            selectFields = "CurrentPrice,WinningPrice,Upset"
            SQL = "select * from price where PID='" & ProductID & "' and CID='" & CData(0) & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
            Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
            
            
            If OldOrderDate = OrderDate Or OldOrderDate = "" Then
                'fill data as first
                If OldOrderDate = "" And OrderDate <> "" Then
                    DayDiff = DateDiff("d", Format(txtCurrentDate.Text, "yyyy/MM/") & "01", OrderDate)
                    For i = 0 To DayDiff - 1
                        Body = "<tr>"
                        Body = Body & "<td>" & DateTime.DateAdd("d", i, Format(txtCurrentDate.Text, "yyyy/MM/") & "01") & "</td>"
                        Body = Body & "<td>0</td>"
                        Body = Body & "<td>0</td>"
                        Body = Body & "<td>0</td>"
                        Body = Body & "<td>0</td>"
                        Body = Body & "<td>0</td>"
                        Body = Body & "</tr>"
                        Print #1, Body
                    Next
                End If
                
                'save data
                If Not price_rec.EOF Then
                    CurrentPriceCount = CurrentPriceCount + (order_rec.Fields.Item("CurrentCount") * Val(price_rec.Fields.Item("CurrentPrice")))
                    WinningPriceCount = WinningPriceCount + (order_rec.Fields.Item("WinningCount") * Val(price_rec.Fields.Item("WinningPrice")))
                End If
                If Not IsNull(order_rec.Fields.Item("AddMoney")) Then AddMoney = AddMoney + Val(order_rec.Fields.Item("AddMoney"))
                If Not IsNull(order_rec.Fields.Item("BonusMoney")) Then BonusMoney = BonusMoney + Val(order_rec.Fields.Item("BonusMoney"))
            Else
                'mark custom name
                Body = "<tr>"
                Body = Body & "<td>" & OldOrderDate & "</td>"
                Body = Body & "<td>" & CurrentPriceCount & "</td>"
                Body = Body & "<td>" & WinningPriceCount & "</td>"
                Body = Body & "<td>" & AddMoney & "</td>"
                Body = Body & "<td>" & BonusMoney & "</td>"
                Body = Body & "<td>" & CurrentPriceCount + WinningPriceCount + AddMoney + BonusMoney & "</td>"
                PriceCount = PriceCount + CurrentPriceCount + WinningPriceCount + AddMoney + BonusMoney
                Body = Body & "</tr>"
                Print #1, Body
                
                'clear variable data
                CurrentPriceCount = 0
                WinningPriceCount = 0
                AddMoney = 0
                BonusMoney = 0
                
                'fill data per day
                DayDiff = DateDiff("d", OldOrderDate, OrderDate)
                For i = 1 To DayDiff - 1
                    Body = "<tr>"
                    Body = Body & "<td>" & DateTime.DateAdd("d", i, OldOrderDate) & "</td>"
                    Body = Body & "<td>0</td>"
                    Body = Body & "<td>0</td>"
                    Body = Body & "<td>0</td>"
                    Body = Body & "<td>0</td>"
                    Body = Body & "<td>0</td>"
                    Body = Body & "</tr>"
                    Print #1, Body
                Next
            End If
            OldOrderDate = OrderDate
            
            
            order_rec.MoveNext
        Loop
        
        
        'show the last data and fill data to lastday as last
        'if everything is empty, then setting a startup date for full data
        If OrderDate = "" Then OrderDate = Format(txtCurrentDate.Text, "yyyy/MM/") & "01"
        If OldOrderDate = "" Then OldOrderDate = OrderDate
        
        'mark custom name
        Body = "<tr>"
        Body = Body & "<td>" & OldOrderDate & "</td>"
        Body = Body & "<td>" & CurrentPriceCount & "</td>"
        Body = Body & "<td>" & WinningPriceCount & "</td>"
        Body = Body & "<td>" & AddMoney & "</td>"
        Body = Body & "<td>" & BonusMoney & "</td>"
        Body = Body & "<td>" & CurrentPriceCount + WinningPriceCount + AddMoney + BonusMoney & "</td>"
        PriceCount = PriceCount + CurrentPriceCount + WinningPriceCount + AddMoney + BonusMoney
        Body = Body & "</tr>"
        Print #1, Body
        DayDiff = DateDiff("d", OrderDate, Format(txtCurrentDate.Text, "yyyy/MM/") & Date_Is_28_29_30_31(txtCurrentDate.Text))
        For i = 1 To DayDiff
            Body = "<tr>"
            Body = Body & "<td>" & DateTime.DateAdd("d", i, OrderDate) & "</td>"
            Body = Body & "<td>0</td>"
            Body = Body & "<td>0</td>"
            Body = Body & "<td>0</td>"
            Body = Body & "<td>0</td>"
            Body = Body & "<td>0</td>"
            Body = Body & "</tr>"
            Print #1, Body
        Next
        
        
        'show price count
        Body = "<tr><td>總計</td><td>" & PriceCount & "</td></tr>"
        Print #1, Body
        
        Print #1, "</table>"
    Close #1
    
    order_rec.Close
End Sub


Sub CustromYearTransaction(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer
    Dim CurrentCount As Integer, CurrentPrice As Double, WinningCount As Integer, WinningPrice As Double
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    Dim CData() As String
    Dim PData() As String
    Dim beginv As Integer
    Dim endv As Integer
    Dim OrderDate As String, ProductID As String, ProductName As String
    Dim PriceCount As Double
    Dim OldOrderDate As String, CurrentPriceCount As Double, WinningPriceCount As Double, AddMoney As Double, BonusMoney As Double
    Dim DayDiff As Integer
       
       
    'search order
    CData = Split(cmbCName.Text, " ")
    SQL = "select * from [order] where CID='" & CData(0) & "' and (CurrentDate>='" & Format(txtCurrentDate.Text, "yyyy/") & "01/01' and CurrentDate<='" & Format(txtCurrentDate.Text, "yyyy/") & "12/31') order by CurrentDate;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, order_rec)
  
    
    Open TargetPath For Output As #1
        Print #1, "<table>"
    
        'show report datetime
        Body = "<tr><td>日期</td><td colspan=10>" & Format(txtCurrentDate.Text, "yyyy") & "</td></tr>"
        Print #1, Body
        
        'show custom name and product name
        Body = "<tr><td>客戶名稱</td><td colspan=10>" & CData(1) & "</td></tr>"
        Print #1, Body
        Body = "<tr><td>交易日期</td><td>交易金額</td><td>中獎金額</td><td>漲價</td><td>退水金額</td><td>小計</td></tr>"
        Print #1, Body
        
        'show every order with current date
        PriceCount = 0
        CurrentPriceCount = 0
        WinningPriceCount = 0
        AddMoney = 0
        BonusMoney = 0
        'order_rec.MoveFirst
        Do Until order_rec.EOF
            'search price
            OrderDate = order_rec.Fields.Item("CurrentDate")
            ProductID = order_rec.Fields.Item("PID")
            selectFields = "CurrentPrice,WinningPrice,Upset"
            SQL = "select * from price where PID='" & ProductID & "' and CID='" & CData(0) & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
            Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
            
            
            If OldOrderDate = OrderDate Or OldOrderDate = "" Then
                'fill data as first
                If OldOrderDate = "" And OrderDate <> "" Then
                    DayDiff = DateDiff("d", Format(txtCurrentDate.Text, "yyyy/") & "01/01", OrderDate)
                    For i = 0 To DayDiff - 1
                        Body = "<tr>"
                        Body = Body & "<td>" & DateTime.DateAdd("d", i, Format(txtCurrentDate.Text, "yyyy/") & "01/01") & "</td>"
                        Body = Body & "<td>0</td>"
                        Body = Body & "<td>0</td>"
                        Body = Body & "<td>0</td>"
                        Body = Body & "<td>0</td>"
                        Body = Body & "<td>0</td>"
                        Body = Body & "</tr>"
                        Print #1, Body
                    Next
                End If
                
                'save data
                If Not price_rec.EOF Then
                    CurrentPriceCount = CurrentPriceCount + (order_rec.Fields.Item("CurrentCount") * Val(price_rec.Fields.Item("CurrentPrice")))
                    WinningPriceCount = WinningPriceCount + (order_rec.Fields.Item("WinningCount") * Val(price_rec.Fields.Item("WinningPrice")))
                End If
                If Not IsNull(order_rec.Fields.Item("AddMoney")) Then AddMoney = AddMoney + Val(order_rec.Fields.Item("AddMoney"))
                If Not IsNull(order_rec.Fields.Item("BonusMoney")) Then BonusMoney = BonusMoney + Val(order_rec.Fields.Item("BonusMoney"))
            Else
                'mark custom name
                Body = "<tr>"
                Body = Body & "<td>" & OldOrderDate & "</td>"
                Body = Body & "<td>" & CurrentPriceCount & "</td>"
                Body = Body & "<td>" & WinningPriceCount & "</td>"
                Body = Body & "<td>" & AddMoney & "</td>"
                Body = Body & "<td>" & BonusMoney & "</td>"
                Body = Body & "<td>" & CurrentPriceCount + WinningPriceCount + AddMoney + BonusMoney & "</td>"
                PriceCount = PriceCount + CurrentPriceCount + WinningPriceCount + AddMoney + BonusMoney
                Body = Body & "</tr>"
                Print #1, Body
                
                'clear variable data
                CurrentPriceCount = 0
                WinningPriceCount = 0
                AddMoney = 0
                BonusMoney = 0
                
                'fill data per day
                DayDiff = DateDiff("d", OldOrderDate, OrderDate)
                For i = 1 To DayDiff - 1
                    Body = "<tr>"
                    Body = Body & "<td>" & DateTime.DateAdd("d", i, OldOrderDate) & "</td>"
                    Body = Body & "<td>0</td>"
                    Body = Body & "<td>0</td>"
                    Body = Body & "<td>0</td>"
                    Body = Body & "<td>0</td>"
                    Body = Body & "<td>0</td>"
                    Body = Body & "</tr>"
                    Print #1, Body
                Next
            End If
            OldOrderDate = OrderDate
            
            
            order_rec.MoveNext
        Loop
        
        
        'show the last data and fill data to lastday as last
        'if everything is empty, then setting a startup date for full data
        If OrderDate = "" Then OrderDate = Format(txtCurrentDate.Text, "yyyy/") & "01/01"
        If OldOrderDate = "" Then OldOrderDate = OrderDate
        
        'mark custom name
        Body = "<tr>"
        Body = Body & "<td>" & OldOrderDate & "</td>"
        Body = Body & "<td>" & CurrentPriceCount & "</td>"
        Body = Body & "<td>" & WinningPriceCount & "</td>"
        Body = Body & "<td>" & AddMoney & "</td>"
        Body = Body & "<td>" & BonusMoney & "</td>"
        Body = Body & "<td>" & CurrentPriceCount + WinningPriceCount + AddMoney + BonusMoney & "</td>"
        PriceCount = PriceCount + CurrentPriceCount + WinningPriceCount + AddMoney + BonusMoney
        Body = Body & "</tr>"
        Print #1, Body
        DayDiff = DateDiff("d", OrderDate, Format(txtCurrentDate.Text, "yyyy/") & "12/31")
        For i = 1 To DayDiff
            Body = "<tr>"
            Body = Body & "<td>" & DateTime.DateAdd("d", i, OrderDate) & "</td>"
            Body = Body & "<td>0</td>"
            Body = Body & "<td>0</td>"
            Body = Body & "<td>0</td>"
            Body = Body & "<td>0</td>"
            Body = Body & "<td>0</td>"
            Body = Body & "</tr>"
            Print #1, Body
        Next
        
        
        'show price count
        Body = "<tr><td>總計</td><td>" & PriceCount & "</td></tr>"
        Print #1, Body
        
        Print #1, "</table>"
    Close #1
    
    order_rec.Close
End Sub

Private Sub cmdConfirm_Click()
'On Error GoTo errout:
    If txtCurrentDate.Text = "" Then
        MsgBox "請先選擇要列印的時間！"
    ElseIf (basVariable.Parameter = "CustromProductDayReport" Or basVariable.Parameter = "CustromProductWeekReport") And cmbCName.Text = "" And cmbPName.Text = "" Then
        MsgBox "尚未選擇客戶或產品！"
    ElseIf (basVariable.Parameter = "CustromWeekReport" Or basVariable.Parameter = "CustromMonthReport" Or basVariable.Parameter = "CustromYearReport") And cmbCName.Text = "" Then
        MsgBox "尚未選擇客戶！"
    Else
        Dim TargetPath As String
        Dim CData() As String
        Dim PData() As String
        
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
        Case "CustromProductDayReport"
            CData = Split(cmbCName.Text, " ")
            PData = Split(cmbPName.Text, " ")
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyyMMdd") & "_" & CData(1) & "_" & PData(1) & "_客別分產品日報表.xls"
            Call CustromProductDayReport(TargetPath)
        Case "CustromProductWeekReport"
            CData = Split(cmbCName.Text, " ")
            PData = Split(cmbPName.Text, " ")
            TargetPath = TargetPath & Format(DateTime.DateAdd("d", -7, txtCurrentDate.Text), "yyyyMMdd") & "至" & Format(txtCurrentDate.Text, "yyyyMMdd") & "_" & CData(1) & "_" & PData(1) & "_客別分產品週報表.xls"
            Call CustromProductWeekReport(TargetPath)
        Case "CustromProductWeekTransaction"
            CData = Split(cmbCName.Text, " ")
            PData = Split(cmbPName.Text, " ")
            TargetPath = TargetPath & Format(DateTime.DateAdd("d", -7, txtCurrentDate.Text), "yyyyMMdd") & "至" & Format(txtCurrentDate.Text, "yyyyMMdd") & "_" & CData(1) & "_" & PData(1) & "_客別分產品週交易金額表.xls"
            Call CustromProductWeekTransaction(TargetPath)
        Case "CustromWeekReport"
            CData = Split(cmbCName.Text, " ")
            PData = Split(cmbPName.Text, " ")
            TargetPath = TargetPath & Format(DateTime.DateAdd("d", -7, txtCurrentDate.Text), "yyyyMMdd") & "至" & Format(txtCurrentDate.Text, "yyyyMMdd") & "_" & CData(1) & "_客別不分產品週報表.xls"
            Call CustromWeekReport(TargetPath)
        Case "CustromMonthReport"
            CData = Split(cmbCName.Text, " ")
            PData = Split(cmbPName.Text, " ")
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyyMM") & "_" & CData(1) & "_客別不分產品月報表.xls"
            Call CustromMonthReport(TargetPath)
        Case "CustromYearReport"
            CData = Split(cmbCName.Text, " ")
            PData = Split(cmbPName.Text, " ")
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyy") & "_" & CData(1) & "_客別不分產品年報表.xls"
            Call CustromYearReport(TargetPath)
        Case "CustromWeekTransaction"
            CData = Split(cmbCName.Text, " ")
            PData = Split(cmbPName.Text, " ")
            TargetPath = TargetPath & Format(DateTime.DateAdd("d", -7, txtCurrentDate.Text), "yyyyMMdd") & "至" & Format(txtCurrentDate.Text, "yyyyMMdd") & "_" & CData(1) & "_客別不分產品週交易金額表.xls"
            Call CustromWeekTransaction(TargetPath)
        Case "CustromMonthTransaction"
            CData = Split(cmbCName.Text, " ")
            PData = Split(cmbPName.Text, " ")
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyyMM") & "_" & CData(1) & "_客別不分產品月交易金額表.xls"
            Call CustromMonthTransaction(TargetPath)
        Case "CustromYearTransaction"
            CData = Split(cmbCName.Text, " ")
            PData = Split(cmbPName.Text, " ")
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyy") & "_" & CData(1) & "_客別不分產品年交易金額表.xls"
            Call CustromYearTransaction(TargetPath)
        End Select
        
        MsgBox "已輸出報表至" & TargetPath & "！"
    End If

    If False Then
errout:
        MsgBox "資料檔有誤。或無法寫入，因為舊報表未關閉！"
    End If
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
    Case "CustromProductDayReport"
        Label1(0).Caption = "客別分產品日報表列印"
        lblEntry(0).Visible = True
        lblEntry(2).Visible = True
        cmbCName.Visible = True
        cmbPName.Visible = True
        Call ComboBox_LoadFrom_DataBase_ByFile(cmbCName, "CID,CName", "custom", "", "", "")
        Call ComboBox_LoadFrom_DataBase_ByFile(cmbPName, "PID,PName", "product", "", "", "")
        
        Call cmbPName.AddItem("100 539_全")
        Call cmbPName.AddItem("110 港號_全")
        Call cmbPName.AddItem("120 大樂透_全")
    Case "CustromProductWeekReport"
        Label1(0).Caption = "客別分產品週報表列印"
        lblEntry(0).Visible = True
        lblEntry(2).Visible = True
        cmbCName.Visible = True
        cmbPName.Visible = True
        Call ComboBox_LoadFrom_DataBase_ByFile(cmbCName, "CID,CName", "custom", "", "", "")
        Call ComboBox_LoadFrom_DataBase_ByFile(cmbPName, "PID,PName", "product", "", "", "")
        
        Call cmbPName.AddItem("100 539_全")
        Call cmbPName.AddItem("110 港號_全")
        Call cmbPName.AddItem("120 大樂透_全")
    Case "CustromProductWeekTransaction"
        Label1(0).Caption = "客別分產品週交易金額表列印"
        lblEntry(0).Visible = True
        lblEntry(2).Visible = True
        cmbCName.Visible = True
        cmbPName.Visible = True
        Call ComboBox_LoadFrom_DataBase_ByFile(cmbCName, "CID,CName", "custom", "", "", "")
        Call ComboBox_LoadFrom_DataBase_ByFile(cmbPName, "PID,PName", "product", "", "", "")
        
        Call cmbPName.AddItem("100 539_全")
        Call cmbPName.AddItem("110 港號_全")
        Call cmbPName.AddItem("120 大樂透_全")
    Case "CustromWeekReport"
        Label1(0).Caption = "客別不分產品週報表列印"
        lblEntry(0).Visible = True
        cmbCName.Visible = True
        Call ComboBox_LoadFrom_DataBase_ByFile(cmbCName, "CID,CName", "custom", "", "", "")
    Case "CustromMonthReport"
        Label1(0).Caption = "客別不分產品月報表列印"
        lblEntry(0).Visible = True
        cmbCName.Visible = True
        Call ComboBox_LoadFrom_DataBase_ByFile(cmbCName, "CID,CName", "custom", "", "", "")
    Case "CustromYearReport"
        Label1(0).Caption = "客別不分產品年報表列印"
        lblEntry(0).Visible = True
        cmbCName.Visible = True
        Call ComboBox_LoadFrom_DataBase_ByFile(cmbCName, "CID,CName", "custom", "", "", "")
    Case "CustromWeekTransaction"
        Label1(0).Caption = "客別不分產品週交易金額表列印"
        lblEntry(0).Visible = True
        cmbCName.Visible = True
        Call ComboBox_LoadFrom_DataBase_ByFile(cmbCName, "CID,CName", "custom", "", "", "")
    Case "CustromMonthTransaction"
        Label1(0).Caption = "客別不分產品月交易金額表列印"
        lblEntry(0).Visible = True
        cmbCName.Visible = True
        Call ComboBox_LoadFrom_DataBase_ByFile(cmbCName, "CID,CName", "custom", "", "", "")
    Case "CustromYearTransaction"
        Label1(0).Caption = "客別不分產品年交易金額表列印"
        lblEntry(0).Visible = True
        cmbCName.Visible = True
        Call ComboBox_LoadFrom_DataBase_ByFile(cmbCName, "CID,CName", "custom", "", "", "")
    End Select
    
    
    dtpCurrentDate.Value = Format(DateTime.Now, "yyyy/MM/dd")
    txtCurrentDate.Text = Format(DateTime.Now, "yyyy/MM/dd")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmProve.Show
    Unload Me
End Sub
