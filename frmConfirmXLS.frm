VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmConfirmXLS 
   BorderStyle     =   1  '單線固定
   Caption         =   "Report Confirm"
   ClientHeight    =   3990
   ClientLeft      =   6315
   ClientTop       =   7785
   ClientWidth     =   4680
   Icon            =   "frmConfirmXLS.frx":0000
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
      Format          =   37421059
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
Attribute VB_Name = "frmConfirmXLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Call Form_Unload(0)
End Sub

Sub DayReport(ByVal TargetPath As String)
    Dim exap As Excel.Application
    Dim exwb As Excel.Workbook
    Dim exsh As Excel.Worksheet
    Dim row As Integer, col As Integer
    Set exap = New Excel.Application
    Set exwb = exap.Workbooks.Add
    Set exsh = exwb.Sheets.Item(1)
    exsh.Columns.ColumnWidth = 16

    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String
    Dim CurrentCount As Double, CurrentPrice As Long, WinningCount As Double, WinningPrice As Long
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset, rec2 As New adoDB.Recordset
    Dim BuyCurrentCount As Double, SellCurrentCount As Double, BuyCurrentPrice As Long, SellCurrentPrice As Long
    Dim BuyWinningCount As Double, SellWinningCount As Double, BuyWinningPrice As Long, SellWinningPrice As Long
    Dim OrderDate As String
    Dim Proportion As Integer, BonusTarget As String, BonusMoney As Long, BonusProduct As String
    Dim CID As String, CName As String
    Dim PriceCount As Long
    Dim PIDBuyCurrentCount(1024) As Double, PIDSellCurrentCount(1024) As Double, PIDBuyCurrentPrice(1024) As Long, PIDSellCurrentPrice(1024) As Long
    Dim PIDBuyWinningCount(1024) As Double, PIDSellWinningCount(1024) As Double, PIDBuyWinningPrice(1024) As Long, PIDSellWinningPrice(1024) As Long
    
    
    BuyCurrentCount = 0
    BuyWinningCount = 0
    BuyCurrentPrice = 0
    BuyWinningPrice = 0
    SellCurrentCount = 0
    SellWinningCount = 0
    SellCurrentPrice = 0
    SellWinningPrice = 0
    'init variable to account buy and sell for each product
    For i = 0 To Count - 1
        PIDBuyCurrentCount(i) = 0
        PIDBuyCurrentPrice(i) = 0
        PIDBuyWinningCount(i) = 0
        PIDBuyWinningPrice(i) = 0
    Next
    
    
    SQL = "select * from product order by CLng(PID);"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom order by CLng(CID);"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)

    

        'show report datetime
        row = 1: col = 1: exsh.Cells(row, col) = "日期"
        col = col + 1: exsh.Cells(row, col) = txtCurrentDate.Text
        col = 11:   Call exsh.Range(Numeric2CharEN(2) & row, Numeric2CharEN(col) & row).Merge
        row = row + 1
        
        'show product name
        col = 1:  exsh.Cells(row, col) = "產品"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName") & "交易數量"
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName") & "中獎數量"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        col = col + 1: exsh.Cells(row, col) = "客戶金額總計"
        row = row + 1
        
        
        'show every custom order per product
        product_rec.MoveFirst
        Do Until custom_rec.EOF
            'mark custom name
            CID = custom_rec.Fields.Item("CID")
            CName = custom_rec.Fields.Item("CName")
            col = 1:  exsh.Cells(row, col) = CName


            'show every custom order per product
            PriceCount = 0
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
                
                    CurrentCount = CurrentCount + SimpleRound(Val(rec1.Fields.Item("CurrentCount")), 4)
                    WinningCount = WinningCount + SimpleRound(Val(rec1.Fields.Item("WinningCount")), 4)
                    If Not price_rec.EOF Then
                        CurrentPrice = CurrentPrice + SimpleRound(Val(rec1.Fields.Item("CurrentCount")) * Val(price_rec.Fields.Item("CurrentPrice")), 0)
                        WinningPrice = WinningPrice + SimpleRound(Val(rec1.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")), 0)
                    End If
                    
                    'a new variable to account buy and sell for each product
                    If CurrentCount >= 0 Then
                        PIDBuyCurrentCount(i) = PIDBuyCurrentCount(i) + SimpleRound(rec1.Fields.Item("CurrentCount"), 4)
                        PIDBuyCurrentPrice(i) = PIDBuyCurrentPrice(i) + SimpleRound(Val(rec1.Fields.Item("CurrentCount")) * Val(price_rec.Fields.Item("CurrentPrice")), 0)
                        PIDBuyWinningCount(i) = PIDBuyWinningCount(i) + SimpleRound(rec1.Fields.Item("WinningCount"), 4)
                        PIDBuyWinningPrice(i) = PIDBuyWinningPrice(i) + SimpleRound(Val(rec1.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")), 0)
                    Else
                        PIDSellCurrentCount(i) = PIDSellCurrentCount(i) + SimpleRound(rec1.Fields.Item("CurrentCount"), 4)
                        PIDSellCurrentPrice(i) = PIDSellCurrentPrice(i) + SimpleRound(Val(rec1.Fields.Item("CurrentCount")) * Val(price_rec.Fields.Item("CurrentPrice")), 0)
                        PIDSellWinningCount(i) = PIDSellWinningCount(i) + SimpleRound(rec1.Fields.Item("WinningCount"), 4)
                        PIDSellWinningPrice(i) = PIDSellWinningPrice(i) + SimpleRound(Val(rec1.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")), 0)
                    End If
                    
                    rec1.MoveNext
                Loop
                col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentCount, 4)
                col = col + 1: exsh.Cells(row, col) = SimpleRound(WinningCount, 4)
                PriceCount = PriceCount + SimpleRound(CurrentPrice - WinningPrice, 0)
                
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
            col = col + 1: exsh.Cells(row, col) = SimpleRound(PriceCount, 0)
            row = row + 1
            
            custom_rec.MoveNext
        Loop
        
        
        'show buy and sell for each product account
        row = row + 1
        
        col = 1: exsh.Cells(row, col) = "買數量"
        For i = 0 To Count - 1
            col = col + 1: exsh.Cells(row, col) = SimpleRound(PIDBuyCurrentCount(i), 4)
            col = col + 1: exsh.Cells(row, col) = SimpleRound(PIDBuyWinningCount(i), 4)
        Next
        row = row + 1
        col = 1: exsh.Cells(row, col) = "買金額"
        For i = 0 To Count - 1
            col = col + 1: exsh.Cells(row, col) = SimpleRound(PIDBuyCurrentPrice(i), 0)
            col = col + 1: exsh.Cells(row, col) = SimpleRound(PIDBuyWinningPrice(i), 0)
        Next
        row = row + 1
        
        col = 1: exsh.Cells(row, col) = "賣數量"
        For i = 0 To Count - 1
            col = col + 1: exsh.Cells(row, col) = SimpleRound(PIDSellCurrentCount(i), 4)
            col = col + 1: exsh.Cells(row, col) = SimpleRound(PIDSellWinningCount(i), 4)
        Next
        row = row + 1
        col = 1: exsh.Cells(row, col) = "賣金額"
        For i = 0 To Count - 1
            col = col + 1: exsh.Cells(row, col) = SimpleRound(PIDSellCurrentPrice(i), 0)
            col = col + 1: exsh.Cells(row, col) = SimpleRound(PIDSellWinningPrice(i), 0)
        Next
        row = row + 1
        
        
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計買"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyCurrentCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyCurrentPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計買中"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyWinningCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyWinningPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計賣"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellCurrentCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellCurrentPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計賣中"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellWinningCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellWinningPrice, 0)
        row = row + 1
        
        
    Call exwb.SaveAs(TargetPath)
    Call exap.Quit
    
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Sub WeekReport(ByVal TargetPath As String)
    Dim exap As Excel.Application
    Dim exwb As Excel.Workbook
    Dim exsh As Excel.Worksheet
    Dim row As Integer, col As Integer
    Set exap = New Excel.Application
    Set exwb = exap.Workbooks.Add
    Set exsh = exwb.Sheets.Item(1)
    exsh.Columns.ColumnWidth = 16
    
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String
    Dim CurrentCount As Double, CurrentPrice As Long, WinningCount As Double, WinningPrice As Long
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset, rec2 As New adoDB.Recordset
    Dim BuyCurrentCount As Double, SellCurrentCount As Double, BuyCurrentPrice As Long, SellCurrentPrice As Long
    Dim BuyWinningCount As Double, SellWinningCount As Double, BuyWinningPrice As Long, SellWinningPrice As Long
    Dim OrderDate As String
    Dim Proportion As Integer, BonusTarget As String, BonusMoney As Long, BonusProduct As String
    Dim CID As String, CName As String
    
    
    BuyCurrentCount = 0
    BuyWinningCount = 0
    BuyCurrentPrice = 0
    BuyWinningPrice = 0
    SellCurrentCount = 0
    SellWinningCount = 0
    SellCurrentPrice = 0
    SellWinningPrice = 0
    
    
    SQL = "select * from product order by CLng(PID);"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom order by CLng(CID);"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    
    

    
        'show report datetime
        row = 1: col = 1: exsh.Cells(row, col) = "日期"
        col = col + 1: exsh.Cells(row, col) = DateTime.DateAdd("d", -6, txtCurrentDate.Text) & "至" & txtCurrentDate.Text
        col = 11:   Call exsh.Range(Numeric2CharEN(2) & row, Numeric2CharEN(col) & row).Merge
        row = row + 1
        
        'show product name
        col = 1: exsh.Cells(row, col) = "產品"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName")
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName") & "交易金額"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        col = col + 1: exsh.Cells(row, col) = "成數"
        col = col + 1: exsh.Cells(row, col) = "前帳"
        row = row + 1
        
        'show every custom order per product
        product_rec.MoveFirst
        Do Until custom_rec.EOF
            'mark custom name
            CID = custom_rec.Fields.Item("CID")
            CName = custom_rec.Fields.Item("CName")
            col = 1: exsh.Cells(row, col) = CName

            'show every custom order per product
            For i = 0 To Count - 1
                'query a new custom order
                selectFields = "CurrentCount,WinningCount,CurrentDate"
                SQL = "select " & selectFields & " from [order] where [order].PID='" & PIDArray(i) & "' and CID='" & CID & "' and (CurrentDate>='" & Format(DateTime.DateAdd("d", -6, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd") & "' and CurrentDate<='" & txtCurrentDate.Text & "') order by [order].PID;"
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
                
                    CurrentCount = CurrentCount + SimpleRound(Val(rec1.Fields.Item("CurrentCount")), 4)
                    WinningCount = WinningCount + SimpleRound(Val(rec1.Fields.Item("WinningCount")), 4)
                    If Not price_rec.EOF Then
                        CurrentPrice = CurrentPrice + SimpleRound(Val(rec1.Fields.Item("CurrentCount")) * Val(price_rec.Fields.Item("CurrentPrice")), 0)
                        WinningPrice = WinningPrice + SimpleRound(Val(rec1.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")), 0)
                    End If
                    
                    rec1.MoveNext
                Loop
                col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentCount, 4)
                col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentPrice, 0)
                
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
                SQL = "select * from [order] where CID='" & BonusTarget & "' and (CurrentDate>='" & Format(DateTime.DateAdd("d", -6, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd") & "' and CurrentDate<='" & txtCurrentDate.Text & "');"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec2)
                Do Until rec2.EOF
                    'search price, and addition to variable
                    OrderDate = rec2.Fields.Item("CurrentDate")
                    BonusProduct = rec2.Fields.Item("PID")
                    selectFields = "CurrentPrice,WinningPrice,Upset"
                    SQL = "select * from price where PID='" & BonusProduct & "' and CID='" & BonusTarget & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
                    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
                    
                    If Not price_rec.EOF Then
                        BonusMoney = BonusMoney + SimpleRound(Val(rec2.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")), 0)
                    End If
                    rec2.MoveNext
                Loop
                
                col = 1: exsh.Cells(row, col) = "來自" & rec1.Fields.Item("CName") & "成數" & Proportion & "%共" & SimpleRound(BonusMoney * (Proportion / 100), 0) & "元。"
                rec1.MoveNext
            Loop
            
            row = row + 1
            
            custom_rec.MoveNext
        Loop
        
        
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計買"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyCurrentCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyCurrentPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計買中"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyWinningCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyWinningPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計賣"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellCurrentCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellCurrentPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計賣中"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellWinningCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellWinningPrice, 0)
        row = row + 1
        
        
    Call exwb.SaveAs(TargetPath)
    Call exap.Quit

    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Sub MonthReport(ByVal TargetPath As String)
    Dim exap As Excel.Application
    Dim exwb As Excel.Workbook
    Dim exsh As Excel.Worksheet
    Dim row As Integer, col As Integer
    Set exap = New Excel.Application
    Set exwb = exap.Workbooks.Add
    Set exsh = exwb.Sheets.Item(1)
    exsh.Columns.ColumnWidth = 16
    
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String
    Dim CurrentCount As Double, CurrentPrice As Long, WinningCount As Double, WinningPrice As Long
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset, rec2 As New adoDB.Recordset
    Dim BuyCurrentCount As Double, SellCurrentCount As Double, BuyCurrentPrice As Long, SellCurrentPrice As Long
    Dim BuyWinningCount As Double, SellWinningCount As Double, BuyWinningPrice As Long, SellWinningPrice As Long
    Dim OrderDate As String
    Dim Proportion As Integer, BonusTarget As String, BonusMoney As Long, BonusProduct As String
    Dim CID As String, CName As String
    
    
    BuyCurrentCount = 0
    BuyWinningCount = 0
    BuyCurrentPrice = 0
    BuyWinningPrice = 0
    SellCurrentCount = 0
    SellWinningCount = 0
    SellCurrentPrice = 0
    SellWinningPrice = 0
    
    
    SQL = "select * from product order by CLng(PID);"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom order by CLng(CID);"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    
    
    
        'show report datetime
        row = 1: col = 1: exsh.Cells(row, col) = "日期"
        col = col + 1: exsh.Cells(row, col) = Format(txtCurrentDate.Text, "yyyy/MM") & "月"
        col = 11:   Call exsh.Range(Numeric2CharEN(2) & row, Numeric2CharEN(col) & row).Merge
        row = row + 1
        
        'show product name
        col = 1: exsh.Cells(row, col) = "產品"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName")
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName") & "交易金額"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        col = col + 1: exsh.Cells(row, col) = "成數"
        row = row + 1
        
        'show every custom order per product
        product_rec.MoveFirst
        Do Until custom_rec.EOF
            'mark custom name
            CID = custom_rec.Fields.Item("CID")
            CName = custom_rec.Fields.Item("CName")
            col = 1: exsh.Cells(row, col) = CName


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
                
                    CurrentCount = CurrentCount + SimpleRound(Val(rec1.Fields.Item("CurrentCount")), 4)
                    WinningCount = WinningCount + SimpleRound(Val(rec1.Fields.Item("WinningCount")), 4)
                    If Not price_rec.EOF Then
                        CurrentPrice = CurrentPrice + SimpleRound(Val(rec1.Fields.Item("CurrentCount")) * Val(price_rec.Fields.Item("CurrentPrice")), 0)
                        WinningPrice = WinningPrice + SimpleRound(Val(rec1.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")), 0)
                    End If
                    
                    rec1.MoveNext
                Loop
                col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentCount, 4)
                col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentPrice, 0)
                
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
                        BonusMoney = BonusMoney + SimpleRound(Val(rec2.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")), 0)
                    End If
                    rec2.MoveNext
                Loop
                
                col = col + 1: exsh.Cells(row, col) = "來自" & rec1.Fields.Item("CName") & "成數" & Proportion & "%共" & SimpleRound(BonusMoney * (Proportion / 100), 0) & "元。"
                rec1.MoveNext
            Loop
            
            row = row + 1
            custom_rec.MoveNext
        Loop
        
        
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計買"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyCurrentCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyCurrentPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計買中"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyWinningCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyWinningPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計賣"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellCurrentCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellCurrentPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計賣中"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellWinningCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellWinningPrice, 0)
        row = row + 1
        
        
    Call exwb.SaveAs(TargetPath)
    Call exap.Quit
    
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Sub YearReport(ByVal TargetPath As String)
    Dim exap As Excel.Application
    Dim exwb As Excel.Workbook
    Dim exsh As Excel.Worksheet
    Dim row As Integer, col As Integer
    Set exap = New Excel.Application
    Set exwb = exap.Workbooks.Add
    Set exsh = exwb.Sheets.Item(1)
    exsh.Columns.ColumnWidth = 16
    
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String
    Dim CurrentCount As Double, CurrentPrice As Long, WinningCount As Double, WinningPrice As Long
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset, rec2 As New adoDB.Recordset
    Dim BuyCurrentCount As Double, SellCurrentCount As Double, BuyCurrentPrice As Long, SellCurrentPrice As Long
    Dim BuyWinningCount As Double, SellWinningCount As Double, BuyWinningPrice As Long, SellWinningPrice As Long
    Dim OrderDate As String
    Dim Proportion As Integer, BonusTarget As String, BonusMoney As Long, BonusProduct As String
    Dim CID As String, CName As String
    
    
    BuyCurrentCount = 0
    BuyWinningCount = 0
    BuyCurrentPrice = 0
    BuyWinningPrice = 0
    SellCurrentCount = 0
    SellWinningCount = 0
    SellCurrentPrice = 0
    SellWinningPrice = 0
    
    
    SQL = "select * from product order by CLng(PID);"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom order by CLng(CID);"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    
    
    
        'show report datetime
        row = 1: col = 1: exsh.Cells(row, col) = "日期"
        col = col + 1: exsh.Cells(row, col) = Format(txtCurrentDate.Text, "yyyy") & "年"
        col = 11:   Call exsh.Range(Numeric2CharEN(2) & row, Numeric2CharEN(col) & row).Merge
        row = row + 1
        
        'show product name
        col = 1: exsh.Cells(row, col) = "產品"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName")
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName") & "交易金額"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        col = col + 1: exsh.Cells(row, col) = "成數"
        row = row + 1
        
        'show every custom order per product
        product_rec.MoveFirst
        Do Until custom_rec.EOF
            'mark custom name
            CID = custom_rec.Fields.Item("CID")
            CName = custom_rec.Fields.Item("CName")
            col = 1: exsh.Cells(row, col) = CName


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
                
                    CurrentCount = CurrentCount + SimpleRound(Val(rec1.Fields.Item("CurrentCount")), 4)
                    WinningCount = WinningCount + SimpleRound(Val(rec1.Fields.Item("WinningCount")), 4)
                    If Not price_rec.EOF Then
                        CurrentPrice = CurrentPrice + SimpleRound(Val(rec1.Fields.Item("CurrentCount")) * Val(price_rec.Fields.Item("CurrentPrice")), 0)
                        WinningPrice = WinningPrice + SimpleRound(Val(rec1.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")), 0)
                    End If
                    
                    rec1.MoveNext
                Loop
                col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentCount, 4)
                col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentPrice, 0)
                
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
                        BonusMoney = BonusMoney + SimpleRound(Val(rec2.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")), 0)
                    End If
                    rec2.MoveNext
                Loop
                
                col = 1: exsh.Cells(row, col) = "來自" & rec1.Fields.Item("CName") & "成數" & Proportion & "%共" & SimpleRound(BonusMoney * (Proportion / 100), 0) & "元。"
                rec1.MoveNext
            Loop
            
            row = row + 1
            
            custom_rec.MoveNext
        Loop
        
        
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計買"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyCurrentCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyCurrentPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計買中"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyWinningCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyWinningPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計賣"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellCurrentCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellCurrentPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計賣中"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellWinningCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellWinningPrice, 0)
        row = row + 1
        
        
    Call exwb.SaveAs(TargetPath)
    Call exap.Quit
    
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Sub DayAccount(ByVal TargetPath As String)
    Dim exap As Excel.Application
    Dim exwb As Excel.Workbook
    Dim exsh As Excel.Worksheet
    Dim row As Integer, col As Integer
    Set exap = New Excel.Application
    Set exwb = exap.Workbooks.Add
    Set exsh = exwb.Sheets.Item(1)
    exsh.Columns.ColumnWidth = 16
    
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String
    Dim CurrentCount As Double, CurrentPrice As Long, WinningCount As Double, WinningPrice As Long
    Dim CurrentCountAll(1024) As Double, WinningCountAll(1024) As Double
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset, rec2 As New adoDB.Recordset
    Dim BuyCurrentCount As Double, SellCurrentCount As Double, BuyCurrentPrice As Long, SellCurrentPrice As Long
    Dim BuyWinningCount As Double, SellWinningCount As Double, BuyWinningPrice As Long, SellWinningPrice As Long
    Dim OrderDate As String
    Dim Proportion As Integer, BonusTarget As String, BonusMoney As Long, BonusProduct As String
    Dim CID As String, CName As String
    
    
    BuyCurrentCount = 0
    BuyWinningCount = 0
    BuyCurrentPrice = 0
    BuyWinningPrice = 0
    SellCurrentCount = 0
    SellWinningCount = 0
    SellCurrentPrice = 0
    SellWinningPrice = 0
    
    
    SQL = "select * from product order by CLng(PID);"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom order by CLng(CID);"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    


        'show report datetime
        row = 1: col = 1: exsh.Cells(row, col) = txtCurrentDate.Text
        col = 2:   Call exsh.Range(Numeric2CharEN(1) & row, Numeric2CharEN(col) & row).Merge
        col = col + 1: exsh.Cells(row, col) = "日總計"
        row = row + 1
        
        'show product name
        col = 1:  exsh.Cells(row, col) = "產品"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName")
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName") & "交易金額"
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName") & "中"
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName") & "中獎金額"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        row = row + 1
        
        
        'show every custom order per product
        product_rec.MoveFirst
        Do Until custom_rec.EOF
            'mark custom name
            CID = custom_rec.Fields.Item("CID")
            CName = custom_rec.Fields.Item("CName")
            col = 1: exsh.Cells(row, col) = CName


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
                
                    CurrentCount = CurrentCount + SimpleRound(Val(rec1.Fields.Item("CurrentCount")), 4)
                    WinningCount = WinningCount + SimpleRound(Val(rec1.Fields.Item("WinningCount")), 4)
                    If Not price_rec.EOF Then
                        CurrentPrice = CurrentPrice + SimpleRound(Val(rec1.Fields.Item("CurrentCount")) * Val(price_rec.Fields.Item("CurrentPrice")), 0)
                        WinningPrice = WinningPrice + SimpleRound(Val(rec1.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")), 0)
                    End If
                    
                    rec1.MoveNext
                Loop
                col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentCount, 4)
                col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentPrice, 0)
                col = col + 1: exsh.Cells(row, col) = SimpleRound(WinningCount, 4)
                col = col + 1: exsh.Cells(row, col) = SimpleRound(WinningPrice, 0)
                
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
            
            row = row + 1
            
            custom_rec.MoveNext
        Loop
        
        
        col = 1: exsh.Cells(row, col) = "數量總計"
        For i = 0 To Count - 1
            col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentCountAll(i), 4):   col = col + 1
            col = col + 1: exsh.Cells(row, col) = SimpleRound(WinningCountAll(i), 4):   col = col + 1
        Next
        row = row + 1
        
              
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計買"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyCurrentCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyCurrentPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計買中"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyWinningCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyWinningPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計賣"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellCurrentCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellCurrentPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計賣中"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellWinningCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellWinningPrice, 0)
        row = row + 1
        
        
    Call exwb.SaveAs(TargetPath)
    Call exap.Quit
    
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Sub WeekAccount(ByVal TargetPath As String)
    Dim exap As Excel.Application
    Dim exwb As Excel.Workbook
    Dim exsh As Excel.Worksheet
    Dim row As Integer, col As Integer
    Set exap = New Excel.Application
    Set exwb = exap.Workbooks.Add
    Set exsh = exwb.Sheets.Item(1)
    exsh.Columns.ColumnWidth = 16
    
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String
    Dim CurrentCount As Double, CurrentPrice As Long, WinningCount As Double, WinningPrice As Long
    Dim CurrentCountAll(1024) As Double, WinningCountAll(1024) As Double
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset, rec2 As New adoDB.Recordset
    Dim BuyCurrentCount As Double, SellCurrentCount As Double, BuyCurrentPrice As Long, SellCurrentPrice As Long
    Dim BuyWinningCount As Double, SellWinningCount As Double, BuyWinningPrice As Long, SellWinningPrice As Long
    Dim OrderDate As String
    Dim Proportion As Integer, BonusTarget As String, BonusMoney As Long, BonusProduct As String
    Dim CID As String, CName As String
    
    
    BuyCurrentCount = 0
    BuyWinningCount = 0
    BuyCurrentPrice = 0
    BuyWinningPrice = 0
    SellCurrentCount = 0
    SellWinningCount = 0
    SellCurrentPrice = 0
    SellWinningPrice = 0
    
    
    SQL = "select * from product order by CLng(PID);"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom order by CLng(CID);"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    
    
        
        'show report datetime
        row = 1: col = 1: exsh.Cells(row, col) = DateTime.DateAdd("d", -6, txtCurrentDate.Text) & "至" & txtCurrentDate.Text
        col = 2:   Call exsh.Range(Numeric2CharEN(1) & row, Numeric2CharEN(col) & row).Merge
        col = col + 1: exsh.Cells(row, col) = "週總計"
        row = row + 1
        
        'show product name
        col = 1:  exsh.Cells(row, col) = "產品"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName")
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName") & "交易金額"
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName") & "中"
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName") & "中獎金額"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        col = col + 1: exsh.Cells(row, col) = "成數"
        col = col + 1: exsh.Cells(row, col) = "前帳"
        row = row + 1
        
        
        'show every custom order per product
        product_rec.MoveFirst
        Do Until custom_rec.EOF
            'mark custom name
            CID = custom_rec.Fields.Item("CID")
            CName = custom_rec.Fields.Item("CName")
            col = 1: exsh.Cells(row, col) = CName


            'show every custom order per product
            For i = 0 To Count - 1
                'query a new custom order
                selectFields = "CurrentCount,WinningCount,CurrentDate"
                SQL = "select " & selectFields & " from [order] where [order].PID='" & PIDArray(i) & "' and CID='" & CID & "' and (CurrentDate>='" & Format(DateTime.DateAdd("d", -6, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd") & "' and CurrentDate<='" & txtCurrentDate.Text & "') order by [order].PID;"
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
                
                    CurrentCount = CurrentCount + SimpleRound(Val(rec1.Fields.Item("CurrentCount")), 4)
                    WinningCount = WinningCount + SimpleRound(Val(rec1.Fields.Item("WinningCount")), 4)
                    If Not price_rec.EOF Then
                        CurrentPrice = CurrentPrice + SimpleRound(Val(rec1.Fields.Item("CurrentCount")) * Val(price_rec.Fields.Item("CurrentPrice")), 0)
                        WinningPrice = WinningPrice + SimpleRound(Val(rec1.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")), 0)
                    End If
                    
                    rec1.MoveNext
                Loop
                col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentCount, 4)
                col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentPrice, 0)
                col = col + 1: exsh.Cells(row, col) = SimpleRound(WinningCount, 4)
                col = col + 1: exsh.Cells(row, col) = SimpleRound(WinningPrice, 0)
                
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
                SQL = "select * from [order] where CID='" & BonusTarget & "' and (CurrentDate>='" & Format(DateTime.DateAdd("d", -6, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd") & "' and CurrentDate<='" & txtCurrentDate.Text & "');"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec2)
                Do Until rec2.EOF
                    'search price, and addition to variable
                    OrderDate = rec2.Fields.Item("CurrentDate")
                    BonusProduct = rec2.Fields.Item("PID")
                    selectFields = "CurrentPrice,WinningPrice,Upset"
                    SQL = "select * from price where PID='" & BonusProduct & "' and CID='" & BonusTarget & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
                    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
                    
                    If Not price_rec.EOF Then
                        BonusMoney = BonusMoney + SimpleRound(Val(rec2.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")), 0)
                    End If
                    rec2.MoveNext
                Loop
                
                col = col + 1: exsh.Cells(row, col) = "來自" & rec1.Fields.Item("CName") & "成數" & Proportion & "%共" & SimpleRound(BonusMoney * (Proportion / 100), 0) & "元。"
                rec1.MoveNext
            Loop
            
            row = row + 1
            
            custom_rec.MoveNext
        Loop
        
        
        col = 1: exsh.Cells(row, col) = "數量總計"
        For i = 0 To Count - 1
            col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentCountAll(i), 4):   col = col + 1
            col = col + 1: exsh.Cells(row, col) = SimpleRound(WinningCountAll(i), 4):   col = col + 1
        Next
        row = row + 1
        
              
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計買"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyCurrentCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyCurrentPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計買中"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyWinningCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyWinningPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計賣"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellCurrentCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellCurrentPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計賣中"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellWinningCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellWinningPrice, 0)
        row = row + 1
        
        
    Call exwb.SaveAs(TargetPath)
    Call exap.Quit
    
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Sub MonthAccount(ByVal TargetPath As String)
    Dim exap As Excel.Application
    Dim exwb As Excel.Workbook
    Dim exsh As Excel.Worksheet
    Dim row As Integer, col As Integer
    Set exap = New Excel.Application
    Set exwb = exap.Workbooks.Add
    Set exsh = exwb.Sheets.Item(1)
    exsh.Columns.ColumnWidth = 16
    
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String
    Dim CurrentCount As Double, CurrentPrice As Long, WinningCount As Double, WinningPrice As Long
    Dim CurrentCountAll(1024) As Double, WinningCountAll(1024) As Double
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset, rec2 As New adoDB.Recordset
    Dim BuyCurrentCount As Double, SellCurrentCount As Double, BuyCurrentPrice As Long, SellCurrentPrice As Long
    Dim BuyWinningCount As Double, SellWinningCount As Double, BuyWinningPrice As Long, SellWinningPrice As Long
    Dim OrderDate As String
    Dim Proportion As Integer, BonusTarget As String, BonusMoney As Long, BonusProduct As String
    Dim CID As String, CName As String
    
    
    BuyCurrentCount = 0
    BuyWinningCount = 0
    BuyCurrentPrice = 0
    BuyWinningPrice = 0
    SellCurrentCount = 0
    SellWinningCount = 0
    SellCurrentPrice = 0
    SellWinningPrice = 0
    
    
    SQL = "select * from product order by CLng(PID);"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom order by CLng(CID);"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    
    
           
        'show report datetime
        row = 1: col = 1: exsh.Cells(row, col) = Format(txtCurrentDate.Text, "yyyy/MM")
        col = 2:   Call exsh.Range(Numeric2CharEN(1) & row, Numeric2CharEN(col) & row).Merge
        col = col + 1: exsh.Cells(row, col) = "月總計"
        row = row + 1
        
        'show product name
        col = 1:  exsh.Cells(row, col) = "產品"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName")
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName") & "交易金額"
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName") & "中"
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName") & "中獎金額"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        col = col + 1: exsh.Cells(row, col) = "成數"
        row = row + 1
        
        
        'show every custom order per product
        product_rec.MoveFirst
        Do Until custom_rec.EOF
            'mark custom name
            CID = custom_rec.Fields.Item("CID")
            CName = custom_rec.Fields.Item("CName")
            col = 1: exsh.Cells(row, col) = CName


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
                
                    CurrentCount = CurrentCount + SimpleRound(Val(rec1.Fields.Item("CurrentCount")), 4)
                    WinningCount = WinningCount + SimpleRound(Val(rec1.Fields.Item("WinningCount")), 4)
                    If Not price_rec.EOF Then
                        CurrentPrice = CurrentPrice + SimpleRound(Val(rec1.Fields.Item("CurrentCount")) * Val(price_rec.Fields.Item("CurrentPrice")), 0)
                        WinningPrice = WinningPrice + SimpleRound(Val(rec1.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")), 0)
                    End If
                    
                    rec1.MoveNext
                Loop
                col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentCount, 4)
                col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentPrice, 0)
                col = col + 1: exsh.Cells(row, col) = SimpleRound(WinningCount, 4)
                col = col + 1: exsh.Cells(row, col) = SimpleRound(WinningPrice, 0)
                
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
                        BonusMoney = BonusMoney + SimpleRound(Val(rec2.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")), 0)
                    End If
                    rec2.MoveNext
                Loop
                
                col = col + 1: exsh.Cells(row, col) = "來自" & rec1.Fields.Item("CName") & "成數" & Proportion & "%共" & SimpleRound(BonusMoney * (Proportion / 100), 0) & "元。"
                rec1.MoveNext
            Loop
            
            row = row + 1
            
            custom_rec.MoveNext
        Loop
        
        
        
        col = 1: exsh.Cells(row, col) = "數量總計"
        For i = 0 To Count - 1
            col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentCountAll(i), 4):   col = col + 1
            col = col + 1: exsh.Cells(row, col) = SimpleRound(WinningCountAll(i), 4):   col = col + 1
        Next
        row = row + 1
        
              
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計買"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyCurrentCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyCurrentPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計買中"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyWinningCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyWinningPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計賣"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellCurrentCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellCurrentPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計賣中"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellWinningCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellWinningPrice, 0)
        row = row + 1
        
        
    Call exwb.SaveAs(TargetPath)
    Call exap.Quit
    
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Sub YearAccount(ByVal TargetPath As String)
    Dim exap As Excel.Application
    Dim exwb As Excel.Workbook
    Dim exsh As Excel.Worksheet
    Dim row As Integer, col As Integer
    Set exap = New Excel.Application
    Set exwb = exap.Workbooks.Add
    Set exsh = exwb.Sheets.Item(1)
    exsh.Columns.ColumnWidth = 16
    
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String
    Dim CurrentCount As Double, CurrentPrice As Long, WinningCount As Double, WinningPrice As Long
    Dim CurrentCountAll(1024) As Double, WinningCountAll(1024) As Double
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset, rec2 As New adoDB.Recordset
    Dim BuyCurrentCount As Double, SellCurrentCount As Double, BuyCurrentPrice As Long, SellCurrentPrice As Long
    Dim BuyWinningCount As Double, SellWinningCount As Double, BuyWinningPrice As Long, SellWinningPrice As Long
    Dim OrderDate As String
    Dim Proportion As Integer, BonusTarget As String, BonusMoney As Long, BonusProduct As String
    Dim CID As String, CName As String
    
    
    BuyCurrentCount = 0
    BuyWinningCount = 0
    BuyCurrentPrice = 0
    BuyWinningPrice = 0
    SellCurrentCount = 0
    SellWinningCount = 0
    SellCurrentPrice = 0
    SellWinningPrice = 0
    
    
    SQL = "select * from product order by CLng(PID);"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom order by CLng(CID);"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    
    
       
        'show report datetime
        row = 1: col = 1: exsh.Cells(row, col) = Format(txtCurrentDate.Text, "yyyy")
        col = 2:   Call exsh.Range(Numeric2CharEN(1) & row, Numeric2CharEN(col) & row).Merge
        col = col + 1: exsh.Cells(row, col) = "年總計"
        row = row + 1
        
        'show product name
        col = 1:  exsh.Cells(row, col) = "產品"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName")
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName") & "交易金額"
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName") & "中"
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName") & "中獎金額"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        col = col + 1: exsh.Cells(row, col) = "成數"
        row = row + 1
        
        'show every custom order per product
        product_rec.MoveFirst
        Do Until custom_rec.EOF
            'mark custom name
            CID = custom_rec.Fields.Item("CID")
            CName = custom_rec.Fields.Item("CName")
            col = 1:  exsh.Cells(row, col) = CName


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
                
                    CurrentCount = CurrentCount + SimpleRound(Val(rec1.Fields.Item("CurrentCount")), 4)
                    WinningCount = WinningCount + SimpleRound(Val(rec1.Fields.Item("WinningCount")), 4)
                    If Not price_rec.EOF Then
                        CurrentPrice = CurrentPrice + SimpleRound(Val(rec1.Fields.Item("CurrentCount")) * Val(price_rec.Fields.Item("CurrentPrice")), 0)
                        WinningPrice = WinningPrice + SimpleRound(Val(rec1.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")), 0)
                    End If
                    
                    rec1.MoveNext
                Loop
                col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentCount, 4)
                col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentPrice, 0)
                col = col + 1: exsh.Cells(row, col) = SimpleRound(WinningCount, 4)
                col = col + 1: exsh.Cells(row, col) = SimpleRound(WinningPrice, 0)
                
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
                        BonusMoney = BonusMoney + SimpleRound(Val(rec2.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")), 0)
                    End If
                    rec2.MoveNext
                Loop
                
                col = col + 1: exsh.Cells(row, col) = "來自" & rec1.Fields.Item("CName") & "成數" & Proportion & "%共" & SimpleRound(BonusMoney * (Proportion / 100), 0) & "元。"
                rec1.MoveNext
            Loop
            
            row = row + 1
            
            custom_rec.MoveNext
        Loop
        
        
        col = 1: exsh.Cells(row, col) = "數量總計"
        For i = 0 To Count - 1
            col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentCountAll(i), 4):   col = col + 1
            col = col + 1: exsh.Cells(row, col) = SimpleRound(WinningCountAll(i), 4):   col = col + 1
        Next
        row = row + 1
        
              
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計買"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyCurrentCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyCurrentPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計買中"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyWinningCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyWinningPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計賣"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellCurrentCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellCurrentPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計賣中"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellWinningCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellWinningPrice, 0)
        row = row + 1
        
        
    Call exwb.SaveAs(TargetPath)
    Call exap.Quit
    
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub


Sub FourKDayReport(ByVal TargetPath As String)
    Dim exap As Excel.Application
    Dim exwb As Excel.Workbook
    Dim exsh As Excel.Worksheet
    Dim row As Integer, col As Integer
    Set exap = New Excel.Application
    Set exwb = exap.Workbooks.Add
    Set exsh = exwb.Sheets.Item(1)
    exsh.Columns.ColumnWidth = 16
    
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String
    Dim CurrentCount As Double, CurrentPrice As Long, WinningCount As Double, WinningPrice As Long
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset, rec2 As New adoDB.Recordset
    Dim BuyCurrentCount As Double, SellCurrentCount As Double, BuyCurrentPrice As Long, SellCurrentPrice As Long
    Dim BuyWinningCount As Double, SellWinningCount As Double, BuyWinningPrice As Long, SellWinningPrice As Long
    Dim OrderDate As String
    Dim Proportion As Integer, BonusTarget As String, BonusMoney As Long, BonusProduct As String
    Dim CID As String, CName As String
    Dim PriceCount As Long
    
    
    BuyCurrentCount = 0
    BuyWinningCount = 0
    BuyCurrentPrice = 0
    BuyWinningPrice = 0
    SellCurrentCount = 0
    SellWinningCount = 0
    SellCurrentPrice = 0
    SellWinningPrice = 0
    
    
    SQL = "select * from product where PName like '539_4K' order by CLng(PID);" '只要顯示539的4K 'SQL = "select * from product where PName like '%4K%' order by PID;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom order by CLng(CID);"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    
    
        
        'show report datetime
        row = 1: col = 1: exsh.Cells(row, col) = "日期"
        col = col + 1: exsh.Cells(row, col) = txtCurrentDate.Text
        col = 11:   Call exsh.Range(Numeric2CharEN(2) & row, Numeric2CharEN(col) & row).Merge
        row = row + 1
        
        'show product name
        col = 1:  exsh.Cells(row, col) = "產品"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName") & "交易數量"
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName") & "中獎數量"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        col = col + 1: exsh.Cells(row, col) = "客戶金額總計"
        row = row + 1
        
        
        'show every custom order per product
        product_rec.MoveFirst
        Do Until custom_rec.EOF
            'mark custom name
            CID = custom_rec.Fields.Item("CID")
            CName = custom_rec.Fields.Item("CName")
            col = 1: exsh.Cells(row, col) = CName


            'show every custom order per product
            PriceCount = 0
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
                
                    CurrentCount = CurrentCount + SimpleRound(Val(rec1.Fields.Item("CurrentCount")), 4)
                    WinningCount = WinningCount + SimpleRound(Val(rec1.Fields.Item("WinningCount")), 4)
                    If Not price_rec.EOF Then
                        CurrentPrice = CurrentPrice + SimpleRound(Val(rec1.Fields.Item("CurrentCount")) * Val(price_rec.Fields.Item("CurrentPrice")), 0)
                        WinningPrice = WinningPrice + SimpleRound(Val(rec1.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")), 0)
                    End If
                    
                    rec1.MoveNext
                Loop
                col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentCount, 4)
                col = col + 1: exsh.Cells(row, col) = SimpleRound(WinningCount, 4)
                PriceCount = PriceCount + SimpleRound(CurrentPrice - WinningPrice, 0)
                
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
            col = col + 1: exsh.Cells(row, col) = SimpleRound(PriceCount, 0)
            row = row + 1
            
            custom_rec.MoveNext
        Loop
        
        
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計買"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyCurrentCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyCurrentPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計買中"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyWinningCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyWinningPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計賣"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellCurrentCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellCurrentPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計賣中"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellWinningCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellWinningPrice, 0)
        row = row + 1
        
        
    Call exwb.SaveAs(TargetPath)
    Call exap.Quit
    
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Sub FourKWeekReport(ByVal TargetPath As String)
    Dim exap As Excel.Application
    Dim exwb As Excel.Workbook
    Dim exsh As Excel.Worksheet
    Dim row As Integer, col As Integer
    Set exap = New Excel.Application
    Set exwb = exap.Workbooks.Add
    Set exsh = exwb.Sheets.Item(1)
    exsh.Columns.ColumnWidth = 16
    
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String
    Dim CurrentCount As Double, CurrentPrice As Long, WinningCount As Double, WinningPrice As Long
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset, rec2 As New adoDB.Recordset
    Dim BuyCurrentCount As Double, SellCurrentCount As Double, BuyCurrentPrice As Long, SellCurrentPrice As Long
    Dim BuyWinningCount As Double, SellWinningCount As Double, BuyWinningPrice As Long, SellWinningPrice As Long
    Dim OrderDate As String
    Dim Proportion As Integer, BonusTarget As String, BonusMoney As Long, BonusProduct As String
    Dim CID As String, CName As String
    
    
    BuyCurrentCount = 0
    BuyWinningCount = 0
    BuyCurrentPrice = 0
    BuyWinningPrice = 0
    SellCurrentCount = 0
    SellWinningCount = 0
    SellCurrentPrice = 0
    SellWinningPrice = 0
    
    
    SQL = "select * from product where PName like '539_4K' order by CLng(PID);" '只要顯示539的4K 'SQL = "select * from product where PName like '%4K%' order by PID;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom order by CLng(CID);"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    
    
                
        'show report datetime
        row = 1: col = 1: exsh.Cells(row, col) = "日期"
        col = col + 1: exsh.Cells(row, col) = DateTime.DateAdd("d", -6, txtCurrentDate.Text) & "至" & txtCurrentDate.Text
        col = 11:   Call exsh.Range(Numeric2CharEN(2) & row, Numeric2CharEN(col) & row).Merge
        row = row + 1
        
        'show product name
        col = 1:  exsh.Cells(row, col) = "產品"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName")
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName") & "交易金額"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        col = col + 1: exsh.Cells(row, col) = "成數"
        col = col + 1: exsh.Cells(row, col) = "前帳"
        row = row + 1
        
        
        
        'show every custom order per product
        product_rec.MoveFirst
        Do Until custom_rec.EOF
            'mark custom name
            CID = custom_rec.Fields.Item("CID")
            CName = custom_rec.Fields.Item("CName")
            col = 1: exsh.Cells(row, col) = CName


            'show every custom order per product
            For i = 0 To Count - 1
                'query a new custom order
                selectFields = "CurrentCount,WinningCount,CurrentDate"
                SQL = "select " & selectFields & " from [order] where [order].PID='" & PIDArray(i) & "' and CID='" & CID & "' and (CurrentDate>='" & Format(DateTime.DateAdd("d", -6, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd") & "' and CurrentDate<='" & txtCurrentDate.Text & "') order by [order].PID;"
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
                
                    CurrentCount = CurrentCount + SimpleRound(Val(rec1.Fields.Item("CurrentCount")), 4)
                    WinningCount = WinningCount + SimpleRound(Val(rec1.Fields.Item("WinningCount")), 4)
                    If Not price_rec.EOF Then
                        CurrentPrice = CurrentPrice + SimpleRound(Val(rec1.Fields.Item("CurrentCount")) * Val(price_rec.Fields.Item("CurrentPrice")), 0)
                        WinningPrice = WinningPrice + SimpleRound(Val(rec1.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")), 0)
                    End If
                    
                    rec1.MoveNext
                Loop
                col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentCount, 4)
                col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentPrice, 0)
                
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
                SQL = "select * from [order] where CID='" & BonusTarget & "' and (CurrentDate>='" & Format(DateTime.DateAdd("d", -6, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd") & "' and CurrentDate<='" & txtCurrentDate.Text & "');"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec2)
                Do Until rec2.EOF
                    'search price, and addition to variable
                    OrderDate = rec2.Fields.Item("CurrentDate")
                    BonusProduct = rec2.Fields.Item("PID")
                    selectFields = "CurrentPrice,WinningPrice,Upset"
                    SQL = "select * from price where PID='" & BonusProduct & "' and CID='" & BonusTarget & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
                    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
                    
                    If Not price_rec.EOF Then
                        BonusMoney = BonusMoney + SimpleRound(Val(rec2.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")), 0)
                    End If
                    rec2.MoveNext
                Loop
                
                col = col + 1: exsh.Cells(row, col) = "來自" & rec1.Fields.Item("CName") & "成數" & Proportion & "%共" & SimpleRound(BonusMoney * (Proportion / 100), 0) & "元。"
                rec1.MoveNext
            Loop
            
            row = row + 1
            
            custom_rec.MoveNext
        Loop

        
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計買"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyCurrentCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyCurrentPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計買中"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyWinningCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyWinningPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計賣"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellCurrentCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellCurrentPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計賣中"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellWinningCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellWinningPrice, 0)
        row = row + 1
        
        
    Call exwb.SaveAs(TargetPath)
    Call exap.Quit
    
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Sub FourKMonthReport(ByVal TargetPath As String)
    Dim exap As Excel.Application
    Dim exwb As Excel.Workbook
    Dim exsh As Excel.Worksheet
    Dim row As Integer, col As Integer
    Set exap = New Excel.Application
    Set exwb = exap.Workbooks.Add
    Set exsh = exwb.Sheets.Item(1)
    exsh.Columns.ColumnWidth = 16
    
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String
    Dim CurrentCount As Double, CurrentPrice As Long, WinningCount As Double, WinningPrice As Long
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset, rec2 As New adoDB.Recordset
    Dim BuyCurrentCount As Double, SellCurrentCount As Double, BuyCurrentPrice As Long, SellCurrentPrice As Long
    Dim BuyWinningCount As Double, SellWinningCount As Double, BuyWinningPrice As Long, SellWinningPrice As Long
    Dim OrderDate As String
    Dim Proportion As Integer, BonusTarget As String, BonusMoney As Long, BonusProduct As String
    Dim CID As String, CName As String
    
    
    BuyCurrentCount = 0
    BuyWinningCount = 0
    BuyCurrentPrice = 0
    BuyWinningPrice = 0
    SellCurrentCount = 0
    SellWinningCount = 0
    SellCurrentPrice = 0
    SellWinningPrice = 0
    
    
    SQL = "select * from product where PName like '539_4K' order by CLng(PID);" '只要顯示539的4K 'SQL = "select * from product where PName like '%4K%' order by PID;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom order by CLng(CID);"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    

           
        'show report datetime
        row = 1: col = 1: exsh.Cells(row, col) = "日期"
        col = col + 1: exsh.Cells(row, col) = Format(txtCurrentDate.Text, "yyyy/MM") & "月"
        col = 11:   Call exsh.Range(Numeric2CharEN(2) & row, Numeric2CharEN(col) & row).Merge
        row = row + 1
        
        'show product name
        col = 1:  exsh.Cells(row, col) = "產品"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName")
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName") & "交易金額"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        col = col + 1: exsh.Cells(row, col) = "成數"
        row = row + 1
        
        
        'show every custom order per product
        product_rec.MoveFirst
        Do Until custom_rec.EOF
            'mark custom name
            CID = custom_rec.Fields.Item("CID")
            CName = custom_rec.Fields.Item("CName")
            col = 1: exsh.Cells(row, col) = CName


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
                
                    CurrentCount = CurrentCount + SimpleRound(Val(rec1.Fields.Item("CurrentCount")), 4)
                    WinningCount = WinningCount + SimpleRound(Val(rec1.Fields.Item("WinningCount")), 4)
                    If Not price_rec.EOF Then
                        CurrentPrice = CurrentPrice + SimpleRound(Val(rec1.Fields.Item("CurrentCount")) * Val(price_rec.Fields.Item("CurrentPrice")), 0)
                        WinningPrice = WinningPrice + SimpleRound(Val(rec1.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")), 0)
                    End If
                    
                    rec1.MoveNext
                Loop
                col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentCount, 4)
                col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentPrice, 0)
                
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
                        BonusMoney = BonusMoney + SimpleRound(Val(rec2.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")), 0)
                    End If
                    rec2.MoveNext
                Loop
                
                col = col + 1: exsh.Cells(row, col) = "來自" & rec1.Fields.Item("CName") & "成數" & Proportion & "%共" & SimpleRound(BonusMoney * (Proportion / 100), 0) & "元。"
                rec1.MoveNext
            Loop
            
            row = row + 1
            
            custom_rec.MoveNext
        Loop
        
        
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計買"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyCurrentCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyCurrentPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計買中"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyWinningCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyWinningPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計賣"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellCurrentCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellCurrentPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計賣中"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellWinningCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellWinningPrice, 0)
        row = row + 1
        
        
    Call exwb.SaveAs(TargetPath)
    Call exap.Quit
    
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Sub FourKYearReport(ByVal TargetPath As String)
    Dim exap As Excel.Application
    Dim exwb As Excel.Workbook
    Dim exsh As Excel.Worksheet
    Dim row As Integer, col As Integer
    Set exap = New Excel.Application
    Set exwb = exap.Workbooks.Add
    Set exsh = exwb.Sheets.Item(1)
    exsh.Columns.ColumnWidth = 16
    
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String
    Dim CurrentCount As Double, CurrentPrice As Long, WinningCount As Double, WinningPrice As Long
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset, rec2 As New adoDB.Recordset
    Dim BuyCurrentCount As Double, SellCurrentCount As Double, BuyCurrentPrice As Long, SellCurrentPrice As Long
    Dim BuyWinningCount As Double, SellWinningCount As Double, BuyWinningPrice As Long, SellWinningPrice As Long
    Dim OrderDate As String
    Dim Proportion As Integer, BonusTarget As String, BonusMoney As Long, BonusProduct As String
    Dim CID As String, CName As String
    
    
    BuyCurrentCount = 0
    BuyWinningCount = 0
    BuyCurrentPrice = 0
    BuyWinningPrice = 0
    SellCurrentCount = 0
    SellWinningCount = 0
    SellCurrentPrice = 0
    SellWinningPrice = 0
    
    
    SQL = "select * from product where PName like '539_4K' order by CLng(PID);" '只要顯示539的4K 'SQL = "select * from product where PName like '%4K%' order by PID;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom order by CLng(CID);"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    

        
        'show report datetime
        row = 1: col = 1: exsh.Cells(row, col) = "日期"
        col = col + 1: exsh.Cells(row, col) = Format(txtCurrentDate.Text, "yyyy") & "年"
        col = 11:   Call exsh.Range(Numeric2CharEN(2) & row, Numeric2CharEN(col) & row).Merge
        row = row + 1
        
        'show product name
        col = 1:  exsh.Cells(row, col) = "產品"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName")
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName") & "交易金額"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        col = col + 1: exsh.Cells(row, col) = "成數"
        row = row + 1
        
        
        
        'show every custom order per product
        product_rec.MoveFirst
        Do Until custom_rec.EOF
            'mark custom name
            CID = custom_rec.Fields.Item("CID")
            CName = custom_rec.Fields.Item("CName")
            col = 1: exsh.Cells(row, col) = CName


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
                
                    CurrentCount = CurrentCount + SimpleRound(Val(rec1.Fields.Item("CurrentCount")), 4)
                    WinningCount = WinningCount + SimpleRound(Val(rec1.Fields.Item("WinningCount")), 4)
                    If Not price_rec.EOF Then
                        CurrentPrice = CurrentPrice + SimpleRound(Val(rec1.Fields.Item("CurrentCount")) * Val(price_rec.Fields.Item("CurrentPrice")), 0)
                        WinningPrice = WinningPrice + SimpleRound(Val(rec1.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")), 0)
                    End If
                    
                    rec1.MoveNext
                Loop
                col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentCount, 4)
                col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentPrice, 0)
                
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
                        BonusMoney = BonusMoney + SimpleRound(Val(rec2.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")), 0)
                    End If
                    rec2.MoveNext
                Loop
                
                col = col + 1: exsh.Cells(row, col) = "來自" & rec1.Fields.Item("CName") & "成數" & Proportion & "%共" & SimpleRound(BonusMoney * (Proportion / 100), 0) & "元。"
                rec1.MoveNext
            Loop
            
            row = row + 1
            
            custom_rec.MoveNext
        Loop
        
        
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計買"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyCurrentCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyCurrentPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計買中"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyWinningCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyWinningPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計賣"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellCurrentCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellCurrentPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計賣中"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellWinningCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellWinningPrice, 0)
        row = row + 1
        
        
    Call exwb.SaveAs(TargetPath)
    Call exap.Quit
    
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Sub FourKDayAccount(ByVal TargetPath As String)
    Dim exap As Excel.Application
    Dim exwb As Excel.Workbook
    Dim exsh As Excel.Worksheet
    Dim row As Integer, col As Integer
    Set exap = New Excel.Application
    Set exwb = exap.Workbooks.Add
    Set exsh = exwb.Sheets.Item(1)
    exsh.Columns.ColumnWidth = 16
    
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String
    Dim CurrentCount As Double, CurrentPrice As Long, WinningCount As Double, WinningPrice As Long
    Dim CurrentCountAll(1024) As Double, WinningCountAll(1024) As Double
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset, rec2 As New adoDB.Recordset
    Dim BuyCurrentCount As Double, SellCurrentCount As Double, BuyCurrentPrice As Long, SellCurrentPrice As Long
    Dim BuyWinningCount As Double, SellWinningCount As Double, BuyWinningPrice As Long, SellWinningPrice As Long
    Dim OrderDate As String
    Dim Proportion As Integer, BonusTarget As String, BonusMoney As Long, BonusProduct As String
    Dim CID As String, CName As String
    
    
    BuyCurrentCount = 0
    BuyWinningCount = 0
    BuyCurrentPrice = 0
    BuyWinningPrice = 0
    SellCurrentCount = 0
    SellWinningCount = 0
    SellCurrentPrice = 0
    SellWinningPrice = 0
    
    
    SQL = "select * from product where PName like '539_4K' order by CLng(PID);" '只要顯示539的4K 'SQL = "select * from product where PName like '%4K%' order by PID;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom order by CLng(CID);"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    
            
            
        'show report datetime
        row = 1: col = 1: exsh.Cells(row, col) = txtCurrentDate.Text
        col = 2:   Call exsh.Range(Numeric2CharEN(1) & row, Numeric2CharEN(col) & row).Merge
        col = col + 1: exsh.Cells(row, col) = "日總計"
        row = row + 1
        
        'show product name
        col = 1:  exsh.Cells(row, col) = "產品"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName")
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName") & "交易金額"
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName") & "中"
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName") & "中獎金額"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        row = row + 1

        
        
        'show every custom order per product
        product_rec.MoveFirst
        Do Until custom_rec.EOF
            'mark custom name
            CID = custom_rec.Fields.Item("CID")
            CName = custom_rec.Fields.Item("CName")
            col = 1:  exsh.Cells(row, col) = CName


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
                
                    CurrentCount = CurrentCount + SimpleRound(Val(rec1.Fields.Item("CurrentCount")), 4)
                    WinningCount = WinningCount + SimpleRound(Val(rec1.Fields.Item("WinningCount")), 4)
                    If Not price_rec.EOF Then
                        CurrentPrice = CurrentPrice + SimpleRound(Val(rec1.Fields.Item("CurrentCount")) * Val(price_rec.Fields.Item("CurrentPrice")), 0)
                        WinningPrice = WinningPrice + SimpleRound(Val(rec1.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")), 0)
                    End If
                    
                    rec1.MoveNext
                Loop
                col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentCount, 4)
                col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentPrice, 0)
                col = col + 1: exsh.Cells(row, col) = SimpleRound(WinningCount, 4)
                col = col + 1: exsh.Cells(row, col) = SimpleRound(WinningPrice, 0)
                
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
            row = row + 1
            
            custom_rec.MoveNext
        Loop
        
        
        col = 1: exsh.Cells(row, col) = "數量總計"
        For i = 0 To Count - 1
            col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentCountAll(i), 4):   col = col + 1
            col = col + 1: exsh.Cells(row, col) = SimpleRound(WinningCountAll(i), 4):   col = col + 1
        Next
        row = row + 1
        
              
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計買"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyCurrentCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyCurrentPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計買中"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyWinningCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyWinningPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計賣"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellCurrentCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellCurrentPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計賣中"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellWinningCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellWinningPrice, 0)
        row = row + 1
        
        
    Call exwb.SaveAs(TargetPath)
    Call exap.Quit
    
    
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Sub FourKWeekAccount(ByVal TargetPath As String)
    Dim exap As Excel.Application
    Dim exwb As Excel.Workbook
    Dim exsh As Excel.Worksheet
    Dim row As Integer, col As Integer
    Set exap = New Excel.Application
    Set exwb = exap.Workbooks.Add
    Set exsh = exwb.Sheets.Item(1)
    exsh.Columns.ColumnWidth = 16
    
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String
    Dim CurrentCount As Double, CurrentPrice As Long, WinningCount As Double, WinningPrice As Long
    Dim CurrentCountAll(1024) As Double, WinningCountAll(1024) As Double
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset, rec2 As New adoDB.Recordset
    Dim BuyCurrentCount As Double, SellCurrentCount As Double, BuyCurrentPrice As Long, SellCurrentPrice As Long
    Dim BuyWinningCount As Double, SellWinningCount As Double, BuyWinningPrice As Long, SellWinningPrice As Long
    Dim OrderDate As String
    Dim Proportion As Integer, BonusTarget As String, BonusMoney As Long, BonusProduct As String
    Dim CID As String, CName As String
    
    
    BuyCurrentCount = 0
    BuyWinningCount = 0
    BuyCurrentPrice = 0
    BuyWinningPrice = 0
    SellCurrentCount = 0
    SellWinningCount = 0
    SellCurrentPrice = 0
    SellWinningPrice = 0
    
    
    SQL = "select * from product where PName like '539_4K' order by CLng(PID);" '只要顯示539的4K 'SQL = "select * from product where PName like '%4K%' order by PID;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom order by CLng(CID);"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    

            
        'show report datetime
        row = 1: col = 1: exsh.Cells(row, col) = DateTime.DateAdd("d", -6, txtCurrentDate.Text) & "至" & txtCurrentDate.Text
        col = 2:   Call exsh.Range(Numeric2CharEN(1) & row, Numeric2CharEN(col) & row).Merge
        col = col + 1: exsh.Cells(row, col) = ">週總計"
        row = row + 1
        
        'show product name
        col = 1:  exsh.Cells(row, col) = "產品"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName")
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName") & "交易金額"
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName") & "中"
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName") & "中獎金額"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        col = col + 1: exsh.Cells(row, col) = "成數"
        col = col + 1: exsh.Cells(row, col) = "前帳"
        row = row + 1
        
        
        'show every custom order per product
        product_rec.MoveFirst
        Do Until custom_rec.EOF
            'mark custom name
            CID = custom_rec.Fields.Item("CID")
            CName = custom_rec.Fields.Item("CName")
            col = 1: exsh.Cells(row, col) = CName


            'show every custom order per product
            For i = 0 To Count - 1
                'query a new custom order
                selectFields = "CurrentCount,WinningCount,CurrentDate"
                SQL = "select " & selectFields & " from [order] where [order].PID='" & PIDArray(i) & "' and CID='" & CID & "' and (CurrentDate>='" & Format(DateTime.DateAdd("d", -6, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd") & "' and CurrentDate<='" & txtCurrentDate.Text & "') order by [order].PID;"
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
                
                    CurrentCount = CurrentCount + SimpleRound(Val(rec1.Fields.Item("CurrentCount")), 4)
                    WinningCount = WinningCount + SimpleRound(Val(rec1.Fields.Item("WinningCount")), 4)
                    If Not price_rec.EOF Then
                        CurrentPrice = CurrentPrice + SimpleRound(Val(rec1.Fields.Item("CurrentCount")) * Val(price_rec.Fields.Item("CurrentPrice")), 0)
                        WinningPrice = WinningPrice + SimpleRound(Val(rec1.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")), 0)
                    End If
                    
                    rec1.MoveNext
                Loop
                col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentCount, 4)
                col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentPrice, 0)
                col = col + 1: exsh.Cells(row, col) = SimpleRound(WinningCount, 4)
                col = col + 1: exsh.Cells(row, col) = SimpleRound(WinningPrice, 0)
                
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
                SQL = "select * from [order] where CID='" & BonusTarget & "' and (CurrentDate>='" & Format(DateTime.DateAdd("d", -6, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd") & "' and CurrentDate<='" & txtCurrentDate.Text & "');"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec2)
                Do Until rec2.EOF
                    'search price, and addition to variable
                    OrderDate = rec2.Fields.Item("CurrentDate")
                    BonusProduct = rec2.Fields.Item("PID")
                    selectFields = "CurrentPrice,WinningPrice,Upset"
                    SQL = "select * from price where PID='" & BonusProduct & "' and CID='" & BonusTarget & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
                    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
                    
                    If Not price_rec.EOF Then
                        BonusMoney = BonusMoney + SimpleRound(Val(rec2.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")), 0)
                    End If
                    rec2.MoveNext
                Loop
                
                col = col + 1: exsh.Cells(row, col) = "來自" & rec1.Fields.Item("CName") & "成數" & Proportion & "%共" & SimpleRound(BonusMoney * (Proportion / 100), 0) & "元。"
                rec1.MoveNext
            Loop
            row = row + 1
            
            custom_rec.MoveNext
        Loop
        
        
        
        col = 1: exsh.Cells(row, col) = "數量總計"
        For i = 0 To Count - 1
            col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentCountAll(i), 4):   col = col + 1
            col = col + 1: exsh.Cells(row, col) = SimpleRound(WinningCountAll(i), 4):   col = col + 1
        Next
        row = row + 1
        
              
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計買"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyCurrentCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyCurrentPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計買中"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyWinningCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyWinningPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計賣"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellCurrentCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellCurrentPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計賣中"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellWinningCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellWinningPrice, 0)
        row = row + 1
        
        
    Call exwb.SaveAs(TargetPath)
    Call exap.Quit
    
    
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Sub FourKMonthAccount(ByVal TargetPath As String)
    Dim exap As Excel.Application
    Dim exwb As Excel.Workbook
    Dim exsh As Excel.Worksheet
    Dim row As Integer, col As Integer
    Set exap = New Excel.Application
    Set exwb = exap.Workbooks.Add
    Set exsh = exwb.Sheets.Item(1)
    exsh.Columns.ColumnWidth = 16
    
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String
    Dim CurrentCount As Double, CurrentPrice As Long, WinningCount As Double, WinningPrice As Long
    Dim CurrentCountAll(1024) As Double, WinningCountAll(1024) As Double
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset, rec2 As New adoDB.Recordset
    Dim BuyCurrentCount As Double, SellCurrentCount As Double, BuyCurrentPrice As Long, SellCurrentPrice As Long
    Dim BuyWinningCount As Double, SellWinningCount As Double, BuyWinningPrice As Long, SellWinningPrice As Long
    Dim OrderDate As String
    Dim Proportion As Integer, BonusTarget As String, BonusMoney As Long, BonusProduct As String
    Dim CID As String, CName As String
    
    
    BuyCurrentCount = 0
    BuyWinningCount = 0
    BuyCurrentPrice = 0
    BuyWinningPrice = 0
    SellCurrentCount = 0
    SellWinningCount = 0
    SellCurrentPrice = 0
    SellWinningPrice = 0
    
    
    SQL = "select * from product where PName like '539_4K' order by CLng(PID);" '只要顯示539的4K 'SQL = "select * from product where PName like '%4K%' order by PID;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom order by CLng(CID);"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    
       
       
        'show report datetime
        row = 1: col = 1: exsh.Cells(row, col) = Format(txtCurrentDate.Text, "yyyy/MM")
        col = 2:   Call exsh.Range(Numeric2CharEN(1) & row, Numeric2CharEN(col) & row).Merge
        col = col + 1: exsh.Cells(row, col) = ">月總計"
        row = row + 1
        
        'show product name
        col = 1:  exsh.Cells(row, col) = "產品"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName")
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName") & "交易金額"
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName") & "中"
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName") & "中獎金額"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        col = col + 1: exsh.Cells(row, col) = "成數"
        row = row + 1
        
        
        'show every custom order per product
        product_rec.MoveFirst
        Do Until custom_rec.EOF
            'mark custom name
            CID = custom_rec.Fields.Item("CID")
            CName = custom_rec.Fields.Item("CName")
            col = 1: exsh.Cells(row, col) = CName


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
                
                    CurrentCount = CurrentCount + SimpleRound(Val(rec1.Fields.Item("CurrentCount")), 4)
                    WinningCount = WinningCount + SimpleRound(Val(rec1.Fields.Item("WinningCount")), 4)
                    If Not price_rec.EOF Then
                        CurrentPrice = CurrentPrice + SimpleRound(Val(rec1.Fields.Item("CurrentCount")) * Val(price_rec.Fields.Item("CurrentPrice")), 0)
                        WinningPrice = WinningPrice + SimpleRound(Val(rec1.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")), 0)
                    End If
                    
                    rec1.MoveNext
                Loop
                col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentCount, 4)
                col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentPrice, 0)
                col = col + 1: exsh.Cells(row, col) = SimpleRound(WinningCount, 4)
                col = col + 1: exsh.Cells(row, col) = SimpleRound(WinningPrice, 0)
                
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
                        BonusMoney = BonusMoney + SimpleRound(Val(rec2.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")), 0)
                    End If
                    rec2.MoveNext
                Loop
                
                col = col + 1: exsh.Cells(row, col) = "來自" & rec1.Fields.Item("CName") & "成數" & Proportion & "%共" & SimpleRound(BonusMoney * (Proportion / 100), 0) & "元。"
                rec1.MoveNext
            Loop
            
            row = row + 1
            
            custom_rec.MoveNext
        Loop
        
        
        col = 1: exsh.Cells(row, col) = "數量總計"
        For i = 0 To Count - 1
            col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentCountAll(i), 4):   col = col + 1
            col = col + 1: exsh.Cells(row, col) = SimpleRound(WinningCountAll(i), 4):   col = col + 1
        Next
        row = row + 1
        
              
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計買"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyCurrentCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyCurrentPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計買中"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyWinningCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyWinningPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計賣"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellCurrentCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellCurrentPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計賣中"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellWinningCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellWinningPrice, 0)
        row = row + 1
        
        
    Call exwb.SaveAs(TargetPath)
    Call exap.Quit
    
    
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub

Sub FourKYearAccount(ByVal TargetPath As String)
    Dim exap As Excel.Application
    Dim exwb As Excel.Workbook
    Dim exsh As Excel.Worksheet
    Dim row As Integer, col As Integer
    Set exap = New Excel.Application
    Set exwb = exap.Workbooks.Add
    Set exsh = exwb.Sheets.Item(1)
    exsh.Columns.ColumnWidth = 16
    
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer, PID As String
    Dim CurrentCount As Double, CurrentPrice As Long, WinningCount As Double, WinningPrice As Long
    Dim CurrentCountAll(1024) As Double, WinningCountAll(1024) As Double
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset, rec2 As New adoDB.Recordset
    Dim BuyCurrentCount As Double, SellCurrentCount As Double, BuyCurrentPrice As Long, SellCurrentPrice As Long
    Dim BuyWinningCount As Double, SellWinningCount As Double, BuyWinningPrice As Long, SellWinningPrice As Long
    Dim OrderDate As String
    Dim Proportion As Integer, BonusTarget As String, BonusMoney As Long, BonusProduct As String
    Dim CID As String, CName As String
    
    
    BuyCurrentCount = 0
    BuyWinningCount = 0
    BuyCurrentPrice = 0
    BuyWinningPrice = 0
    SellCurrentCount = 0
    SellWinningCount = 0
    SellCurrentPrice = 0
    SellWinningPrice = 0
    
    
    SQL = "select * from product where PName like '539_4K' order by CLng(PID);" '只要顯示539的4K 'SQL = "select * from product where PName like '%4K%' order by PID;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    
    SQL = "select * from custom order by CLng(CID);"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
    
    
        'show report datetime
        row = 1: col = 1: exsh.Cells(row, col) = Format(txtCurrentDate.Text, "yyyy")
        col = 2:   Call exsh.Range(Numeric2CharEN(1) & row, Numeric2CharEN(col) & row).Merge
        col = col + 1: exsh.Cells(row, col) = ">年總計"
        row = row + 1
        
        'show product name
        col = 1:  exsh.Cells(row, col) = "產品"
        Count = 0
        Do Until product_rec.EOF
            PIDArray(Count) = product_rec.Fields.Item("PID")
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName")
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName") & "交易金額"
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName") & "中"
            col = col + 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName") & "中獎金額"
            Count = Count + 1
            product_rec.MoveNext
        Loop
        col = col + 1: exsh.Cells(row, col) = "成數"
        row = row + 1
        
        
        'show every custom order per product
        product_rec.MoveFirst
        Do Until custom_rec.EOF
            'mark custom name
            CID = custom_rec.Fields.Item("CID")
            CName = custom_rec.Fields.Item("CName")
            col = 1: exsh.Cells(row, col) = CName


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
                
                    CurrentCount = CurrentCount + SimpleRound(Val(rec1.Fields.Item("CurrentCount")), 4)
                    WinningCount = WinningCount + SimpleRound(Val(rec1.Fields.Item("WinningCount")), 4)
                    If Not price_rec.EOF Then
                        CurrentPrice = CurrentPrice + SimpleRound(Val(rec1.Fields.Item("CurrentCount")) * Val(price_rec.Fields.Item("CurrentPrice")), 0)
                        WinningPrice = WinningPrice + SimpleRound(Val(rec1.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")), 0)
                    End If
                    
                    rec1.MoveNext
                Loop
                col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentCount, 4)
                col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentPrice, 0)
                col = col + 1: exsh.Cells(row, col) = SimpleRound(WinningCount, 4)
                col = col + 1: exsh.Cells(row, col) = SimpleRound(WinningPrice, 0)
                
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
                        BonusMoney = BonusMoney + SimpleRound(Val(rec2.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")), 0)
                    End If
                    rec2.MoveNext
                Loop
                
                col = col + 1: exsh.Cells(row, col) = "來自" & rec1.Fields.Item("CName") & "成數" & Proportion & "%共" & SimpleRound(BonusMoney * (Proportion / 100), 0) & "元。"
                rec1.MoveNext
            Loop
            
            row = row + 1
            
            custom_rec.MoveNext
        Loop
        
        

        col = 1: exsh.Cells(row, col) = "數量總計"
        For i = 0 To Count - 1
            col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentCountAll(i), 4):   col = col + 1
            col = col + 1: exsh.Cells(row, col) = SimpleRound(WinningCountAll(i), 4):   col = col + 1
        Next
        row = row + 1
        
              
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計買"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyCurrentCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyCurrentPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計買中"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyWinningCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BuyWinningPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計賣"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellCurrentCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellCurrentPrice, 0)
        row = row + 1
        col = 1: exsh.Cells(row, col) = "小計賣中"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellWinningCount, 4)
        col = col + 2: exsh.Cells(row, col) = "金額"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(SellWinningPrice, 0)
        row = row + 1
        
        
    Call exwb.SaveAs(TargetPath)
    Call exap.Quit
    
    
    product_rec.Close
    custom_rec.Close
    rec1.Close
End Sub


Sub CustomProductDayReportDetail(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer
    Dim CurrentCount As Double, CurrentPrice As Long, WinningCount As Double, WinningPrice As Long
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    Dim CData() As String
    Dim PData() As String
    Dim beginv As Integer
    Dim endv As Integer
    Dim OrderDate As String, ProductID As String, ProductName As String
    Dim PriceCount As Long
    
        
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
            
            Body = Body & "<td>" & SimpleRound(order_rec.Fields.Item("CurrentCount"), 4) & "</td>"
            If price_rec.EOF Then
                Body = Body & "<td>0</td>"
            Else
                Body = Body & "<td>" & SimpleRound(order_rec.Fields.Item("CurrentCount") * Val(price_rec.Fields.Item("CurrentPrice")), 0) & "</td>"
                PriceCount = PriceCount + SimpleRound(order_rec.Fields.Item("CurrentCount") * Val(price_rec.Fields.Item("CurrentPrice")), 0)
            End If
            
            Body = Body & "<td>" & SimpleRound(order_rec.Fields.Item("WinningCount"), 4) & "</td>"
            If price_rec.EOF Then
                Body = Body & "<td>0</td>"
            Else
                Body = Body & "<td>" & SimpleRound(order_rec.Fields.Item("WinningCount") * Val(price_rec.Fields.Item("WinningPrice")), 0) & "</td>"
                PriceCount = PriceCount - SimpleRound(order_rec.Fields.Item("WinningCount") * Val(price_rec.Fields.Item("WinningPrice")), 0)
            End If
            
            Body = Body & "<td>" & SimpleRound(Val(order_rec.Fields.Item("AddMoney")), 0) & "</td>"
            If Not IsNull(order_rec.Fields.Item("AddMoney")) Then PriceCount = PriceCount + SimpleRound(Val(order_rec.Fields.Item("AddMoney")), 0)
            
            Body = Body & "<td>" & SimpleRound(Val(order_rec.Fields.Item("BonusMoney")), 0) & "</td>"
            '不計退水金額 'If Not IsNull(order_rec.Fields.Item("BonusMoney")) Then PriceCount = PriceCount - simpleround(val(order_rec.Fields.Item("BonusMoney")),0)
            
            Body = Body & "<td>" & order_rec.Fields.Item("Note") & "</td>"
              

            Body = Body & "</tr>"
            Print #1, Body
            
            order_rec.MoveNext
        Loop
        
        'show price count
        Body = "<tr><td>應收</td><td>" & SimpleRound(PriceCount, 0) & "</td></tr>"
        Print #1, Body
        
        Print #1, "</table>"
    Close #1
    
    order_rec.Close
End Sub

Sub CustomProductWeekReportDetail(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer
    Dim CurrentCount As Double, CurrentPrice As Long, WinningCount As Double, WinningPrice As Long
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    Dim CData() As String
    Dim PData() As String
    Dim beginv As Integer
    Dim endv As Integer
    Dim OrderDate As String, ProductID As String, ProductName As String
    Dim PriceCount As Long
    
       
    'search order
    CData = Split(cmbCName.Text, " ")
    PData = Split(cmbPName.Text, " ")
    If Val(PData(0)) >= 100 Then
        beginv = Mid(PData(0), 2, 1)
        endv = beginv & "9"
        beginv = beginv & "0"
        SQL = "select * from [order] where cint(PID)>=" & beginv & " and cint(PID)<=" & endv & " and CID='" & CData(0) & "' and (CurrentDate>='" & Format(DateTime.DateAdd("d", -6, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd") & "' and CurrentDate<='" & txtCurrentDate.Text & "') order by CurrentDate;"
    Else
        SQL = "select * from [order] where PID='" & PData(0) & "' and CID='" & CData(0) & "' and (CurrentDate>='" & Format(DateTime.DateAdd("d", -6, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd") & "' and CurrentDate<='" & txtCurrentDate.Text & "') order by CurrentDate;"
    End If
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, order_rec)
  
    
    Open TargetPath For Output As #1
        
        Print #1, "<table>"
    
        'show report datetime
        Body = "<tr><td>日期</td><td colspan=10>" & DateTime.DateAdd("d", -6, txtCurrentDate.Text) & "至" & txtCurrentDate.Text & "</td></tr>"
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
            
            Body = Body & "<td>" & SimpleRound(order_rec.Fields.Item("CurrentCount"), 4) & "</td>"
            If price_rec.EOF Then
                Body = Body & "<td>0</td>"
            Else
                Body = Body & "<td>" & SimpleRound(order_rec.Fields.Item("CurrentCount") * Val(price_rec.Fields.Item("CurrentPrice")), 0) & "</td>"
                PriceCount = PriceCount + SimpleRound(order_rec.Fields.Item("CurrentCount") * Val(price_rec.Fields.Item("CurrentPrice")), 0)
            End If
            
            Body = Body & "<td>" & SimpleRound(order_rec.Fields.Item("WinningCount"), 4) & "</td>"
            If price_rec.EOF Then
                Body = Body & "<td>0</td>"
            Else
                Body = Body & "<td>" & SimpleRound(order_rec.Fields.Item("WinningCount") * Val(price_rec.Fields.Item("WinningPrice")), 0) & "</td>"
                PriceCount = PriceCount - SimpleRound(order_rec.Fields.Item("WinningCount") * Val(price_rec.Fields.Item("WinningPrice")), 0)
            End If
            
            Body = Body & "<td>" & SimpleRound(Val(order_rec.Fields.Item("AddMoney")), 0) & "</td>"
            If Not IsNull(order_rec.Fields.Item("AddMoney")) Then PriceCount = PriceCount + SimpleRound(Val(order_rec.Fields.Item("AddMoney")), 0)
            
            Body = Body & "<td>" & SimpleRound(Val(order_rec.Fields.Item("BonusMoney")), 0) & "</td>"
            '不計退水金額 'If Not IsNull(order_rec.Fields.Item("BonusMoney")) Then PriceCount = PriceCount - simpleround(val(order_rec.Fields.Item("BonusMoney")),0)
            
            Body = Body & "<td>" & order_rec.Fields.Item("Note") & "</td>"
              

            Body = Body & "</tr>"
            Print #1, Body
            
            order_rec.MoveNext
        Loop
        
        'show price count
        Body = "<tr><td>應收</td><td>" & SimpleRound(PriceCount, 0) & "</td></tr>"
        Print #1, Body
        
        Print #1, "</table>"
    Close #1
    
    order_rec.Close
End Sub


Sub CustomProductDayReport(ByVal TargetPath As String)
    Dim exap As Excel.Application
    Dim exwb As Excel.Workbook
    Dim exsh As Excel.Worksheet
    Dim row As Integer, col As Integer
    Set exap = New Excel.Application
    Set exwb = exap.Workbooks.Add
    Set exsh = exwb.Sheets.Item(1)
    exsh.Columns.ColumnWidth = 16
    
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer
    Dim CurrentCount As Double, CurrentPrice As Long, WinningCount As Double, WinningPrice As Long
    Dim AddMoney As Long, BonusMoney As Long, Note As String
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    Dim CData() As String
    Dim PData() As String
    Dim beginv As Integer
    Dim endv As Integer
    Dim OrderDate As String, ProductID As String, ProductName As String
    Dim PriceCount As Long
    Dim GroupFlag As Integer, GroupMax As Integer
    Dim Current(100) As Double, Winning(100) As Double
    
        
    'search order and product
    CData = Split(cmbCName.Text, " ")
    PData = Split(cmbPName.Text, " ")
    If Val(PData(0)) >= 100 Then
        beginv = Mid(PData(0), 2, 1)
        endv = beginv & "9"
        beginv = beginv & "0"
               
        SQL = "select * from product where cint(PID)>=" & beginv & " and cint(PID)<=" & endv & " order by CLng(PID);"
        Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
        
        'get the highest group, and setting for row count
        SQL = "select top 1 MAX(Group) AS HighestGroup from [order] where cint(PID)>=" & beginv & " and cint(PID)<=" & endv & " and CurrentDate='" & txtCurrentDate.Text & "';"
        Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
    Else
        SQL = "select * from product where PID='" & PData(0) & "' order by CLng(PID);"
        Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
        
        'get the highest group, and setting for row count
        SQL = "select top 1 MAX(Group) AS HighestGroup from [order] where PID='" & PData(0) & "' and CurrentDate='" & txtCurrentDate.Text & "';"
        Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, rec1)
    End If
    GroupMax = Val(rec1.Fields.Item("HighestGroup"))
       
    
    
        'show report datetime
        row = 1: col = 1: exsh.Cells(row, col) = "日期"
        col = col + 1: exsh.Cells(row, col) = txtCurrentDate.Text
        col = 11:   Call exsh.Range(Numeric2CharEN(2) & row, Numeric2CharEN(col) & row).Merge
        row = row + 1
        
        'show custom name and product name
        col = 1: exsh.Cells(row, col) = "客戶名稱"
        col = col + 1: exsh.Cells(row, col) = CData(1)
        col = 11:   Call exsh.Range(Numeric2CharEN(2) & row, Numeric2CharEN(col) & row).Merge
        row = row + 1
        col = 1: exsh.Cells(row, col) = "產品名稱"
        col = col + 1: exsh.Cells(row, col) = PData(1)
        col = 11:   Call exsh.Range(Numeric2CharEN(2) & row, Numeric2CharEN(col) & row).Merge
        row = row + 1
        col = 1: exsh.Cells(row, col) = "明細"
        For i = 0 To GroupMax
            col = col + 1: exsh.Cells(row, col) = (i + 1)
        Next
        col = col + 1: exsh.Cells(row, col) = "數量總計"
        col = col + 1: exsh.Cells(row, col) = "金額總計"
        row = row + 1
        
        
        'list choice product
        AddMoney = 0
        BonusMoney = 0
        Note = ""
        PriceCount = 0
        Do Until product_rec.EOF
            'get product ID
            ProductID = product_rec.Fields.Item("PID")
            
            
            'search order
            SQL = "select * from [order] where PID='" & ProductID & "' and CID='" & CData(0) & "' and CurrentDate='" & txtCurrentDate.Text & "' order by CurrentDate,CLng(PID),Group;"
            Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, order_rec)
            
            
            
            'show every order with current date
            GroupFlag = 0
            CurrentCount = 0
            CurrentPrice = 0
            WinningCount = 0
            WinningPrice = 0
            'order_rec.MoveFirst
            Do Until order_rec.EOF
                'search price
                OrderDate = order_rec.Fields.Item("CurrentDate")
                selectFields = "CurrentPrice,WinningPrice,Upset"
                SQL = "select * from price where PID='" & ProductID & "' and CID='" & CData(0) & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
                
                
                
                'mark custom name
                CurrentCount = CurrentCount + Val(order_rec.Fields.Item("CurrentCount"))
                If Not price_rec.EOF Then
                    CurrentPrice = CurrentPrice + (Val(order_rec.Fields.Item("CurrentCount")) * Val(price_rec.Fields.Item("CurrentPrice")))
                End If
                
                WinningCount = WinningCount + Val(order_rec.Fields.Item("WinningCount"))
                If Not price_rec.EOF Then
                    WinningPrice = WinningPrice + (Val(order_rec.Fields.Item("WinningCount")) * Val(price_rec.Fields.Item("WinningPrice")))
                End If

                
                AddMoney = AddMoney + Val(order_rec.Fields.Item("AddMoney"))
                '不計退水金額 'BonusMoney = BonusMoney + Val(order_rec.Fields.Item("BonusMoney"))
                
                Note = Note & order_rec.Fields.Item("Note")
                
                
                'Debug.Print order_rec.Fields.Item("pid") & " ";
                'Debug.Print order_rec.Fields.Item("currentcount");
                'Debug.Print order_rec.Fields.Item("Group");
                'Debug.Print GroupFlag
                
                
                'save count
                If Val(order_rec.Fields.Item("Group")) <> GroupFlag Then
                    For i = GroupFlag To Val(order_rec.Fields.Item("Group")) - 1
                        Current(i) = 0
                        Winning(i) = 0
                    Next
                    GroupFlag = Val(order_rec.Fields.Item("Group"))
                End If
                Current(GroupFlag) = order_rec.Fields.Item("CurrentCount")
                Winning(GroupFlag) = order_rec.Fields.Item("WinningCount")
                GroupFlag = GroupFlag + 1
                
                            
                'move to next record
                order_rec.MoveNext
            Loop
            
            
            'fill data to last
            For i = GroupFlag To GroupMax
                Current(i) = 0
                Winning(i) = 0
            Next
            
            
            'write to xml
            col = 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName")
            For i = 0 To GroupMax
                col = col + 1: exsh.Cells(row, col) = Current(i)
            Next
            col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentCount, 4)
            col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentPrice, 0)
            row = row + 1
            PriceCount = PriceCount + CurrentPrice
            
            col = 1: exsh.Cells(row, col) = product_rec.Fields.Item("PName") & "中"
            For i = 0 To GroupMax
                col = col + 1: exsh.Cells(row, col) = Winning(i)
            Next
            col = col + 1: exsh.Cells(row, col) = SimpleRound(WinningCount, 4)
            col = col + 1: exsh.Cells(row, col) = SimpleRound(WinningPrice, 0)
            row = row + 1
            PriceCount = PriceCount - WinningPrice
            
            product_rec.MoveNext
        Loop
        
        
        col = 1: exsh.Cells(row, col) = "漲價"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(AddMoney, 0)
        row = row + 1
        PriceCount = PriceCount + AddMoney
        
        
        'col = 1: exsh.Cells(row, col) = "退水"
        'col = col + 1: exsh.Cells(row, col) = SimpleRound( BonusMoney, 0)
        'row = row + 1
        '不計退水金額 'PriceCount = PriceCount - BonusMoney
        
        
        'show price count
        col = 1: exsh.Cells(row, col) = "應收"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(PriceCount, 0)
        row = row + 1
        
        
    Call exwb.SaveAs(TargetPath)
    Call exap.Quit
    
    
    order_rec.Close
    product_rec.Close
End Sub

Sub CustomProductWeekReport(ByVal TargetPath As String)

End Sub


Sub CustomProductWeekTransaction(ByVal TargetPath As String)
    Dim exap As Excel.Application
    Dim exwb As Excel.Workbook
    Dim exsh As Excel.Worksheet
    Dim row As Integer, col As Integer
    Set exap = New Excel.Application
    Set exwb = exap.Workbooks.Add
    Set exsh = exwb.Sheets.Item(1)
    exsh.Columns.ColumnWidth = 16
    
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer
    Dim CurrentCount As Double, CurrentPrice As Long, WinningCount As Double, WinningPrice As Long
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    Dim CData() As String
    Dim PData() As String
    Dim beginv As Integer
    Dim endv As Integer
    Dim OrderDate As String, ProductID As String, ProductName As String
    Dim PriceCount As Long
    Dim OldOrderDate As String, CurrentPriceCount As Long, WinningPriceCount As Long, AddMoney As Long, BonusMoney As Long
    Dim DayDiff As Integer
       
       
    'search order
    CData = Split(cmbCName.Text, " ")
    PData = Split(cmbPName.Text, " ")
    If Val(PData(0)) >= 100 Then
        beginv = Mid(PData(0), 2, 1)
        endv = beginv & "9"
        beginv = beginv & "0"
        SQL = "select * from [order] where cint(PID)>=" & beginv & " and cint(PID)<=" & endv & " and CID='" & CData(0) & "' and (CurrentDate>='" & Format(DateTime.DateAdd("d", -6, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd") & "' and CurrentDate<='" & txtCurrentDate.Text & "') order by CurrentDate;"
    Else
        SQL = "select * from [order] where PID='" & PData(0) & "' and CID='" & CData(0) & "' and (CurrentDate>='" & Format(DateTime.DateAdd("d", -6, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd") & "' and CurrentDate<='" & txtCurrentDate.Text & "') order by CurrentDate;"
    End If
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, order_rec)
  
    
        'show report datetime
        row = 1: col = 1: exsh.Cells(row, col) = "日期"
        col = col + 1: exsh.Cells(row, col) = DateTime.DateAdd("d", -6, txtCurrentDate.Text) & "至" & txtCurrentDate.Text
        col = 11:   Call exsh.Range(Numeric2CharEN(2) & row, Numeric2CharEN(col) & row).Merge
        row = row + 1
        
        'show custom name and product name
        col = 1: exsh.Cells(row, col) = "客戶名稱"
        col = col + 1: exsh.Cells(row, col) = CData(1)
        col = 11:   Call exsh.Range(Numeric2CharEN(2) & row, Numeric2CharEN(col) & row).Merge
        row = row + 1
        col = 1: exsh.Cells(row, col) = "產品名稱"
        col = col + 1: exsh.Cells(row, col) = PData(1)
        col = 11:   Call exsh.Range(Numeric2CharEN(2) & row, Numeric2CharEN(col) & row).Merge
        row = row + 1
        col = 1: exsh.Cells(row, col) = "交易日期"
        col = col + 1: exsh.Cells(row, col) = "交易金額"
        col = col + 1: exsh.Cells(row, col) = "交中獎金額"
        col = col + 1: exsh.Cells(row, col) = "漲價"
        col = col + 1: exsh.Cells(row, col) = "退水金額"
        col = col + 1: exsh.Cells(row, col) = "小計"
        row = row + 1
        
        
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
                    DayDiff = DateDiff("d", Format(DateTime.DateAdd("d", -6, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd"), OrderDate)
                    For i = 0 To DayDiff - 1
                        col = 1: exsh.Cells(row, col) = DateTime.DateAdd("d", i, Format(DateTime.DateAdd("d", -6, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd"))
                        col = col + 1: exsh.Cells(row, col) = 0
                        col = col + 1: exsh.Cells(row, col) = 0
                        col = col + 1: exsh.Cells(row, col) = 0
                        col = col + 1: exsh.Cells(row, col) = 0
                        col = col + 1: exsh.Cells(row, col) = 0
                        row = row + 1
                    Next
                End If
                
                'save data
                If Not price_rec.EOF Then
                    CurrentPriceCount = CurrentPriceCount + SimpleRound(order_rec.Fields.Item("CurrentCount") * Val(price_rec.Fields.Item("CurrentPrice")), 0)
                    WinningPriceCount = WinningPriceCount + SimpleRound(order_rec.Fields.Item("WinningCount") * Val(price_rec.Fields.Item("WinningPrice")), 0)
                End If
                If Not IsNull(order_rec.Fields.Item("AddMoney")) Then AddMoney = AddMoney + SimpleRound(order_rec.Fields.Item("AddMoney"), 0)
                If Not IsNull(order_rec.Fields.Item("BonusMoney")) Then BonusMoney = BonusMoney + SimpleRound(order_rec.Fields.Item("BonusMoney"), 0)
            Else
                'mark custom name
                col = 1: exsh.Cells(row, col) = OldOrderDate
                col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentPriceCount, 0)
                col = col + 1: exsh.Cells(row, col) = SimpleRound(WinningPriceCount, 0)
                col = col + 1: exsh.Cells(row, col) = SimpleRound(AddMoney, 0)
                col = col + 1: exsh.Cells(row, col) = SimpleRound(BonusMoney, 0)
                col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentPriceCount - WinningPriceCount + AddMoney + BonusMoney, 0)
                row = row + 1
                PriceCount = PriceCount + CurrentPriceCount - WinningPriceCount + AddMoney + BonusMoney
                
                'clear variable data
                CurrentPriceCount = 0
                WinningPriceCount = 0
                AddMoney = 0
                BonusMoney = 0
                
                'fill data per day
                DayDiff = DateDiff("d", OldOrderDate, OrderDate)
                For i = 1 To DayDiff - 1
                    col = 1: exsh.Cells(row, col) = DateTime.DateAdd("d", i, OldOrderDate)
                    col = col + 1: exsh.Cells(row, col) = 0
                    col = col + 1: exsh.Cells(row, col) = 0
                    col = col + 1: exsh.Cells(row, col) = 0
                    col = col + 1: exsh.Cells(row, col) = 0
                    col = col + 1: exsh.Cells(row, col) = 0
                    row = row + 1
                Next
            End If
            OldOrderDate = OrderDate
            
            
            order_rec.MoveNext
        Loop
        
        
        
        'show the last data and fill data to lastday as last
        'if everything is empty, then setting a startup date for full data
        If OrderDate = "" Then OrderDate = Format(DateTime.DateAdd("d", -6, Format(txtCurrentDate.Text, "yyyy/MM/dd")))
        If OldOrderDate = "" Then OldOrderDate = OrderDate
        
        'mark custom name
        col = 1: exsh.Cells(row, col) = OldOrderDate
        col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentPriceCount, 0)
        col = col + 1: exsh.Cells(row, col) = SimpleRound(WinningPriceCount, 0)
        col = col + 1: exsh.Cells(row, col) = SimpleRound(AddMoney, 0)
        col = col + 1: exsh.Cells(row, col) = SimpleRound(BonusMoney, 0)
        col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentPriceCount - WinningPriceCount + AddMoney + BonusMoney, 0)
        row = row + 1
        PriceCount = PriceCount + CurrentPriceCount - WinningPriceCount + AddMoney + BonusMoney
        DayDiff = DateDiff("d", OrderDate, txtCurrentDate.Text)
        For i = 1 To DayDiff - 1
            col = 1: exsh.Cells(row, col) = DateTime.DateAdd("d", i, OrderDate)
            col = col + 1: exsh.Cells(row, col) = 0
            col = col + 1: exsh.Cells(row, col) = 0
            col = col + 1: exsh.Cells(row, col) = 0
            col = col + 1: exsh.Cells(row, col) = 0
            col = col + 1: exsh.Cells(row, col) = 0
            row = row + 1
        Next
        
        
        'show price count
        col = 1: exsh.Cells(row, col) = "總計"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(PriceCount, 0)
        
        
    Call exwb.SaveAs(TargetPath)
    Call exap.Quit
    
    
    order_rec.Close
End Sub

Sub CustomWeekReport(ByVal TargetPath As String)
    Dim exap As Excel.Application
    Dim exwb As Excel.Workbook
    Dim exsh As Excel.Worksheet
    Dim row As Integer, col As Integer
    Set exap = New Excel.Application
    Set exwb = exap.Workbooks.Add
    Set exsh = exwb.Sheets.Item(1)
    exsh.Columns.ColumnWidth = 16
    
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer
    Dim CurrentCount As Double, CurrentPrice As Long, WinningCount As Double, WinningPrice As Long
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    Dim CData() As String
    Dim PData() As String
    Dim beginv As Integer
    Dim endv As Integer
    Dim OrderDate As String, ProductID As String, ProductName As String
    Dim PriceCount As Long
    
    
    'search order
    CData = Split(cmbCName.Text, " ")
    SQL = "select * from [order] where CID='" & CData(0) & "' and (CurrentDate>='" & Format(DateTime.DateAdd("d", -6, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd") & "' and CurrentDate<='" & txtCurrentDate.Text & "') order by CurrentDate;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, order_rec)
  
    
        'show report datetime
        row = 1: col = 1: exsh.Cells(row, col) = "日期"
        col = col + 1: exsh.Cells(row, col) = DateTime.DateAdd("d", -6, txtCurrentDate.Text) & "至" & txtCurrentDate.Text
        col = 11:   Call exsh.Range(Numeric2CharEN(2) & row, Numeric2CharEN(col) & row).Merge
        row = row + 1
        
        'show custom name and product name
        col = 1: exsh.Cells(row, col) = "客戶名稱"
        col = col + 1: exsh.Cells(row, col) = CData(1)
        col = 11:   Call exsh.Range(Numeric2CharEN(2) & row, Numeric2CharEN(col) & row).Merge
        row = row + 1
        col = 1: exsh.Cells(row, col) = "交易流水號"
        col = col + 1: exsh.Cells(row, col) = "交易日期"
        col = col + 1: exsh.Cells(row, col) = "產品名稱"
        col = col + 1: exsh.Cells(row, col) = "交易數量"
        col = col + 1: exsh.Cells(row, col) = "交易金額"
        col = col + 1: exsh.Cells(row, col) = "中獎數量"
        col = col + 1: exsh.Cells(row, col) = "中獎金額"
        col = col + 1: exsh.Cells(row, col) = "漲價"
        col = col + 1: exsh.Cells(row, col) = "退水金額"
        col = col + 1: exsh.Cells(row, col) = "備註"
        row = row + 1
        
        
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
            col = 1: exsh.Cells(row, col) = order_rec.Fields.Item("SwiftCode")
            col = col + 1: exsh.Cells(row, col) = OrderDate
            col = col + 1: exsh.Cells(row, col) = ProductName
            
            col = col + 1: exsh.Cells(row, col) = SimpleRound(order_rec.Fields.Item("CurrentCount"), 4)
            If price_rec.EOF Then
                col = col + 1: exsh.Cells(row, col) = 0
            Else
                col = col + 1: exsh.Cells(row, col) = SimpleRound(order_rec.Fields.Item("CurrentCount") * Val(price_rec.Fields.Item("CurrentPrice")), 0)
                PriceCount = PriceCount + SimpleRound(order_rec.Fields.Item("CurrentCount") * Val(price_rec.Fields.Item("CurrentPrice")), 0)
            End If
            
            col = col + 1: exsh.Cells(row, col) = SimpleRound(order_rec.Fields.Item("WinningCount"), 4)
            If price_rec.EOF Then
                col = col + 1: exsh.Cells(row, col) = 0
            Else
                col = col + 1: exsh.Cells(row, col) = SimpleRound(order_rec.Fields.Item("WinningCount") * Val(price_rec.Fields.Item("WinningPrice")), 0)
                PriceCount = PriceCount - SimpleRound(order_rec.Fields.Item("WinningCount") * Val(price_rec.Fields.Item("WinningPrice")), 0)
            End If
            
            col = col + 1: exsh.Cells(row, col) = SimpleRound(Val(order_rec.Fields.Item("AddMoney")), 0)
            If Not IsNull(order_rec.Fields.Item("AddMoney")) Then PriceCount = PriceCount + SimpleRound(Val(order_rec.Fields.Item("AddMoney")), 0)
            
            col = col + 1: exsh.Cells(row, col) = SimpleRound(Val(order_rec.Fields.Item("BonusMoney")), 0)
            '不計退水金額 'If Not IsNull(order_rec.Fields.Item("BonusMoney")) Then PriceCount = PriceCount - simpleround(val(order_rec.Fields.Item("BonusMoney")),0)
            
            col = col + 1: exsh.Cells(row, col) = order_rec.Fields.Item("Note")
              

            row = row + 1
            
            
            order_rec.MoveNext
        Loop
        
        
        'show price count
        col = 1: exsh.Cells(row, col) = "應收"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(PriceCount, 0)
        
    Call exwb.SaveAs(TargetPath)
    Call exap.Quit
    
    
    order_rec.Close
End Sub

Sub CustomMonthReport(ByVal TargetPath As String)
    Dim exap As Excel.Application
    Dim exwb As Excel.Workbook
    Dim exsh As Excel.Worksheet
    Dim row As Integer, col As Integer
    Set exap = New Excel.Application
    Set exwb = exap.Workbooks.Add
    Set exsh = exwb.Sheets.Item(1)
    exsh.Columns.ColumnWidth = 16
    
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer
    Dim CurrentCount As Double, CurrentPrice As Long, WinningCount As Double, WinningPrice As Long
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    Dim CData() As String
    Dim PData() As String
    Dim beginv As Integer
    Dim endv As Integer
    Dim OrderDate As String, ProductID As String, ProductName As String
    Dim PriceCount As Long
    
    
    'search order
    CData = Split(cmbCName.Text, " ")
    SQL = "select * from [order] where CID='" & CData(0) & "' and (CurrentDate>='" & Format(txtCurrentDate.Text, "yyyy/MM/") & "01' and CurrentDate<='" & Format(txtCurrentDate.Text, "yyyy/MM/") & Date_Is_28_29_30_31(txtCurrentDate.Text) & "') order by CurrentDate;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, order_rec)
  
          
        'show report datetime
        row = 1: col = 1: exsh.Cells(row, col) = "日期"
        col = col + 1: exsh.Cells(row, col) = Format(txtCurrentDate.Text, "yyyy/MM")
        col = 11:   Call exsh.Range(Numeric2CharEN(2) & row, Numeric2CharEN(col) & row).Merge
        row = row + 1
        
        'show custom name and product name
        col = 1: exsh.Cells(row, col) = "客戶名稱"
        col = col + 1: exsh.Cells(row, col) = CData(1)
        col = 11:   Call exsh.Range(Numeric2CharEN(2) & row, Numeric2CharEN(col) & row).Merge
        row = row + 1
        col = 1: exsh.Cells(row, col) = "交易流水號"
        col = col + 1: exsh.Cells(row, col) = "交易日期"
        col = col + 1: exsh.Cells(row, col) = "產品名稱"
        col = col + 1: exsh.Cells(row, col) = "交易數量"
        col = col + 1: exsh.Cells(row, col) = "交易金額"
        col = col + 1: exsh.Cells(row, col) = "中獎數量"
        col = col + 1: exsh.Cells(row, col) = "中獎金額"
        col = col + 1: exsh.Cells(row, col) = "漲價"
        col = col + 1: exsh.Cells(row, col) = "退水金額"
        col = col + 1: exsh.Cells(row, col) = "備註"
        row = row + 1
        
        
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
            col = 1: exsh.Cells(row, col) = order_rec.Fields.Item("SwiftCode")
            col = col + 1: exsh.Cells(row, col) = OrderDate
            col = col + 1: exsh.Cells(row, col) = ProductName
            
            col = col + 1: exsh.Cells(row, col) = SimpleRound(order_rec.Fields.Item("CurrentCount"), 4)
            If price_rec.EOF Then
                col = col + 1: exsh.Cells(row, col) = 0
            Else
                col = col + 1: exsh.Cells(row, col) = SimpleRound(order_rec.Fields.Item("CurrentCount") * Val(price_rec.Fields.Item("CurrentPrice")), 0)
                PriceCount = PriceCount + SimpleRound(order_rec.Fields.Item("CurrentCount") * Val(price_rec.Fields.Item("CurrentPrice")), 0)
            End If
            
            col = col + 1: exsh.Cells(row, col) = SimpleRound(order_rec.Fields.Item("WinningCount"), 4)
            If price_rec.EOF Then
                col = col + 1: exsh.Cells(row, col) = 0
            Else
                col = col + 1: exsh.Cells(row, col) = SimpleRound(order_rec.Fields.Item("WinningCount") * Val(price_rec.Fields.Item("WinningPrice")), 0)
                PriceCount = PriceCount - SimpleRound(order_rec.Fields.Item("WinningCount") * Val(price_rec.Fields.Item("WinningPrice")), 0)
            End If
            
            col = col + 1: exsh.Cells(row, col) = SimpleRound(Val(order_rec.Fields.Item("AddMoney")), 0)
            If Not IsNull(order_rec.Fields.Item("AddMoney")) Then PriceCount = PriceCount + SimpleRound(Val(order_rec.Fields.Item("AddMoney")), 0)
            
            col = col + 1: exsh.Cells(row, col) = SimpleRound(Val(order_rec.Fields.Item("BonusMoney")), 0)
            '不計退水金額 'If Not IsNull(order_rec.Fields.Item("BonusMoney")) Then PriceCount = PriceCount - simpleround(val(order_rec.Fields.Item("BonusMoney")),0)
            
            col = col + 1: exsh.Cells(row, col) = order_rec.Fields.Item("Note")
              

            row = row + 1
            
            
            order_rec.MoveNext
        Loop
        
        'show price count
        col = 1: exsh.Cells(row, col) = "應收"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(PriceCount, 0)
        
    Call exwb.SaveAs(TargetPath)
    Call exap.Quit
    
    
    
    order_rec.Close
End Sub

Sub CustomYearReport(ByVal TargetPath As String)
    Dim exap As Excel.Application
    Dim exwb As Excel.Workbook
    Dim exsh As Excel.Worksheet
    Dim row As Integer, col As Integer
    Set exap = New Excel.Application
    Set exwb = exap.Workbooks.Add
    Set exsh = exwb.Sheets.Item(1)
    exsh.Columns.ColumnWidth = 16
    
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer
    Dim CurrentCount As Double, CurrentPrice As Long, WinningCount As Double, WinningPrice As Long
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    Dim CData() As String
    Dim PData() As String
    Dim beginv As Integer
    Dim endv As Integer
    Dim OrderDate As String, ProductID As String, ProductName As String
    Dim PriceCount As Long
    
    
    'search order
    CData = Split(cmbCName.Text, " ")
    SQL = "select * from [order] where CID='" & CData(0) & "' and (CurrentDate>='" & Format(txtCurrentDate.Text, "yyyy/") & "01/01' and CurrentDate<='" & Format(txtCurrentDate.Text, "yyyy/") & "12/31') order by CurrentDate;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, order_rec)
  
    
        'show report datetime
        row = 1: col = 1: exsh.Cells(row, col) = "日期"
        col = col + 1: exsh.Cells(row, col) = Format(txtCurrentDate.Text, "yyyy")
        col = 11:   Call exsh.Range(Numeric2CharEN(2) & row, Numeric2CharEN(col) & row).Merge
        row = row + 1
        
        'show custom name and product name
        col = 1: exsh.Cells(row, col) = "客戶名稱"
        col = col + 1: exsh.Cells(row, col) = CData(1)
        col = 11:   Call exsh.Range(Numeric2CharEN(2) & row, Numeric2CharEN(col) & row).Merge
        row = row + 1
        col = 1: exsh.Cells(row, col) = "交易流水號"
        col = col + 1: exsh.Cells(row, col) = "交易日期"
        col = col + 1: exsh.Cells(row, col) = "產品名稱"
        col = col + 1: exsh.Cells(row, col) = "交易數量"
        col = col + 1: exsh.Cells(row, col) = "交易金額"
        col = col + 1: exsh.Cells(row, col) = "中獎數量"
        col = col + 1: exsh.Cells(row, col) = "中獎金額"
        col = col + 1: exsh.Cells(row, col) = "漲價"
        col = col + 1: exsh.Cells(row, col) = "退水金額"
        col = col + 1: exsh.Cells(row, col) = "備註"
        row = row + 1
        
        
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
            col = 1: exsh.Cells(row, col) = order_rec.Fields.Item("SwiftCode")
            col = col + 1: exsh.Cells(row, col) = OrderDate
            col = col + 1: exsh.Cells(row, col) = ProductName
            
            col = col + 1: exsh.Cells(row, col) = SimpleRound(order_rec.Fields.Item("CurrentCount"), 4)
            If price_rec.EOF Then
                col = col + 1: exsh.Cells(row, col) = 0
            Else
                col = col + 1: exsh.Cells(row, col) = SimpleRound(order_rec.Fields.Item("CurrentCount") * Val(price_rec.Fields.Item("CurrentPrice")), 0)
                PriceCount = PriceCount + SimpleRound(order_rec.Fields.Item("CurrentCount") * Val(price_rec.Fields.Item("CurrentPrice")), 0)
            End If
            
            col = col + 1: exsh.Cells(row, col) = SimpleRound(order_rec.Fields.Item("WinningCount"), 4)
            If price_rec.EOF Then
                col = col + 1: exsh.Cells(row, col) = 0
            Else
                col = col + 1: exsh.Cells(row, col) = SimpleRound(order_rec.Fields.Item("WinningCount") * Val(price_rec.Fields.Item("WinningPrice")), 0)
                PriceCount = PriceCount - SimpleRound(order_rec.Fields.Item("WinningCount") * Val(price_rec.Fields.Item("WinningPrice")), 0)
            End If
            
            col = col + 1: exsh.Cells(row, col) = SimpleRound(Val(order_rec.Fields.Item("AddMoney")), 0)
            If Not IsNull(order_rec.Fields.Item("AddMoney")) Then PriceCount = PriceCount + SimpleRound(Val(order_rec.Fields.Item("AddMoney")), 0)
            
            col = col + 1: exsh.Cells(row, col) = SimpleRound(Val(order_rec.Fields.Item("BonusMoney")), 0)
            '不計退水金額 'If Not IsNull(order_rec.Fields.Item("BonusMoney")) Then PriceCount = PriceCount - simpleround(val(order_rec.Fields.Item("BonusMoney")),0)
            
            col = col + 1: exsh.Cells(row, col) = order_rec.Fields.Item("Note")
              

            row = row + 1
            
            
            order_rec.MoveNext
        Loop
        
        
        'show price count
        col = 1: exsh.Cells(row, col) = "應收"
        col = col + 1: exsh.Cells(row, col) = SimpleRound(PriceCount, 0)
        
    Call exwb.SaveAs(TargetPath)
    Call exap.Quit
    
    
    order_rec.Close
End Sub

Sub CustomWeekTransaction(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer
    Dim CurrentCount As Double, CurrentPrice As Long, WinningCount As Double, WinningPrice As Long
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    Dim CData() As String
    Dim PData() As String
    Dim beginv As Integer
    Dim endv As Integer
    Dim OrderDate As String, ProductID As String, ProductName As String
    Dim PriceCount As Long
    Dim OldOrderDate As String, CurrentPriceCount As Long, WinningPriceCount As Long, AddMoney As Long, BonusMoney As Long
    Dim DayDiff As Integer
       
       
    'search order
    CData = Split(cmbCName.Text, " ")
    SQL = "select * from [order] where CID='" & CData(0) & "' and (CurrentDate>='" & Format(DateTime.DateAdd("d", -6, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd") & "' and CurrentDate<='" & txtCurrentDate.Text & "') order by CurrentDate;"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, order_rec)
  
    
    Open TargetPath For Output As #1
        
        Print #1, "<table>"
    
        'show report datetime
        Body = "<tr><td>日期</td><td colspan=10>" & DateTime.DateAdd("d", -6, txtCurrentDate.Text) & "至" & txtCurrentDate.Text & "</td></tr>"
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
                    DayDiff = DateDiff("d", Format(DateTime.DateAdd("d", -6, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd"), OrderDate)
                    For i = 0 To DayDiff - 1
                        Body = "<tr>"
                        Body = Body & "<td>" & DateTime.DateAdd("d", i, Format(DateTime.DateAdd("d", -6, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd")) & "</td>"
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
                    CurrentPriceCount = CurrentPriceCount + SimpleRound(order_rec.Fields.Item("CurrentCount") * Val(price_rec.Fields.Item("CurrentPrice")), 0)
                    WinningPriceCount = WinningPriceCount + SimpleRound(order_rec.Fields.Item("WinningCount") * Val(price_rec.Fields.Item("WinningPrice")), 0)
                End If
                If Not IsNull(order_rec.Fields.Item("AddMoney")) Then AddMoney = AddMoney + SimpleRound(order_rec.Fields.Item("AddMoney"), 0)
                If Not IsNull(order_rec.Fields.Item("BonusMoney")) Then BonusMoney = BonusMoney + SimpleRound(order_rec.Fields.Item("BonusMoney"), 0)
            Else
                'mark custom name
                Body = "<tr>"
                Body = Body & "<td>" & OldOrderDate & "</td>"
                Body = Body & "<td>" & SimpleRound(CurrentPriceCount, 0) & "</td>"
                Body = Body & "<td>" & SimpleRound(WinningPriceCount, 0) & "</td>"
                Body = Body & "<td>" & SimpleRound(AddMoney, 0) & "</td>"
                Body = Body & "<td>" & SimpleRound(BonusMoney, 0) & "</td>"
                Body = Body & "<td>" & SimpleRound(CurrentPriceCount + WinningPriceCount + AddMoney + BonusMoney, 0) & "</td>"
                PriceCount = PriceCount + SimpleRound(CurrentPriceCount + WinningPriceCount + AddMoney + BonusMoney, 0)
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
        If OrderDate = "" Then OrderDate = Format(DateTime.DateAdd("d", -6, Format(txtCurrentDate.Text, "yyyy/MM/dd")))
        If OldOrderDate = "" Then OldOrderDate = OrderDate
        
        'mark custom name
        Body = "<tr>"
        Body = Body & "<td>" & OldOrderDate & "</td>"
        Body = Body & "<td>" & SimpleRound(CurrentPriceCount, 0) & "</td>"
        Body = Body & "<td>" & SimpleRound(WinningPriceCount, 0) & "</td>"
        Body = Body & "<td>" & SimpleRound(AddMoney, 0) & "</td>"
        Body = Body & "<td>" & SimpleRound(BonusMoney, 0) & "</td>"
        Body = Body & "<td>" & SimpleRound(CurrentPriceCount + WinningPriceCount + AddMoney + BonusMoney, 0) & "</td>"
        PriceCount = PriceCount + SimpleRound(CurrentPriceCount + WinningPriceCount + AddMoney + BonusMoney, 0)
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
        Body = "<tr><td>總計</td><td>" & SimpleRound(PriceCount, 0) & "</td></tr>"
        Print #1, Body
        
        Print #1, "</table>"
    Close #1
    
    order_rec.Close
End Sub


Sub CustomMonthTransaction(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer
    Dim CurrentCount As Double, CurrentPrice As Long, WinningCount As Double, WinningPrice As Long
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    Dim CData() As String
    Dim PData() As String
    Dim beginv As Integer
    Dim endv As Integer
    Dim OrderDate As String, ProductID As String, ProductName As String
    Dim PriceCount As Long
    Dim OldOrderDate As String, CurrentPriceCount As Long, WinningPriceCount As Long, AddMoney As Long, BonusMoney As Long
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
                    CurrentPriceCount = CurrentPriceCount + SimpleRound(order_rec.Fields.Item("CurrentCount") * Val(price_rec.Fields.Item("CurrentPrice")), 0)
                    WinningPriceCount = WinningPriceCount + SimpleRound(order_rec.Fields.Item("WinningCount") * Val(price_rec.Fields.Item("WinningPrice")), 0)
                End If
                If Not IsNull(order_rec.Fields.Item("AddMoney")) Then AddMoney = AddMoney + SimpleRound(order_rec.Fields.Item("AddMoney"), 0)
                If Not IsNull(order_rec.Fields.Item("BonusMoney")) Then BonusMoney = BonusMoney + SimpleRound(order_rec.Fields.Item("BonusMoney"), 0)
            Else
                'mark custom name
                Body = "<tr>"
                Body = Body & "<td>" & OldOrderDate & "</td>"
                Body = Body & "<td>" & SimpleRound(CurrentPriceCount, 0) & "</td>"
                Body = Body & "<td>" & SimpleRound(WinningPriceCount, 0) & "</td>"
                Body = Body & "<td>" & SimpleRound(AddMoney, 0) & "</td>"
                Body = Body & "<td>" & SimpleRound(BonusMoney, 0) & "</td>"
                Body = Body & "<td>" & SimpleRound(CurrentPriceCount + WinningPriceCount + AddMoney + BonusMoney, 0) & "</td>"
                PriceCount = PriceCount + SimpleRound(CurrentPriceCount + WinningPriceCount + AddMoney + BonusMoney, 0)
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
        Body = Body & "<td>" & SimpleRound(CurrentPriceCount, 0) & "</td>"
        Body = Body & "<td>" & SimpleRound(WinningPriceCount, 0) & "</td>"
        Body = Body & "<td>" & SimpleRound(AddMoney, 0) & "</td>"
        Body = Body & "<td>" & SimpleRound(BonusMoney, 0) & "</td>"
        Body = Body & "<td>" & SimpleRound(CurrentPriceCount + WinningPriceCount + AddMoney + BonusMoney, 0) & "</td>"
        PriceCount = PriceCount + SimpleRound(CurrentPriceCount + WinningPriceCount + AddMoney + BonusMoney, 0)
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
        Body = "<tr><td>總計</td><td>" & SimpleRound(PriceCount, 0) & "</td></tr>"
        Print #1, Body
        
        Print #1, "</table>"
    Close #1
    
    order_rec.Close
End Sub


Sub CustomYearTransaction(ByVal TargetPath As String)
    Dim selectFields As String
    Dim Body As String, i As Integer, PIDArray(1024) As String, Count As Integer
    Dim CurrentCount As Double, CurrentPrice As Long, WinningCount As Double, WinningPrice As Long
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    Dim CData() As String
    Dim PData() As String
    Dim beginv As Integer
    Dim endv As Integer
    Dim OrderDate As String, ProductID As String, ProductName As String
    Dim PriceCount As Long
    Dim OldOrderDate As String, CurrentPriceCount As Long, WinningPriceCount As Long, AddMoney As Long, BonusMoney As Long
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
                    CurrentPriceCount = CurrentPriceCount + SimpleRound(order_rec.Fields.Item("CurrentCount") * Val(price_rec.Fields.Item("CurrentPrice")), 0)
                    WinningPriceCount = WinningPriceCount + SimpleRound(order_rec.Fields.Item("WinningCount") * Val(price_rec.Fields.Item("WinningPrice")), 0)
                End If
                If Not IsNull(order_rec.Fields.Item("AddMoney")) Then AddMoney = AddMoney + SimpleRound(order_rec.Fields.Item("AddMoney"), 0)
                If Not IsNull(order_rec.Fields.Item("BonusMoney")) Then BonusMoney = BonusMoney + SimpleRound(order_rec.Fields.Item("BonusMoney"), 0)
            Else
                'mark custom name
                Body = "<tr>"
                Body = Body & "<td>" & OldOrderDate & "</td>"
                Body = Body & "<td>" & SimpleRound(CurrentPriceCount, 0) & "</td>"
                Body = Body & "<td>" & SimpleRound(WinningPriceCount, 0) & "</td>"
                Body = Body & "<td>" & SimpleRound(AddMoney, 0) & "</td>"
                Body = Body & "<td>" & SimpleRound(BonusMoney, 0) & "</td>"
                Body = Body & "<td>" & SimpleRound(CurrentPriceCount + WinningPriceCount + AddMoney + BonusMoney, 0) & "</td>"
                PriceCount = PriceCount + SimpleRound(CurrentPriceCount + WinningPriceCount + AddMoney + BonusMoney, 0)
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
        Body = Body & "<td>" & SimpleRound(CurrentPriceCount, 0) & "</td>"
        Body = Body & "<td>" & SimpleRound(WinningPriceCount, 0) & "</td>"
        Body = Body & "<td>" & SimpleRound(AddMoney, 0) & "</td>"
        Body = Body & "<td>" & SimpleRound(BonusMoney, 0) & "</td>"
        Body = Body & "<td>" & SimpleRound(CurrentPriceCount + WinningPriceCount + AddMoney + BonusMoney, 0) & "</td>"
        PriceCount = PriceCount + SimpleRound(CurrentPriceCount + WinningPriceCount + AddMoney + BonusMoney, 0)
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
        Body = "<tr><td>總計</td><td>" & SimpleRound(PriceCount, 0) & "</td></tr>"
        Print #1, Body
        
        Print #1, "</table>"
    Close #1
    
    order_rec.Close
End Sub



Sub ProductWeekTransaction(ByVal TargetPath As String)
    Dim exap As Excel.Application
    Dim exwb As Excel.Workbook
    Dim exsh As Excel.Worksheet
    Dim row As Integer, col As Integer
    Set exap = New Excel.Application
    Set exwb = exap.Workbooks.Add
    Set exsh = exwb.Sheets.Item(1)
    exsh.Columns.ColumnWidth = 16
    
    Const Day As Integer = 6
    
    Dim selectFields As String
    Dim Body As String, i As Integer, n As Integer
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    Dim PData() As String
    Dim beginv As Integer
    Dim endv As Integer
    Dim OrderDate As String, ProductID As String, ProductName As String
    Dim PriceCount As Long
    Dim OldOrderDate As String, CurrentPriceCount(Day) As Double, WinningPriceCount(Day) As Double, AddMoney As Long, BonusMoney As Long
    Dim DayDiff As Integer
    Dim SaveOrderDate(Day) As String
    
    
    PData = Split(cmbPName.Text, " ")
    
    
    'list all custom
    SQL = "select * from custom order by CLng(CID);"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
  
    
        
        'show report datetime
        row = 1: col = 1: exsh.Cells(row, col) = "日期"
        col = col + 1: exsh.Cells(row, col) = DateTime.DateAdd("d", -6, txtCurrentDate.Text) & "至" & txtCurrentDate.Text
        col = 11:   Call exsh.Range(Numeric2CharEN(2) & row, Numeric2CharEN(col) & row).Merge
        row = row + 1
        
        'show product name
        col = 1: exsh.Cells(row, col) = "客戶名稱"
        For i = 6 To 0 Step -1
            col = col + 1: exsh.Cells(row, col) = DateTime.DateAdd("d", -i, txtCurrentDate.Text)
        Next
        col = col + 1: exsh.Cells(row, col) = "小計"
        row = row + 1
        
       
        'list all custom again
        custom_rec.MoveFirst
        Do Until custom_rec.EOF
            'search order
            If Val(PData(0)) >= 100 Then
                beginv = Mid(PData(0), 2, 1)
                endv = beginv & "9"
                beginv = beginv & "0"
                SQL = "select * from [order] where cint(PID)>=" & beginv & " and cint(PID)<=" & endv & " and CID='" & custom_rec.Fields.Item("CID") & "' and (CurrentDate>='" & Format(DateTime.DateAdd("d", -6, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd") & "' and CurrentDate<='" & txtCurrentDate.Text & "') order by CurrentDate;"
            Else
                SQL = "select * from [order] where PID='" & PData(0) & "' and CID='" & custom_rec.Fields.Item("CID") & "' and (CurrentDate>='" & Format(DateTime.DateAdd("d", -6, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd") & "' and CurrentDate<='" & txtCurrentDate.Text & "') order by CurrentDate;" ''SQL = "select * from [order] where CID='" & custom_rec.Fields.Item("CID") & "' and (CurrentDate>='" & Format(DateTime.DateAdd("d", -6, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd") & "' and CurrentDate<='" & txtCurrentDate.Text & "') order by CurrentDate;"
            End If
            Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, order_rec)
            
            
            'show every order with current date
            n = -1
            OldOrderDate = ""
            'order_rec.MoveFirst
            Do Until order_rec.EOF
            
                'search price
                OrderDate = order_rec.Fields.Item("CurrentDate")
                ProductID = order_rec.Fields.Item("PID")
                selectFields = "CurrentPrice,WinningPrice,Upset"
                SQL = "select * from price where PID='" & ProductID & "' and CID='" & custom_rec.Fields.Item("CID") & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc,CLng(SwiftCode);"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
                
                
                If OldOrderDate = OrderDate Then
                    'save data per same data
                    If Not price_rec.EOF Then
                        CurrentPriceCount(n) = CurrentPriceCount(n) + (order_rec.Fields.Item("CurrentCount") * Val(price_rec.Fields.Item("CurrentPrice")))
                        'WinningPriceCount(n) = WinningPriceCount(n) + (order_rec.Fields.Item("WinningCount") * Val(price_rec.Fields.Item("WinningPrice")))
                    End If
                Else
                    n = n + 1
                    SaveOrderDate(n) = OrderDate
                    OldOrderDate = OrderDate
                    
                    'save data per new day
                    If Not price_rec.EOF Then
                        CurrentPriceCount(n) = (order_rec.Fields.Item("CurrentCount") * Val(price_rec.Fields.Item("CurrentPrice")))
                        'WinningPriceCount(i) =  (order_rec.Fields.Item("WinningCount") * Val(price_rec.Fields.Item("WinningPrice")))
                    Else
                        CurrentPriceCount(n) = 0
                        'WinningPriceCount(i) =  0
                    End If
                End If
                

                order_rec.MoveNext
            Loop
            
          
            'list all custom name
            n = 0
            PriceCount = 0
            col = 1: exsh.Cells(row, col) = custom_rec.Fields.Item("CName")
            For i = 6 To 0 Step -1
                If SaveOrderDate(n) = Format(DateTime.DateAdd("d", -i, txtCurrentDate.Text), "yyyy/MM/dd") Then
                    PriceCount = PriceCount + CurrentPriceCount(n)
                    
                    col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentPriceCount(n), 0)
                    'col = col + 1: exsh.Cells(row, col) = SimpleRound(WinningPriceCount(n) ,0)
                    
                    CurrentPriceCount(n) = 0
                    'WinningPriceCount(n) =  0
                    
                    n = n + 1
                Else
                    col = col + 1: exsh.Cells(row, col) = 0
                    'col = col + 1: exsh.Cells(row, col) = 0
                End If
            Next
            col = col + 1: exsh.Cells(row, col) = "=sum(" & Numeric2CharEN(2) & row & ":" & Numeric2CharEN(8) & row & ")" 'SimpleRound(PriceCount, 0)
            row = row + 1
            
            
            custom_rec.MoveNext
        Loop
        
        
    Call exwb.SaveAs(TargetPath)
    Call exap.Quit
    
    
    order_rec.Close
End Sub


Sub WeekTransaction(ByVal TargetPath As String)
    Dim exap As Excel.Application
    Dim exwb As Excel.Workbook
    Dim exsh As Excel.Worksheet
    Dim row As Integer, col As Integer
    Set exap = New Excel.Application
    Set exwb = exap.Workbooks.Add
    Set exsh = exwb.Sheets.Item(1)
    exsh.Columns.ColumnWidth = 16
    
    Const Day As Integer = 6
    
    Dim selectFields As String
    Dim Body As String, i As Integer, n As Integer
    Dim SQL As String
    Dim product_rec As New adoDB.Recordset, price_rec As New adoDB.Recordset, custom_rec As New adoDB.Recordset, order_rec As New adoDB.Recordset
    Dim rec1 As New adoDB.Recordset
    Dim beginv As Integer
    Dim endv As Integer
    Dim OrderDate As String, ProductID As String, ProductName As String
    Dim PriceCount As Long
    Dim OldOrderDate As String, CurrentPriceCount(Day) As Double, WinningPriceCount(Day) As Double, AddMoney As Long, BonusMoney As Long
    Dim DayDiff As Integer
    Dim SaveOrderDate(Day) As String
       
    'list all custom
    SQL = "select * from custom order by CLng(CID);"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, custom_rec)
  
    
       
        'show report datetime
        row = 1: col = 1: exsh.Cells(row, col) = "日期"
        col = col + 1: exsh.Cells(row, col) = DateTime.DateAdd("d", -6, txtCurrentDate.Text) & "至" & txtCurrentDate.Text
        col = 11:   Call exsh.Range(Numeric2CharEN(2) & row, Numeric2CharEN(col) & row).Merge
        row = row + 1
        
        'show product name
        col = 1: exsh.Cells(row, col) = "客戶名稱"
        For i = 6 To 0 Step -1
            col = col + 1: exsh.Cells(row, col) = DateTime.DateAdd("d", -i, txtCurrentDate.Text)
        Next
        col = col + 1: exsh.Cells(row, col) = "前帳"
        col = col + 1: exsh.Cells(row, col) = "小計"
        row = row + 1
        
        
       
        'list all custom again
        custom_rec.MoveFirst
        Do Until custom_rec.EOF
            'search order
            SQL = "select * from [order] where CID='" & custom_rec.Fields.Item("CID") & "' and (CurrentDate>='" & Format(DateTime.DateAdd("d", -6, Format(txtCurrentDate.Text, "yyyy/MM/dd")), "yyyy/MM/dd") & "' and CurrentDate<='" & txtCurrentDate.Text & "') order by CurrentDate;"
            Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, order_rec)
            
            
            'show every order with current date
            n = -1
            OldOrderDate = ""
            'order_rec.MoveFirst
            Do Until order_rec.EOF
            
                'search price
                OrderDate = order_rec.Fields.Item("CurrentDate")
                ProductID = order_rec.Fields.Item("PID")
                selectFields = "CurrentPrice,WinningPrice,Upset"
                SQL = "select * from price where PID='" & ProductID & "' and CID='" & custom_rec.Fields.Item("CID") & "' and CurrentDate<='" & OrderDate & "' order by CurrentDate desc;"
                Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, price_rec)
                
                
                If OldOrderDate = OrderDate Then
                    'save data per same data
                    If Not price_rec.EOF Then
                        CurrentPriceCount(n) = CurrentPriceCount(n) + (order_rec.Fields.Item("CurrentCount") * Val(price_rec.Fields.Item("CurrentPrice")))
                        'WinningPriceCount(n) = WinningPriceCount(n) + (order_rec.Fields.Item("WinningCount") * Val(price_rec.Fields.Item("WinningPrice")))
                    End If
                Else
                    n = n + 1
                    SaveOrderDate(n) = OrderDate
                    OldOrderDate = OrderDate
                    
                    'save data per new day
                    If Not price_rec.EOF Then
                        CurrentPriceCount(n) = (order_rec.Fields.Item("CurrentCount") * Val(price_rec.Fields.Item("CurrentPrice")))
                        'WinningPriceCount(i) =  (order_rec.Fields.Item("WinningCount") * Val(price_rec.Fields.Item("WinningPrice")))
                    Else
                        CurrentPriceCount(n) = 0
                        'WinningPriceCount(i) =  0
                    End If
                End If


                order_rec.MoveNext
            Loop
            
          
            'list all custom name
            n = 0
            PriceCount = 0
            col = 1: exsh.Cells(row, col) = custom_rec.Fields.Item("CName")
            For i = 6 To 0 Step -1
                If SaveOrderDate(n) = Format(DateTime.DateAdd("d", -i, txtCurrentDate.Text), "yyyy/MM/dd") Then
                    PriceCount = PriceCount + CurrentPriceCount(n)
                    
                    col = col + 1: exsh.Cells(row, col) = SimpleRound(CurrentPriceCount(n), 0)
                    'col = col + 1: exsh.Cells(row, col) = SimpleRound(WinningPriceCount(n) ,0)
                    
                    CurrentPriceCount(n) = 0
                    'WinningPriceCount(n) =  0
                    
                    n = n + 1
                Else
                    col = col + 1: exsh.Cells(row, col) = 0
                    'col = col + 1: exsh.Cells(row, col) = 0
                End If
            Next
            col = col + 1
            col = col + 1: exsh.Cells(row, col) = "=sum(" & Numeric2CharEN(2) & row & ":" & Numeric2CharEN(9) & row & ")" 'SimpleRound(PriceCount, 0)
            row = row + 1
            
            
            
            custom_rec.MoveNext
        Loop
        
        
    Call exwb.SaveAs(TargetPath)
    Call exap.Quit
    
    
    order_rec.Close
End Sub




Private Sub cmdConfirm_Click()
'On Error GoTo errout:
    If txtCurrentDate.Text = "" Then
        MsgBox "請先選擇要列印的時間！"
    ElseIf (basVariable.Parameter = "CustomProductDayReport" Or basVariable.Parameter = "CustomProductWeekReport") And cmbCName.Text = "" And cmbPName.Text = "" Then
        MsgBox "尚未選擇客戶或產品！"
    ElseIf (basVariable.Parameter = "CustomWeekReport" Or basVariable.Parameter = "CustomMonthReport" Or basVariable.Parameter = "CustomYearReport") And cmbCName.Text = "" Then
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
            TargetPath = TargetPath & Format(DateTime.DateAdd("d", -6, txtCurrentDate.Text), "yyyyMMdd") & "至" & Format(txtCurrentDate.Text, "yyyyMMdd") & "_週報表.xls"
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
            TargetPath = TargetPath & Format(DateTime.DateAdd("d", -6, txtCurrentDate.Text), "yyyyMMdd") & "至" & Format(txtCurrentDate.Text, "yyyyMMdd") & "_週總帳.xls"
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
            TargetPath = TargetPath & Format(DateTime.DateAdd("d", -6, txtCurrentDate.Text), "yyyyMMdd") & "至" & Format(txtCurrentDate.Text, "yyyyMMdd") & "_4K週報表.xls"
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
            TargetPath = TargetPath & Format(DateTime.DateAdd("d", -6, txtCurrentDate.Text), "yyyyMMdd") & "至" & Format(txtCurrentDate.Text, "yyyyMMdd") & "_4K週總帳.xls"
            Call FourKWeekAccount(TargetPath)
        Case "FourKMonthAccount"
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyyMM") & "_4K月總帳.xls"
            Call FourKMonthAccount(TargetPath)
        Case "FourKYearAccount"
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyy") & "_4K年總帳.xls"
            Call FourKYearAccount(TargetPath)
        Case "CustomProductDayReportDetail"
            CData = Split(cmbCName.Text, " ")
            PData = Split(cmbPName.Text, " ")
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyyMMdd") & "_" & CData(1) & "_" & PData(1) & "_分客別、分產品日明細表.xls"
            Call CustomProductDayReportDetail(TargetPath)
        Case "CustomProductWeekReportDetail"
            CData = Split(cmbCName.Text, " ")
            PData = Split(cmbPName.Text, " ")
            TargetPath = TargetPath & Format(DateTime.DateAdd("d", -6, txtCurrentDate.Text), "yyyyMMdd") & "至" & Format(txtCurrentDate.Text, "yyyyMMdd") & "_" & CData(1) & "_" & PData(1) & "_分客別、分產品週明細表.xls"
            Call CustomProductWeekReportDetail(TargetPath)
        Case "CustomProductDayReport"
            CData = Split(cmbCName.Text, " ")
            PData = Split(cmbPName.Text, " ")
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyyMMdd") & "_" & CData(1) & "_" & PData(1) & "_分客別、分產品日報表.xls"
            Call CustomProductDayReport(TargetPath)
        Case "CustomProductWeekReport"
            CData = Split(cmbCName.Text, " ")
            PData = Split(cmbPName.Text, " ")
            TargetPath = TargetPath & Format(DateTime.DateAdd("d", -6, txtCurrentDate.Text), "yyyyMMdd") & "至" & Format(txtCurrentDate.Text, "yyyyMMdd") & "_" & CData(1) & "_" & PData(1) & "_分客別、分產品週報表.xls"
            Call CustomProductWeekReport(TargetPath)
        Case "CustomProductWeekTransaction"
            CData = Split(cmbCName.Text, " ")
            PData = Split(cmbPName.Text, " ")
            TargetPath = TargetPath & Format(DateTime.DateAdd("d", -6, txtCurrentDate.Text), "yyyyMMdd") & "至" & Format(txtCurrentDate.Text, "yyyyMMdd") & "_" & CData(1) & "_" & PData(1) & "_分客別、分產品週交易金額表.xls"
            Call CustomProductWeekTransaction(TargetPath)
        Case "CustomWeekReport"
            CData = Split(cmbCName.Text, " ")
            PData = Split(cmbPName.Text, " ")
            TargetPath = TargetPath & Format(DateTime.DateAdd("d", -6, txtCurrentDate.Text), "yyyyMMdd") & "至" & Format(txtCurrentDate.Text, "yyyyMMdd") & "_" & CData(1) & "_分客別、不分產品週報表.xls"
            Call CustomWeekReport(TargetPath)
        Case "CustomMonthReport"
            CData = Split(cmbCName.Text, " ")
            PData = Split(cmbPName.Text, " ")
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyyMM") & "_" & CData(1) & "_分客別、不分產品月報表.xls"
            Call CustomMonthReport(TargetPath)
        Case "CustomYearReport"
            CData = Split(cmbCName.Text, " ")
            PData = Split(cmbPName.Text, " ")
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyy") & "_" & CData(1) & "_分客別、不分產品年報表.xls"
            Call CustomYearReport(TargetPath)
        Case "CustomWeekTransaction"
            CData = Split(cmbCName.Text, " ")
            PData = Split(cmbPName.Text, " ")
            TargetPath = TargetPath & Format(DateTime.DateAdd("d", -6, txtCurrentDate.Text), "yyyyMMdd") & "至" & Format(txtCurrentDate.Text, "yyyyMMdd") & "_" & CData(1) & "_分客別、不分產品週交易金額表.xls"
            Call CustomWeekTransaction(TargetPath)
        Case "CustomMonthTransaction"
            CData = Split(cmbCName.Text, " ")
            PData = Split(cmbPName.Text, " ")
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyyMM") & "_" & CData(1) & "_分客別、不分產品月交易金額表.xls"
            Call CustomMonthTransaction(TargetPath)
        Case "CustomYearTransaction"
            CData = Split(cmbCName.Text, " ")
            PData = Split(cmbPName.Text, " ")
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyy") & "_" & CData(1) & "_分客別、不分產品年交易金額表.xls"
            Call CustomYearTransaction(TargetPath)
        Case "ProductWeekTransaction"
            TargetPath = TargetPath & Format(DateTime.DateAdd("d", -6, txtCurrentDate.Text), "yyyyMMdd") & "至" & Format(txtCurrentDate.Text, "yyyyMMdd") & "_不分客別、分產品週交易金額表.xls"
            Call ProductWeekTransaction(TargetPath)
        Case "WeekTransaction"
            TargetPath = TargetPath & Format(DateTime.DateAdd("d", -6, txtCurrentDate.Text), "yyyyMMdd") & "至" & Format(txtCurrentDate.Text, "yyyyMMdd") & "_不分客別、不分產品週交易金額表.xls"
            Call WeekTransaction(TargetPath)
        Case "MonthTransaction"
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyyMM") & "_不分客別、不分產品月交易金額表.xls"
            'Call MonthTransaction(TargetPath)
        Case "YearTransaction"
            TargetPath = TargetPath & Format(txtCurrentDate.Text, "yyyy") & "_不分客別、不分產品年交易金額表.xls"
            'Call YearTransaction(TargetPath)
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
    Case "CustomProductDayReportDetail"
        Label1(0).Caption = "分客別、分產品日明細表列印"
        lblEntry(0).Visible = True
        lblEntry(2).Visible = True
        cmbCName.Visible = True
        cmbPName.Visible = True
        Call ComboBox_LoadFrom_DataBase_ByFile(cmbCName, "CID,CName", "custom", "", "", "")
        Call ComboBox_LoadFrom_DataBase_ByFile(cmbPName, "PID,PName", "product", "", "", "")
        
        Call cmbPName.AddItem("100 539_全")
        Call cmbPName.AddItem("110 港號_全")
        Call cmbPName.AddItem("120 大樂透_全")
    Case "CustomProductWeekReportDetail"
        Label1(0).Caption = "分客別、分產品週明細表列印"
        lblEntry(0).Visible = True
        lblEntry(2).Visible = True
        cmbCName.Visible = True
        cmbPName.Visible = True
        Call ComboBox_LoadFrom_DataBase_ByFile(cmbCName, "CID,CName", "custom", "", "", "")
        Call ComboBox_LoadFrom_DataBase_ByFile(cmbPName, "PID,PName", "product", "", "", "")
        
        Call cmbPName.AddItem("100 539_全")
        Call cmbPName.AddItem("110 港號_全")
        Call cmbPName.AddItem("120 大樂透_全")
    Case "CustomProductDayReport"
        Label1(0).Caption = "分客別、分產品日報表列印"
        lblEntry(0).Visible = True
        lblEntry(2).Visible = True
        cmbCName.Visible = True
        cmbPName.Visible = True
        Call ComboBox_LoadFrom_DataBase_ByFile(cmbCName, "CID,CName", "custom", "", "", "")
        Call ComboBox_LoadFrom_DataBase_ByFile(cmbPName, "PID,PName", "product", "", "", "")
        
        Call cmbPName.AddItem("100 539_全")
        Call cmbPName.AddItem("110 港號_全")
        Call cmbPName.AddItem("120 大樂透_全")
    Case "CustomProductWeekReport"
        Label1(0).Caption = "分客別、分產品週報表列印"
        lblEntry(0).Visible = True
        lblEntry(2).Visible = True
        cmbCName.Visible = True
        cmbPName.Visible = True
        Call ComboBox_LoadFrom_DataBase_ByFile(cmbCName, "CID,CName", "custom", "", "", "")
        Call ComboBox_LoadFrom_DataBase_ByFile(cmbPName, "PID,PName", "product", "", "", "")
        
        Call cmbPName.AddItem("100 539_全")
        Call cmbPName.AddItem("110 港號_全")
        Call cmbPName.AddItem("120 大樂透_全")
    Case "CustomProductWeekTransaction"
        Label1(0).Caption = "分客別、分產品週交易金額表列印"
        lblEntry(0).Visible = True
        lblEntry(2).Visible = True
        cmbCName.Visible = True
        cmbPName.Visible = True
        Call ComboBox_LoadFrom_DataBase_ByFile(cmbCName, "CID,CName", "custom", "", "", "")
        Call ComboBox_LoadFrom_DataBase_ByFile(cmbPName, "PID,PName", "product", "", "", "")
        
        Call cmbPName.AddItem("100 539_全")
        Call cmbPName.AddItem("110 港號_全")
        Call cmbPName.AddItem("120 大樂透_全")
    Case "CustomWeekReport"
        Label1(0).Caption = "分客別、不分產品週報表列印"
        lblEntry(0).Visible = True
        cmbCName.Visible = True
        Call ComboBox_LoadFrom_DataBase_ByFile(cmbCName, "CID,CName", "custom", "", "", "")
    Case "CustomMonthReport"
        Label1(0).Caption = "分客別、不分產品月報表列印"
        lblEntry(0).Visible = True
        cmbCName.Visible = True
        Call ComboBox_LoadFrom_DataBase_ByFile(cmbCName, "CID,CName", "custom", "", "", "")
    Case "CustomYearReport"
        Label1(0).Caption = "分客別、不分產品年報表列印"
        lblEntry(0).Visible = True
        cmbCName.Visible = True
        Call ComboBox_LoadFrom_DataBase_ByFile(cmbCName, "CID,CName", "custom", "", "", "")
    Case "CustomWeekTransaction"
        Label1(0).Caption = "分客別、不分產品週交易金額表列印"
        lblEntry(0).Visible = True
        cmbCName.Visible = True
        Call ComboBox_LoadFrom_DataBase_ByFile(cmbCName, "CID,CName", "custom", "", "", "")
    Case "CustomMonthTransaction"
        Label1(0).Caption = "分客別、不分產品月交易金額表列印"
        lblEntry(0).Visible = True
        cmbCName.Visible = True
        Call ComboBox_LoadFrom_DataBase_ByFile(cmbCName, "CID,CName", "custom", "", "", "")
    Case "CustomYearTransaction"
        Label1(0).Caption = "分客別、不分產品年交易金額表列印"
        lblEntry(0).Visible = True
        cmbCName.Visible = True
        Call ComboBox_LoadFrom_DataBase_ByFile(cmbCName, "CID,CName", "custom", "", "", "")
    Case "ProductWeekTransaction"
        Label1(0).Caption = "不分客別、分產品週交易金額表列印"
        lblEntry(2).Visible = True
        cmbPName.Visible = True
        Call ComboBox_LoadFrom_DataBase_ByFile(cmbPName, "PID,PName", "product", "", "", "")
        
        Call cmbPName.AddItem("100 539_全")
        Call cmbPName.AddItem("110 港號_全")
        Call cmbPName.AddItem("120 大樂透_全")
    Case "WeekTransaction"
        Label1(0).Caption = "不分客別、不分產品週交易金額表列印"
    Case "MonthTransaction"
        Label1(0).Caption = "不分客別、不分產品月交易金額表列印"
    Case "YearTransaction"
        Label1(0).Caption = "不分客別、不分產品年交易金額表列印"
    End Select
    
    
    dtpCurrentDate.Value = Format(DateTime.Now, "yyyy/MM/dd")
    txtCurrentDate.Text = Format(DateTime.Now, "yyyy/MM/dd")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmProve.Show
    Unload Me
End Sub
