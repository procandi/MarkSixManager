VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmOrder 
   Caption         =   "產品購買明細"
   ClientHeight    =   10560
   ClientLeft      =   1965
   ClientTop       =   3420
   ClientWidth     =   15105
   Icon            =   "frmOrder.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   10560
   ScaleWidth      =   15105
   Begin Threed.SSPanel pnlFilter 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      _Version        =   65536
      _ExtentX        =   26485
      _ExtentY        =   1931
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      Outline         =   -1  'True
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFC0C0&
         Caption         =   "刪除"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9600
         Style           =   1  '圖片外觀
         TabIndex        =   18
         Top             =   600
         Width           =   1215
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
         Left            =   7320
         MaxLength       =   10
         TabIndex        =   14
         Top             =   120
         Width           =   1650
      End
      Begin VB.CommandButton cmdAppend 
         BackColor       =   &H00FFC0C0&
         Caption         =   "交易"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10920
         Style           =   1  '圖片外觀
         TabIndex        =   13
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtWinningCount 
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
         Left            =   7320
         TabIndex        =   11
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtPName 
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
         Left            =   1200
         TabIndex        =   7
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtCurrentCount 
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
         Left            =   4200
         TabIndex        =   6
         Top             =   600
         Width           =   1935
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00FFC0C0&
         Caption         =   "更新清單"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13560
         Style           =   1  '圖片外觀
         TabIndex        =   3
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00FFC0C0&
         Caption         =   "清除條件"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12240
         Style           =   1  '圖片外觀
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&X 關　　閉"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12960
         Style           =   1  '圖片外觀
         TabIndex        =   1
         Top             =   120
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker dtpCurrentDate 
         Height          =   360
         Left            =   7320
         TabIndex        =   15
         Top             =   120
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
         Format          =   11337731
         CurrentDate     =   37058
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
         Index           =   0
         Left            =   6240
         TabIndex        =   16
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "中獎數量"
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
         Left            =   6240
         TabIndex        =   12
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblName 
         Alignment       =   2  '置中對齊
         BackColor       =   &H80000015&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   "王小明"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   5295
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "購買數量"
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
         Index           =   23
         Left            =   3120
         TabIndex        =   5
         Top             =   600
         Width           =   1095
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
         Index           =   7
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   9255
      Left            =   0
      TabIndex        =   9
      Top             =   1320
      Width           =   15015
      _Version        =   65536
      _ExtentX        =   26485
      _ExtentY        =   16325
      _StockProps     =   15
      BackColor       =   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      Outline         =   -1  'True
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   9015
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   15901
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   17
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lblAddCount 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   "預設顯示最近1年資料"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   12480
      TabIndex        =   17
      Top             =   1080
      Width           =   2535
   End
End
Attribute VB_Name = "frmOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type BindData
    SwiftCode As String
    PName As String
    CurrentDate As String
    CurrentCount_Car As Double
    WinningCount_Car As Double
    CurrentCount_2K As Double
    WinningCount_2K As Double
    CurrentCount_3K As Double
    WinningCount_3K As Double
    CurrentCount_4K As Double
    WinningCount_4K As Double
    CurrentCount_Special As Double
    WinningCount_Special As Double
    AddMoney As Double
    BonusMoney As Double
    Note As String
End Type

Dim selectFields As String

Private Sub cmdAppend_Click()
    frmOrderAddNew.Show
    'Me.Hide
End Sub

Private Sub cmdClear_Click()
    txtPName.Text = ""
    txtCurrentDate.Text = ""
    txtCurrentCount.Text = ""
    txtWinningCount.Text = ""
End Sub

Private Sub cmdDelete_Click()
    Dim condition As String, SQL As String
    Dim PID(15) As String
    Dim product_rec As New adoDB.Recordset
    
    SQL = "select * from product';"
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, product_rec)
    Do Until product_rec.EOF
        Select Case product_rec("PName")
        Case "539_車"
            PID(0) = product_rec("PID")
        Case "539_2K"
            PID(1) = product_rec("PID")
        Case "539_3K"
            PID(2) = product_rec("PID")
        Case "539_4K"
            PID(3) = product_rec("PID")
        Case "539_3包"
            PID(4) = product_rec("PID")
        Case "港號_車"
            PID(5) = product_rec("PID")
        Case "港號_2K"
            PID(6) = product_rec("PID")
        Case "港號_3K"
            PID(7) = product_rec("PID")
        Case "港號_4K"
            PID(8) = product_rec("PID")
        Case "港號_特"
            PID(9) = product_rec("PID")
        Case "大樂透_車"
            PID(10) = product_rec("PID")
        Case "大樂透_2K"
            PID(11) = product_rec("PID")
        Case "大樂透_3K"
            PID(12) = product_rec("PID")
        Case "大樂透_4K"
            PID(13) = product_rec("PID")
        Case "大樂透_特"
            PID(14) = product_rec("PID")
        End Select
        product_rec.MoveNext
    Loop
    product_rec.Close
    
    
    Select Case basVariable.SelectPName
    Case "539"
        condition = "'" & PID(0) & "'"
        For i = 1 To 4
            condition = condition & ",'" & PID(i) & "'"
        Next
    Case "港號"
        condition = "'" & PID(5) & "'"
        For i = 6 To 9
            condition = condition & ",'" & PID(i) & "'"
        Next
    Case "大樂透"
        condition = "'" & PID(10) & "'"
        For i = 11 To 14
            condition = condition & ",'" & PID(i) & "'"
        Next
    End Select
    
    SQL = "delete from [order] where CurrentDate='" & basVariable.SelectDate & "' and PID in (" & condition & ")"
    basDataBase.Connection.Execute SQL
    

    cmdDelete.Enabled = False
    
    Call cmdRefresh_Click
End Sub

'add function to refresh database and datagrid
Private Sub cmdRefresh_Click()
    Dim condition As String, SQL As String
    Dim recdata(365) As BindData, i As Integer, n As Integer
    Dim rs As adoDB.Recordset
    Dim rex As RegExp
    
    Set rs = New adoDB.Recordset
    Set rex = New RegExp
    rex.Pattern = "_.*"
    
    condition = ""

    If txtPName.Text <> "" Then
        condition = condition & IIf(condition = "", "", "and ") & "PName like '" & rex.Replace(txtPName.Text, "") & "%' "
    End If
    If txtCurrentDate.Text <> "" Then
        condition = condition & IIf(condition = "", "", "and ") & "CurrentDate='" & txtCurrentDate.Text & "' "
    End If
    If txtCurrentCount.Text <> "" Then
        condition = condition & IIf(condition = "", "", "and ") & "CurrentCount='" & txtCurrentCount.Text & "' "
    End If
    If txtWinningCount.Text <> "" Then
        condition = condition & IIf(condition = "", "", "and ") & "WinningCount='" & txtWinningCount.Text & "' "
    End If

    If condition = "" Then
        SQL = "select " & selectFields & " from [order],product where [order].PID=product.PID and [order].CID='" & basVariable.SelectCID & "' order by CurrentDate desc, product.PName;"
    Else
        SQL = "select " & selectFields & " from [order],product where [order].PID=product.PID and [order].CID='" & basVariable.SelectCID & "' and " & condition & " order by CurrentDate desc, product.PName;"
    End If
    Call basDataBase.OpenRecordset(SQL, basDataBase.Connection, basDataBase.Recordset)
    
       
    
    'get the Grid's Recordset and add a new Record. the limit number between -9999.9999 to 9999.9999
    rs.Fields.Append "SwiftCode", adVarChar, 10
    rs.Fields.Append "CurrentDate", adVarChar, 10
    rs.Fields.Append "PName", adVarChar, 6
    rs.Fields.Append "CurrentCount_Car", adVarChar, 10
    rs.Fields.Append "WinningCount_Car", adVarChar, 10
    rs.Fields.Append "CurrentCount_2K", adVarChar, 10
    rs.Fields.Append "WinningCount_2K", adVarChar, 10
    rs.Fields.Append "CurrentCount_3K", adVarChar, 10
    rs.Fields.Append "WinningCount_3K", adVarChar, 10
    rs.Fields.Append "CurrentCount_4K", adVarChar, 10
    rs.Fields.Append "WinningCount_4K", adVarChar, 10
    rs.Fields.Append "CurrentCount_Special", adVarChar, 10
    rs.Fields.Append "WinningCount_Special", adVarChar, 10
    rs.Fields.Append "AddMoney", adVarChar, 15
    rs.Fields.Append "BonusMoney", adVarChar, 15
    rs.Fields.Append "Note", adVarChar, 150
    rs.Open
    
         
    'add data
    For i = 0 To 365
        recdata(i).CurrentCount_Car = 0
        recdata(i).WinningCount_Car = 0
        recdata(i).CurrentCount_2K = 0
        recdata(i).WinningCount_2K = 0
        recdata(i).CurrentCount_3K = 0
        recdata(i).WinningCount_3K = 0
        recdata(i).CurrentCount_4K = 0
        recdata(i).WinningCount_4K = 0
        recdata(i).CurrentCount_Special = 0
        recdata(i).WinningCount_Special = 0
    Next
    n = 0
    Do Until basDataBase.Recordset.EOF
        If recdata(n).CurrentDate <> basDataBase.Recordset.Fields.Item("CurrentDate") Or rex.Replace(recdata(n).PName, "") <> rex.Replace(basDataBase.Recordset.Fields.Item("PName"), "") Then
            n = n + 1
            If n = 365 Then Exit Do 'data line limit unless one year
            
            recdata(n).SwiftCode = basDataBase.Recordset.Fields.Item("SwiftCode")
            recdata(n).CurrentDate = basDataBase.Recordset.Fields.Item("CurrentDate")
            recdata(n).PName = rex.Replace(basDataBase.Recordset.Fields.Item("PName"), "")
            
            recdata(n).AddMoney = Val(basDataBase.Recordset.Fields.Item("AddMoney"))
            recdata(n).BonusMoney = Val(basDataBase.Recordset.Fields.Item("BonusMoney"))
            recdata(n).Note = basDataBase.Recordset.Fields.Item("Note")
         End If
         
        Select Case basDataBase.Recordset.Fields.Item("PName")
        Case "539_車", "港號_車", "大樂透_車"
            recdata(n).CurrentCount_Car = recdata(n).CurrentCount_Car + Val(basDataBase.Recordset.Fields.Item("CurrentCount"))
            recdata(n).WinningCount_Car = recdata(n).WinningCount_Car + Val(basDataBase.Recordset.Fields.Item("WinningCount"))
        Case "539_2K", "港號_2K", "大樂透_2K"
            recdata(n).CurrentCount_2K = recdata(n).CurrentCount_2K + Val(basDataBase.Recordset.Fields.Item("CurrentCount"))
            recdata(n).WinningCount_2K = recdata(n).WinningCount_2K + Val(basDataBase.Recordset.Fields.Item("WinningCount"))
        Case "539_3K", "港號_3K", "大樂透_3K"
            recdata(n).CurrentCount_3K = recdata(n).WinningCount_3K + Val(basDataBase.Recordset.Fields.Item("CurrentCount"))
            recdata(n).WinningCount_3K = recdata(n).WinningCount_3K + Val(basDataBase.Recordset.Fields.Item("WinningCount"))
        Case "539_4K", "港號_4K", "大樂透_4K"
            recdata(n).CurrentCount_4K = recdata(n).WinningCount_4K + Val(basDataBase.Recordset.Fields.Item("CurrentCount"))
            recdata(n).WinningCount_4K = recdata(n).WinningCount_4K + Val(basDataBase.Recordset.Fields.Item("WinningCount"))
        Case "539_3包", "港號_特", "大樂透_特"
            recdata(n).CurrentCount_Special = recdata(n).WinningCount_Special + Val(basDataBase.Recordset.Fields.Item("CurrentCount"))
            recdata(n).WinningCount_Special = recdata(n).WinningCount_Special + Val(basDataBase.Recordset.Fields.Item("WinningCount"))
        End Select
        
        basDataBase.Recordset.MoveNext
    Loop
    
    For i = 1 To n
        rs.AddNew
        rs.Fields("SwiftCode").Value = recdata(i).SwiftCode
        rs.Fields("CurrentDate").Value = recdata(i).CurrentDate
        rs.Fields("PName").Value = recdata(i).PName
        rs.Fields("CurrentCount_Car").Value = Format(recdata(i).CurrentCount_Car, "0.0000")
        rs.Fields("WinningCount_Car").Value = Format(recdata(i).WinningCount_Car, "0.0000")
        rs.Fields("CurrentCount_2K").Value = Format(recdata(i).CurrentCount_2K, "0.0000")
        rs.Fields("WinningCount_2K").Value = Format(recdata(i).WinningCount_2K, "0.0000")
        rs.Fields("CurrentCount_3K").Value = Format(recdata(i).CurrentCount_3K, "0.0000")
        rs.Fields("WinningCount_3K").Value = Format(recdata(i).WinningCount_3K, "0.0000")
        rs.Fields("CurrentCount_4K").Value = Format(recdata(i).CurrentCount_4K, "0.0000")
        rs.Fields("WinningCount_4K").Value = Format(recdata(i).WinningCount_4K, "0.0000")
        rs.Fields("CurrentCount_Special").Value = Format(recdata(i).CurrentCount_Special, "0.0000")
        rs.Fields("WinningCount_Special").Value = Format(recdata(i).WinningCount_Special, "0.0000")
        rs.Fields("AddMoney").Value = Format(recdata(i).AddMoney, "0.0000")
        rs.Fields("BonusMoney").Value = Format(recdata(i).BonusMoney, "0.0000")
        rs.Fields("Note").Value = recdata(i).Note
    Next
    
    
    'bind recordset to Grid
    Set DataGrid1.DataSource = rs
    
    
    RefreshDataGridHeader
End Sub

Private Sub dtpCurrentDate_CloseUp()
    txtCurrentDate.Text = Format(dtpCurrentDate.Value, "yyyy/MM/dd")
End Sub

'do refresh database and datagrid when form paint
Private Sub Form_Paint()
    Call cmdRefresh_Click
End Sub

Private Sub cmdClose_Click()
    Call Form_Unload(0)
End Sub

'get something system needed when user click datagrid row
Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If basDataBase.Recordset.RecordCount > 0 Then
        If DataGrid1.Columns("交易流水號") <> "" Then
            cmdDelete.Enabled = True
            basVariable.CurrentSwiftCode = DataGrid1.Columns("交易流水號")
            basVariable.SelectPName = DataGrid1.Columns("產品名稱")
            basVariable.SelectDate = DataGrid1.Columns("交易日期")
        End If
        
        If DataGrid1.SelBookmarks.Count <> 0 Then Call DataGrid1.SelBookmarks.Remove(0)
        Call DataGrid1.SelBookmarks.Add(DataGrid1.Bookmark)
    End If
End Sub

'import database and export to datagrid when form load
Private Sub Form_Load()
    cmdDelete.Enabled = False
    
    'Don't allow user to modify data in grid directly.
    With DataGrid1
        .AllowAddNew = False
        .AllowDelete = False
        .AllowUpdate = False
    End With


    lblName(0).Caption = basVariable.SelectCName
    selectFields = "SwiftCode,CID,[order].PID,PName,CurrentDate,CurrentCount,WinningCount,AddMoney,BonusMoney,Note"
    
    
    dtpCurrentDate.Value = Format(DateTime.Now, "yyyy/MM/dd")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmCustom.Show
    Unload Me
End Sub

'a function to batch rename datagrid header
Sub RefreshDataGridHeader()
    DataGrid1.Columns("SwiftCode").Caption = "交易流水號"
    DataGrid1.Columns("CurrentDate").Caption = "交易日期"
    'DataGrid1.Columns("CID").Caption = "客戶編號"
    'DataGrid1.Columns("PID").Caption = "產品編號"
    DataGrid1.Columns("PName").Caption = "產品名稱"
    DataGrid1.Columns("CurrentCount_Car").Caption = "車交易數量"
    DataGrid1.Columns("WinningCount_Car").Caption = "車中獎數量"
    DataGrid1.Columns("CurrentCount_2K").Caption = "2K交易數量"
    DataGrid1.Columns("WinningCount_2K").Caption = "2K中獎數量"
    DataGrid1.Columns("CurrentCount_3K").Caption = "3K交易數量"
    DataGrid1.Columns("WinningCount_3K").Caption = "3K中獎數量"
    DataGrid1.Columns("CurrentCount_4K").Caption = "4K交易數量"
    DataGrid1.Columns("WinningCount_4K").Caption = "4K中獎數量"
    DataGrid1.Columns("CurrentCount_Special").Caption = "3包或特交易數量"
    DataGrid1.Columns("WinningCount_Special").Caption = "3包或特中獎數量"
    DataGrid1.Columns("AddMoney").Caption = "漲價"
    DataGrid1.Columns("BonusMoney").Caption = "退水金額"
    DataGrid1.Columns("Note").Caption = "備註"
End Sub

