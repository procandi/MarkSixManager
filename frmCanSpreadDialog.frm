VERSION 5.00
Begin VB.Form frmCanSpreadDialog 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "套餐範本名稱"
   ClientHeight    =   945
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5850
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   15.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.TextBox Text1 
      Height          =   435
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   4095
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "確定"
      Height          =   615
      Left            =   4440
      TabIndex        =   0
      Top             =   150
      Width           =   1215
   End
End
Attribute VB_Name = "frmCanSpreadDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Activate()
    Text1.SetFocus
End Sub

Private Sub OKButton_Click()
    If Trim(Text1.Text) <> "" Then
        frmCanSpread.xSpreadCanName = Trim(Text1.Text)
    End If
    Unload Me
End Sub
