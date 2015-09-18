VERSION 5.00
Begin VB.UserControl UserControl1 
   BackColor       =   &H00FF0000&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox Picture1 
      Height          =   3015
      Left            =   0
      ScaleHeight     =   2955
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   0
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   3120
      Width           =   4815
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
