VERSION 5.00
Begin VB.Form frmConfirm 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "Report Confirm"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmConfirm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4680
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.OptionButton optStatus 
      BackColor       =   &H00808080&
      Caption         =   "���i������"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   2
      Left            =   840
      TabIndex        =   6
      Top             =   1680
      Width           =   2895
   End
   Begin VB.OptionButton optStatus 
      BackColor       =   &H00808080&
      Caption         =   "���i�w�����A�Ȥ��o�G"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   1
      Left            =   840
      TabIndex        =   5
      Top             =   1200
      Width           =   2895
   End
   Begin VB.CommandButton cmdConfirm 
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
      Index           =   1
      Left            =   2400
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   3
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
      Index           =   0
      Left            =   120
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   2
      Top             =   2520
      Width           =   2175
   End
   Begin VB.OptionButton optStatus 
      BackColor       =   &H00808080&
      Caption         =   "���i�w�����A�����o�G"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   0
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Value           =   -1  'True
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  '�m�����
      BackColor       =   &H00808080&
      BackStyle       =   0  '�z��
      Caption         =   "�ˬd���i�s�ɽT�{"
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
      TabIndex        =   4
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
