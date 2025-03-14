VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ";"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   19065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   19065
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSenha 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   6885
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "Digite"
      Top             =   3660
      Width           =   4095
   End
   Begin VB.TextBox txtUsuario 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   0
      Left            =   6855
      TabIndex        =   1
      Text            =   "Digite o seu usuário"
      Top             =   2475
      Width           =   4095
   End
   Begin VB.Label lblSenha 
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6885
      TabIndex        =   3
      Top             =   3240
      Width           =   930
   End
   Begin VB.Label lblUsuario 
      Caption         =   "Usuário"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6825
      TabIndex        =   0
      Top             =   2025
      Width           =   930
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
