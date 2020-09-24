VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4170
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtServer 
      Appearance      =   0  'Flat
      BackColor       =   &H00E9E9E9&
      Height          =   360
      Left            =   1275
      TabIndex        =   0
      Text            =   "localhost"
      Top             =   105
      Width           =   2655
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00E9E9E9&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   3045
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1560
      Width           =   870
   End
   Begin VB.CommandButton cmdLogin 
      BackColor       =   &H00E9E9E9&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   870
   End
   Begin VB.TextBox txtPass 
      Appearance      =   0  'Flat
      BackColor       =   &H00E9E9E9&
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1260
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "admin"
      Top             =   1065
      Width           =   2655
   End
   Begin VB.TextBox txtUser 
      Appearance      =   0  'Flat
      BackColor       =   &H00E9E9E9&
      Height          =   360
      Left            =   1260
      TabIndex        =   1
      Text            =   "admin"
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label lblServer 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Server:"
      Height          =   240
      Left            =   135
      TabIndex        =   7
      Top             =   150
      Width           =   945
   End
   Begin VB.Label lblPass 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Password:"
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   1110
      Width           =   945
   End
   Begin VB.Label lblUser 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Username:"
      Height          =   240
      Left            =   135
      TabIndex        =   3
      Top             =   645
      Width           =   945
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Server, Username As String, Password As String

Public Function GetUserAndPass(Optional Cap As String = "Login")
Me.Caption = Cap
Me.Show vbModal
User = Username
Pass = Password
Ser = Server
End Function

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdLogin_Click()
Username = txtUser
Password = txtPass
Server = txtServer
Unload Me

End Sub

