VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmClient 
   BackColor       =   &H00FFFFFF&
   Caption         =   "News Client 1.0"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5940
   Icon            =   "frmClient.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   5940
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox filMain 
      Height          =   1065
      Left            =   3855
      Pattern         =   "~*.*"
      TabIndex        =   10
      Top             =   3045
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picPanel 
      Align           =   1  'Align Top
      BackColor       =   &H00E9E9E9&
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   5940
      TabIndex        =   0
      Top             =   0
      Width           =   5940
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2595
         Picture         =   "frmClient.frx":1CFA
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Click to refresh the contents"
         Top             =   60
         Width           =   1095
      End
      Begin VB.CommandButton cmdSignup 
         Caption         =   "Signup"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   135
         Picture         =   "frmClient.frx":2B3C
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Click to signup with a server"
         Top             =   60
         Width           =   1095
      End
      Begin VB.CommandButton cmdLogin 
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   1380
         Picture         =   "frmClient.frx":3406
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Click to login to a server"
         Top             =   45
         Width           =   1095
      End
   End
   Begin VB.PictureBox picNews 
      BackColor       =   &H00EEBEA2&
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   4665
      TabIndex        =   8
      Top             =   390
      Width           =   4665
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DETAILS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   75
         TabIndex        =   9
         Top             =   60
         Width           =   825
      End
   End
   Begin SHDocVwCtl.WebBrowser wbMain 
      Height          =   2895
      Left            =   615
      TabIndex        =   7
      Top             =   2745
      Width           =   4140
      ExtentX         =   7302
      ExtentY         =   5106
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComctlLib.ImageList imgMain 
      Left            =   495
      Top             =   3855
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient.frx":3CD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient.frx":426A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient.frx":4804
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient.frx":4D9E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BackColor       =   &H00EEBEA2&
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   5940
      TabIndex        =   2
      Top             =   1005
      Width           =   5940
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NEWS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   6
         Top             =   60
         Width           =   555
      End
   End
   Begin MSComctlLib.TreeView tvMain 
      Height          =   3240
      Left            =   45
      TabIndex        =   1
      Top             =   1440
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   5715
      _Version        =   393217
      Indentation     =   353
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imgMain"
      Appearance      =   0
   End
   Begin MSWinsockLib.Winsock wsConnect 
      Left            =   4140
      Top             =   1545
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Module: frmClient
'@Desc: Main client form
'@Author: Anoop - anoopj13@yahoo.com

Option Explicit


Private Sub cmdLogin_Click()

If LCase(cmdLogin.Caption) = "login" Then
    Connected = False
    frmLogin.GetUserAndPass "Login"
    Action = "login"
    On Error Resume Next
    Me.wsConnect.Close
    Me.wsConnect.Connect Ser, 1001
    cmdLogin.Caption = "Logout"
Else
    On Error Resume Next
    Connected = False
    wsConnect.Close
    Setgui
    cmdLogin.Caption = "Login"
End If

End Sub

Private Sub cmdRefresh_Click()
On Error Resume Next
Me.tvMain.Nodes.Clear
SendCommand "cat"

End Sub

Private Sub cmdSignup_Click()


'Check whether this is just a message
Connected = False
frmLogin.GetUserAndPass "Sign up"
Action = "sign"
On Error Resume Next
Me.wsConnect.Close
Me.wsConnect.Connect Ser, 1001

End Sub

Private Sub Form_Load()
On Error Resume Next
filMain.Path = App.Path
Connected = False
Buffer = ""
frmClient.wbMain.Navigate2 App.Path & "\help.htm"
Setgui
End Sub

Private Sub Form_Resize()
On Error Resume Next

Me.tvMain.Move 0, Me.picTop.Top + picTop.Height, Me.ScaleWidth, Me.ScaleHeight - wbMain.Height - picNews.Height - tvMain.Top
picNews.Top = tvMain.Top + tvMain.Height
picNews.Left = 0
picNews.Width = Me.ScaleWidth
wbMain.Left = 0
wbMain.Width = Me.ScaleWidth

wbMain.Top = picNews.Top + picNews.Height


End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Dim I As Integer

For I = 0 To filMain.ListCount - 1
On Error Resume Next
    Kill App.Path & "\" & filMain.List(I)
Next

End Sub

Private Sub tvMain_DblClick()
On Error Resume Next
Dim Selnode As Node
Set Selnode = tvMain.SelectedItem
If Err Then Exit Sub
'Get the news data
SendCommand "newsdata" & Chr$(10) & Right(Selnode.Key, Len(Selnode.Key) - 6)

End Sub

Private Sub wsConnect_Close()
Connected = False
On Error Resume Next

Setgui

End Sub

Private Sub wsConnect_Connect()

Buffer = ""
If Connected = False Then
    If Action = "sign" Then
        SendCommand "Signup" & Chr$(10) & User & Chr$(10) & Pass
    Else
        SendCommand "Login" & Chr$(10) & User & Chr$(10) & Pass
    End If
    Connected = True
End If
Setgui

End Sub

Private Sub wsConnect_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim Data As String
Dim SplitData() As String
Dim I As Integer

wsConnect.GetData Data, vbString


If InStr(1, Data, Chr$(13)) Then
SplitData = Split(Buffer & Data, Chr$(13))

    For I = LBound(SplitData()) To UBound(SplitData()) - 1
        HandleCommand SplitData(I)
        Buffer = SplitData(UBound(SplitData()))
    Next

Else
Buffer = Buffer & Data
End If


End Sub


Private Sub wsConnect_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Connected = False
On Error Resume Next
MsgBox "An error occurred. " & Description, vbCritical + vbOKOnly, "Connection Error"
Setgui
End Sub

Public Function Setgui()
If Connected = False Then
    Me.tvMain.Nodes.Clear
    Me.cmdSignup.Enabled = True
    Me.cmdRefresh.Enabled = False
    cmdLogin.Caption = "Login"
    User = ""
    Pass = ""
    Ser = ""
    Action = ""
    Me.lblInfo(1).Caption = "Info"
    
Else
    Me.cmdSignup.Enabled = False
    Me.cmdRefresh.Enabled = True
    
End If

End Function

Public Function SendCommand(Command As String)
    On Error Resume Next
    wsConnect.SendData Command & Chr$(13)
    If Err Then
        MsgBox "An error occurred. " & Err.Description, vbCritical + vbOKOnly, "Connection Error"
        wsConnect.Close
        Setgui
    End If
End Function
