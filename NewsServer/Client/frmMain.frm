VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   Caption         =   "Client"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8415
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock wsConnect 
      Left            =   2145
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   435
      Left            =   7365
      TabIndex        =   2
      Top             =   4470
      Width           =   915
   End
   Begin VB.TextBox txtCommand 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   435
      Left            =   120
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   4485
      Width           =   7155
   End
   Begin VB.TextBox txtConsole 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   4185
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   60
      Width           =   8190
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSend_Click()

If Trim(txtCommand.Text) = "" Then Exit Sub

'Check whether this is just a message


If Connected = False Then
    ShowMessage "user", Me.txtCommand
    If StartsWith(Me.txtCommand, "connect") Then
        HandleCommand Replace(Me.txtCommand.Text, "|", Chr$(10))
        ShowMessage "cli", "Opening connection..."
    Else
        ShowMessage "cli", "Not connected to server. Use the 'Connect' command to connect"
    End If

Else
'We are connected
    On Error Resume Next
    Me.wsConnect.SendData Replace(Me.txtCommand.Text, "|", Chr$(10)) & Chr$(13)
    
    If Err Then
    On Error Resume Next
        ShowMessage "user", Me.txtCommand
        ShowMessage "srv", "Connection Error. Disconnected"
        wsConnect.Close
        Connected = False
    End If

End If

txtCommand.Text = ""

End Sub

Private Sub Form_Load()
Connected = False
Buffer = ""

End Sub

Private Sub Form_Resize()
On Error Resume Next
Me.txtConsole.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - txtCommand.Height - 100
Me.txtCommand.Move 0, Me.txtConsole.Height + txtConsole.Top + 50, Me.ScaleWidth - cmdSend.Width - 100
Me.cmdSend.Move Me.txtCommand.Width + Me.txtCommand.Left + 50, Me.txtCommand.Top
End Sub

Private Sub txtCommand_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = vbKeyReturn Then
        
End If

End Sub

Private Sub wsConnect_Connect()
Connected = True
ShowMessage "cli", "Connected To Server.."
Buffer = ""

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
ShowMessage "cli", "Socket Error " & Description

End Sub

