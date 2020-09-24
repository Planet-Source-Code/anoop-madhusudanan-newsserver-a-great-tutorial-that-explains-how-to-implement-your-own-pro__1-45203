VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "I-News Server"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4020
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   4020
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   75
      TabIndex        =   4
      Top             =   1455
      Width           =   3855
   End
   Begin VB.TextBox txtLog 
      Height          =   2850
      Left            =   75
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1650
      Width           =   3885
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2700
      TabIndex        =   1
      Top             =   810
      Width           =   1125
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   810
      Width           =   1125
   End
   Begin MSWinsockLib.Winsock wsListen 
      Left            =   1605
      Top             =   1785
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsAccept 
      Index           =   0
      Left            =   1620
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmServer.frx":08CA
      Top             =   15
      Width           =   480
   End
   Begin VB.Label lblInfo 
      Caption         =   "The server is not running"
      Height          =   420
      Left            =   780
      TabIndex        =   2
      Top             =   225
      Width           =   2790
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Module: frmServer
'@Desc: Main Server Form
'@Author: Anoop - anoopj13@yahoo.com

Option Explicit

'------------------------------------------------------------------------------------
'SESSION LAYER PROTOCOL EXAMPLE
'------------------------------------------------------------------------------------
'By Anoop Madusudanan, http://www.inetsindia.com/anoopvision, anoopj13@yahoo.com
'Friday, May 02, 2003
'
'------------------------------------------------------------------------------------
'WHY THIS ARTICLE?
'------------------------------------------------------------------------------------
'
'To teach you how to design session layer protocols like HTTP and FTP
'(if you are interested ;) ).
'
'Read Readme.htm with this package for a better overview.
'
'The zip file contains 3 projects
'
'a) NewsServer - This project
'b) I-News GUI client - A client with a GUI (in the ClientGUI folder)
'c) I-News Client - A telnet like client (in the Client folder)
'
'
'------------------------------------------------------------------------------------
'READ THIS!!!
'------------------------------------------------------------------------------------
'
'I-News server is a TCP/IP based server suit, with limited functionality to
'serve news over a network. Rather than working as a news server, this application
'is to demonstrate the designing and implementation of of custom application protocols
'over TCP/IP. This package contains the Server in the server directory,
'along with two clients. In the folder ClientGUI, there is a visual client,
'and in the folder Client, there is a simple text based client.
'
'The server and client exchanges messages as plain text. Each argument
'is separated with ASCII character 10 (chr$(10)) and each lines are separated
'with ASCII character 13 (chr$(13)) . It is possible to send and receive data in
'various chunks, because both server and client supports buffered collection of data.
'That is, even after receiving the data, the algorithm will keep data in a buffer as
'long as it receives a chr$(13) character. Then, the command is executed.
'
'For example, data can be send like this (& represents the appending operator in VB,
'and ws is a socket).
'
'ws.SendData "login" & Chr$(10) & "user" & Chr$(10) & "password" & Chr$(13)
'
'Also, sending data as two chunks has the same result as above.
'
'ws.SendData "login" & Chr$(10) & "user"
'ws.SendData Chr$(10) & "password" & Chr(13)

'------------------------------------------------------------------------------------
'HOW TO START
'------------------------------------------------------------------------------------
'
'1) Start this project in Visual Basic IDE and click Start button to start the server
'2) Start another instance of Visual Basic, open the project ClientGUI, and start it.
'3) Click Login button in ClientGUI and type 'localhost' as server, 'admin' as username and 'admin' as password. You are in
'
'------------------------------------------------------------------------------------
'AND FINALLY..
'------------------------------------------------------------------------------------
'
'1) Visit my site at http://www.inetsindia.com/anoopvision for more code and tutorials
'2) Give me your vote for this at PSC
'------------------------------------------------------------------------------------
'
'Regards, An 'OOP' - anoopj13@yahoo.com
'
'------------------------------------------------------------------------------------




Public Function UserConnectionListen()
'@Desc: Starts listening
On Error GoTo FunctionError

wsListen.LocalPort = LISTEN_PORT
wsListen.Listen



Exit Function
FunctionError:
MsgBox "Unable to start listening.", vbInformation + vbOKOnly
Unload Me

End Function

Public Function HandleError(Error As ErrObject)
'@Desc: Shows message or something

End Function

Private Sub cmdStart_Click()
On Error Resume Next

InitUsers
cmdStop.Enabled = True
cmdStart.Enabled = False

UserConnectionListen

If Err Then
    Me.lblInfo = "Server starting error"
    cmdStop_Click
Else
    Me.lblInfo = "The server is running.."
End If

End Sub

Private Sub cmdStop_Click()
On Error Resume Next

SendDataAll "msg", "Server is going to shut down"

cmdStop.Enabled = False
cmdStart.Enabled = True

Dim I As Integer

For I = 0 To MAX_CONNECTION
    CloseUserConnection I
Next

wsListen.Close



End Sub

Private Sub Form_Load()
On Error Resume Next
'Start automatically
cmdStart_Click

End Sub

Private Sub wsAccept_Close(Index As Integer)
On Error Resume Next
CloseUserConnection Index

End Sub

Private Sub wsAccept_DataArrival(Index As Integer, ByVal bytesTotal As Long)
'@Desc: Data arrived
On Error Resume Next
Dim Data As String
Dim SplitData() As String
Dim I As Integer

wsAccept(Index).GetData Data, vbString


If InStr(1, Data, Chr$(13)) Then
SplitData = Split(wsAccept(Index).Tag & Data, Chr$(13))

    For I = LBound(SplitData()) To UBound(SplitData()) - 1
        HandleCommand SplitData(I), Index
        wsAccept(Index).Tag = SplitData(UBound(SplitData()))
    Next

Else
wsAccept(Index).Tag = wsAccept(Index).Tag & Data
End If


End Sub


Private Sub wsAccept_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
CloseUserConnection Index
End Sub

Private Sub wsAccept_SendComplete(Index As Integer)
GlobalUsers(Index).Sending = False
End Sub

Private Sub wsListen_ConnectionRequest(ByVal requestID As Long)
'@Desc: We got a request

Call modConnection.AddUserConnection(requestID)
    

End Sub

Public Function ShowStats(Msg)
'@Desc: Put some info in the text box

On Error Resume Next
Me.txtLog.Text = Me.txtLog.Text & Msg & vbCrLf
If Err Then Me.txtLog.Text = ""
End Function

