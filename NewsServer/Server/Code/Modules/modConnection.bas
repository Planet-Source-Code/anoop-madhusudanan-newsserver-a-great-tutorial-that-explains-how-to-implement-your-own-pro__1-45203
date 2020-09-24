Attribute VB_Name = "modConnection"
'@Module: modConnection
'@Desc: User connection management
'@Author: Anoop - anoopj13@yahoo.com

Option Explicit

Public MAX_CONNECTION As Long

Public Type UserConnection
       Connected As Boolean
       SocketIndex As Integer
       Username As String
       Sending As Boolean
End Type

Public GlobalUsers() As UserConnection


Public Function AddUserConnection(Request As Long) As Boolean
'@Desc: Loads a new socket and add a new connection

AddUserConnection = True
Dim I As Integer
For I = 0 To MAX_CONNECTION
    If GlobalUsers(I).Connected = False Then
        On Error Resume Next
        Unload frmServer.wsAccept(I)
        Load frmServer.wsAccept(I)
        Err.Clear
        
        frmServer.wsAccept(I).Accept Request
        frmServer.wsAccept(I).Tag = ""
        GlobalUsers(I).Connected = True
        
        frmServer.ShowStats "Connection: " & frmServer.wsAccept(I).RemoteHostIP & ":" & frmServer.wsAccept(I).RemotePort
        If Err Then AddUserConnection = False
        Exit Function
    End If
Next

End Function

Public Function CloseUserConnection(Index As Integer)
'@Desc: Closes a user connection

On Error Resume Next
GlobalUsers(Index).Sending = False
frmServer.ShowStats "Connection Close : " & frmServer.wsAccept(Index).RemoteHostIP & ":" & frmServer.wsAccept(Index).RemotePort
frmServer.wsAccept(Index).Close
Unload frmServer.wsAccept(Index)
GlobalUsers(Index).Connected = False
GlobalUsers(Index).Username = ""


End Function

Public Function SendData(Command As String, Data As String, Index As Integer)
'Desc: Sends data to a socket
On Error Resume Next
Dim Dummy

If GlobalUsers(Index).Connected = True Then
    frmServer.wsAccept(Index).SendData Command & Chr$(10) & Data & Chr$(13)
    GlobalUsers(Index).Sending = True
End If

End Function


Public Function SendDataAll(Command As String, Data As String)
'Desc: Sends data to all open sockets
On Error Resume Next
Dim Index As Integer

For Index = 0 To MAX_CONNECTION
        SendData Command, Data, Index
Next

End Function


Public Function ReplaceChars(Data)
ReplaceChars = Replace(Data, Chr$(10), "&chr10;")
ReplaceChars = Replace(ReplaceChars, Chr$(13), "&chr13;")
End Function

Public Function ReverseChars(Data)
ReverseChars = Replace(Data, "&chr10;", Chr$(10))
ReverseChars = Replace(ReverseChars, "&chr13;", Chr$(13))
End Function


