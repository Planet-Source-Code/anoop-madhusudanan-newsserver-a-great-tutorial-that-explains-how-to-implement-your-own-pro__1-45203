Attribute VB_Name = "modMain"
'@Module: modMain
'@Desc: Interfaces Main
'@Author: Anoop - anoopj13@yahoo.com

Option Explicit

Public LISTEN_PORT As Long


'DB Constants
Public GlobalDatabase As ADODB.Connection
Public Dbstring As String


Public Sub Main()
'@Desc: We start here

    InitDatabase
    frmServer.Show
    
    
End Sub

Public Sub InitUsers()
'@Desc: Initialses the GlobalUsers array

On Error Resume Next

'Initialise global variables
MAX_CONNECTION = 50
LISTEN_PORT = 1001

ReDim GlobalUsers(MAX_CONNECTION)

Dim I As Integer

For I = 0 To MAX_CONNECTION
    GlobalUsers(I).Connected = False
    GlobalUsers(I).Username = ""
Next

End Sub

Public Function InitDatabase() As Boolean
'@Desc: Initialise the database
On Error GoTo DbError
Set GlobalDatabase = New ADODB.Connection
GlobalDatabase.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\News.mdb;Persist Security Info=False"
Dbstring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\News.mdb;Persist Security Info=False"
InitDatabase = True

Exit Function
DbError:
InitDatabase = False
MsgBox Err.Description

End Function

