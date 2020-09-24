Attribute VB_Name = "modCommand"
'@Module: modCommands
'@Desc: Interfaces related to command strings
'@Author: Anoop - anoopj13@yahoo.com

Option Explicit
Public Buffer As String
Public Connected As Boolean
Public User, Pass, Ser, Action


Public Function HandleCommand(LocalCommand As String)
'@Desc: Handles the commands
'@Comment: Use the chr$(10) command

Dim StringArray() As String
StringArray = Split(LocalCommand, Chr$(10))

Dim I As Long

For I = LBound(StringArray()) To UBound(StringArray())
    StringArray(I) = ReverseChars(Trim(StringArray(I)))
Next

'Check whether user is valid


On Error GoTo ParseError

Select Case LCase(StringArray(0))
    Case "connect"
        On Error Resume Next
        frmClient.wsConnect.Close
        Dim Host As String, Port As String
        Port = "1001"
        Host = StringArray(1)
        Port = StringArray(2)
        frmClient.wsConnect.Connect Host, Port
    Case "msg"
        ShowMessage "srv", StringArray(1)
    Case "err"
        ShowMessage "err", StringArray(1)
    Case "ok"
        On Error Resume Next
        frmClient.tvMain.Nodes.Clear
        frmClient.SendCommand "cat"
    Case "cat"
        'We got categories
        frmClient.tvMain.Nodes.Add , , "Key" & StringArray(2), StringArray(1), 1, 2
        frmClient.SendCommand "news" & Chr$(10) & StringArray(2)
    Case "news"
        frmClient.tvMain.Nodes.Add "Key" & StringArray(3), tvwChild, "SubKey" & StringArray(2), StringArray(1), 4
        
    Case "newsdata"
        Dim tmpfile As String
        tmpfile = App.Path & "\" & Format(Now, "~dd_mm_yy_hh_mm_ss") & ".html"
        Open tmpfile For Output As #1
            Print #1, StringArray(2)
        Close #1
        frmClient.lblInfo(1).Caption = "Details : " & StringArray(1)
        frmClient.wbMain.Navigate2 tmpfile
End Select

ParseError:

End Function

Function ShowMessage(Command As String, Message As String)
'@Desc: Shows a message to the window
Dim ShowMsg As String
ShowMsg = ""

Select Case LCase(Command)
    Case "cli"
        ShowMsg = "Client "
    Case "srv"
        ShowMsg = "Server "
    Case "err"
        ShowMsg = "Server "
    Case "user"
        ShowMsg = "Command"
End Select
On Error Resume Next
If Command = "err" Then
    MsgBox Message, vbCritical + vbOKOnly, "Error"
    frmClient.wsConnect.Close
    frmClient.tvMain.Nodes.Clear
    Connected = False
    frmClient.Setgui
Else
    '
End If



End Function

Public Function ReplaceChars(Data)
ReplaceChars = Replace(Data, Chr$(10), "&chr10;")
ReplaceChars = Replace(ReplaceChars, Chr$(13), "&chr13;")
End Function

Public Function ReverseChars(Data)
ReverseChars = Replace(Data, "&chr10;", Chr$(10))
ReverseChars = Replace(ReverseChars, "&chr13;", Chr$(13))
End Function

