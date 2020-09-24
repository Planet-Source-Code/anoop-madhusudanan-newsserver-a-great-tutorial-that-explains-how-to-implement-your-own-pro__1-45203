Attribute VB_Name = "modMain"
Public Connected As Boolean
Public Buffer As String

Public Function HandleCommand(LocalCommand As String)
'@Desc: Handles the commands
'@Comment: Use the chr$(10) command

Dim StringArray() As String
StringArray = Split(LocalCommand, Chr$(10))

Dim I As Long

For I = LBound(StringArray()) To UBound(StringArray())
    StringArray(I) = Trim(StringArray(I))
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
    Case "cat"
        ShowMessage "srv", StringArray(2) & "-" & StringArray(1)
    Case "news"
        ShowMessage "srv", StringArray(2) & "-" & StringArray(1)
    Case "newsdata"
        ShowMessage "srv", StringArray(1) & vbCrLf & "---------------------------------------------------" & vbCrLf & StringArray(2)
End Select

ParseError:

End Function

Function ShowMessage(Command As String, Message As String)
'@Desc: Shows a message to the window
Dim ShowMsg As String
ShowMsg = ""

Select Case Command
    Case "cli"
        ShowMsg = "Client "
    Case "srv"
        ShowMsg = "Server "
    Case "user"
        ShowMsg = "Command"
End Select
On Error Resume Next

If ShowMsg <> "" Then
    frmClient.txtConsole.Text = frmClient.txtConsole.Text & "[" & ShowMsg & "] " & Message & vbCrLf
End If

If Err Then
    frmClient.txtConsole.Text = ""
    frmClient.txtConsole.Text = frmClient.txtConsole.Text & "[" & ShowMsg & "] " & Message & vbCrLf
End If

frmClient.txtConsole.SelStart = Len(frmClient.txtConsole.Text) - 1

End Function


Public Function StartsWith(String1 As String, String2 As String, Optional CheckCase As Boolean = False) As Boolean
'@Desc: Checks whether String1 starts with String2
StartsWith = False

If CheckCase = False Then
    If Left(LCase(String1), Len(String2)) = LCase(String2) Then StartsWith = True
Else
    If Left(String1, Len(String2)) = String2 Then StartsWith = True
End If

End Function

