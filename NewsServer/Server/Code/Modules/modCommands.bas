Attribute VB_Name = "modCommands"
'@Module: modCommands
'@Desc: Interfaces related to command strings
'@Author: Anoop - anoopj13@yahoo.com

Option Explicit

Public Function HandleCommand(LocalCommand As String, Index As Integer)
'@Desc: Handles the commands
'@Comment: Use the chr$(10) command

Dim StringArray() As String
StringArray = Split(LocalCommand, Chr$(10))

Dim I As Integer

For I = LBound(StringArray()) To UBound(StringArray())
    StringArray(I) = ReverseChars(Trim(StringArray(I)))
Next

'Check whether user is valid

If LCase(StringArray(0)) <> "login" And LCase(StringArray(0)) <> "signup" Then
    If GlobalUsers(Index).Username = "" Then
       SendData "err", "Should Signup or Login First. ", Index
       Exit Function
    End If
Else
    If GlobalUsers(Index).Username <> "" Then
       SendData "msg", "You are already logged in", Index
       Exit Function
    End If
End If

On Error GoTo ParseError

    Select Case LCase(StringArray(0))
        Case "login"
            
            If HandleCommandLogin(StringArray(1), StringArray(2)) Then
                    GlobalUsers(Index).Username = StringArray(1)
                SendData "msg", "Login Success At " & Now, Index
                SendData "ok", "", Index
            Else
                SendData "err", "Invalid username or password", Index
            End If
            
        Case "signup"
            If HandleCommandSignup(StringArray(1), StringArray(2), Index) Then
                    GlobalUsers(Index).Username = StringArray(1)
                    SendData "msg", "Singup Success At " & Now, Index
                    SendData "ok", "", Index
                    
            Else
                 SendData "err", "Singup Failure. Username already exists", Index
            End If
            
       Case "cat"
            HandleCommandCategory Index
            
       Case "news"
            HandleCommandNews StringArray(1), Index
            
       Case "newsdata"
            HandleCommandNewsData StringArray(1), Index
            
            
       Case Else
            SendData "msg", "Command not supported", Index
            
    End Select

Exit Function
ParseError:
    SendData "err", "Command Error. Required Parameter Missing", Index

End Function

Public Function HandleCommandLogin(Username As String, Password As String) As Boolean
    
 '@Desc: Allows the user to login
 
    Dim LocalDb As New Connection
    Dim LocalRecordset As New ADODB.Recordset
    
    Dim QueryString As String
    
    On Error GoTo NoUser
    
    LocalDb.Open Dbstring
    QueryString = "Select * from [user]"
    
    LocalRecordset.Open QueryString, LocalDb
        
    If LocalRecordset.EOF And LocalRecordset.BOF Then
        HandleCommandLogin = False
        GoTo ExitLoop
    End If
    
        Do While Not LocalRecordset.EOF
            If LocalRecordset.Fields("username").Value = Username And LocalRecordset.Fields("password").Value = Password Then
                HandleCommandLogin = True
                GoTo ExitLoop
            End If
            LocalRecordset.MoveNext
        Loop
        
ExitLoop:


    LocalRecordset.Close
    LocalDb.Close
    Exit Function
    
NoUser:
    HandleCommandLogin = False
    
End Function

Public Function HandleCommandCategory(Index As Integer)
    
 '@Desc: Allows the user to view news category
 
    Dim LocalDb As New Connection
    Dim LocalRecordset As New ADODB.Recordset
    
    Dim QueryString As String
    
    On Error GoTo NoUser
    
    LocalDb.Open Dbstring
    QueryString = "Select * from [category]"
    
    LocalRecordset.Open QueryString, LocalDb
        
    If LocalRecordset.EOF And LocalRecordset.BOF Then
        SendData "msg", "No categories in this server", Index
        GoTo ExitLoop
    End If
    
        Do While Not LocalRecordset.EOF
        On Error Resume Next
            SendData "cat", CStr(LocalRecordset.Fields("Categoryname").Value) & Chr$(10) & CStr(LocalRecordset.Fields("CategoryId").Value), Index
            LocalRecordset.MoveNext
        Loop
        
ExitLoop:


    LocalRecordset.Close
    LocalDb.Close
    Exit Function
    
NoUser:
    
End Function


Public Function HandleCommandNews(Category, Index As Integer)
    
 '@Desc: Allows the user to view news
 
    Dim LocalDb As New Connection
    Dim LocalRecordset As New ADODB.Recordset
    
    Dim QueryString As String
    
    On Error GoTo NoUser
    
    LocalDb.Open Dbstring
    QueryString = "Select * from [news] where categoryid=" & Category
    
    LocalRecordset.Open QueryString, LocalDb
        
    If LocalRecordset.EOF And LocalRecordset.BOF Then
        SendData "msg", "No news in this category", Index
        GoTo ExitLoop
    End If
    
        Do While Not LocalRecordset.EOF
        On Error Resume Next
            SendData "news", CStr(LocalRecordset.Fields("Subject").Value) & Chr$(10) & CStr(LocalRecordset.Fields("newsid").Value) & Chr$(10) & CStr(LocalRecordset.Fields("categoryid").Value), Index
            LocalRecordset.MoveNext
        Loop
        
ExitLoop:


    LocalRecordset.Close
    LocalDb.Close
    Exit Function
    
NoUser:
    
End Function

Public Function HandleCommandNewsData(Id, Index As Integer)
    
 '@Desc: Allows the user to view news data
 
    Dim LocalDb As New Connection
    Dim LocalRecordset As New ADODB.Recordset
    
    Dim QueryString As String
    
    On Error GoTo NoUser
    
    LocalDb.Open Dbstring
    QueryString = "Select * from [news] where newsid=" & Id
    Debug.Print Id
    
    LocalRecordset.Open QueryString, LocalDb
        
    If LocalRecordset.EOF And LocalRecordset.BOF Then
        GoTo ExitLoop
        SendData "msg", "No news data for the requested id", Index
    End If
    
        Do While Not LocalRecordset.EOF
        On Error Resume Next
        Dim mDate As String
        mDate = ""
        mDate = CStr(LocalRecordset.Fields("Date").Value)
            SendData "newsdata", CStr(LocalRecordset.Fields("Subject").Value) & Chr$(10) & ReplaceChars(CStr(LocalRecordset.Fields("Description").Value)) & Chr$(10) & mDate, Index
            LocalRecordset.MoveNext
        Loop
        
ExitLoop:


    LocalRecordset.Close
    LocalDb.Close
    Exit Function
    
NoUser:
    
End Function


Public Function HandleCommandSignup(Username As String, Password As String, Index As Integer) As Boolean
    
'@Desc: Allows the user to signup
     
    Dim LocalRecordset As ADODB.Recordset
    Dim QueryString As String
    
    On Error GoTo NoUser
    
    Set LocalRecordset = New ADODB.Recordset
    
    
    QueryString = "INSERT INTO [User] ( UserName, [Password] ) Values (" & QT(Username) & ", " & QT(Password) & ");"
    
    LocalRecordset.Open QueryString, GlobalDatabase
    
    On Error Resume Next
    LocalRecordset.Close
    HandleCommandSignup = True
    Exit Function
    
NoUser:
    HandleCommandSignup = False
    
    
End Function


Public Function QT(St As String)
QT = """" & St & """"
End Function
