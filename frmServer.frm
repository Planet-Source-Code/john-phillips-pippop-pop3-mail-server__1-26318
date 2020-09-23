VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   Caption         =   "PipPop Server"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   3120
      Width           =   6375
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3000
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.TextBox Text1 
         Height          =   2055
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   720
         Width           =   6375
      End
      Begin VB.Label Label4 
         Caption         =   "POP3 Window"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Started"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSWinsockLib.Winsock WSServer 
      Index           =   0
      Left            =   1080
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock WSSMTP 
      Index           =   0
      Left            =   240
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label5 
      Caption         =   "SMTP Window"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   4920
      TabIndex        =   3
      Top             =   4920
      Width           =   1575
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' PipPop POP3 Email Server copyright ©2001 CSMG, Inc. (John Phillips)
' Created by John Phillips, MCP
' Computer Systems Management Group, Inc.
' This version of pipop was thrown together real quick for
' upload to Planet-Source-Code
' As it stands right now pipop was only tested on windows 98
' I plan to test it on modify as needed on winows NT , 2000
' This version is not complete and was only tested with one client connecting at
' a single session - multiple client support will be supported in
' future releases - also this version does not support attachments
'
' This was also tested using Outlook express
' so let me know if there is a problem connecting
' with other mail clients
'
' any question about this server please email me at
' vbjack@nyc.rr.com with PIPPOP Code somewhere in the subject Line
'
' the reason i started creating such a server is simple
' I need to provide email support for most of my clients
' and they want an inexpensive way of doing so
' so after finding out that the free POP3 servers arent that easy to understand
' and the easy ones are very expensive for multible clients
' I decided to see if I could create my own
' I will be uploading updates to this server for everyone to use
' provided anyone who uses this code put my name somewhere in the credits
' with CSMG, Inc.

Private mailBody As String
Private SubjectText As String
Private intMax As Long
Dim sCLF As String
Dim sUser(100) As String
Dim sMailFrom As String
Dim sMailTo As String

' message array (1st holds which user we are holding messages for
' 2nd holds the message number and the 3rd holds the message
Dim aMessage(100, 100, 1) As String

Private Sub Form_Load()

sCLF = Chr(13) & Chr(10)
intMax = 0 ' set the winsock array to 0

' POP3 uses port 110 unless it is a secure connection (encrypted)
' which is not used in this example
' SMTP uses port 25
WSServer(0).LocalPort = "110"
WSSMTP(0).LocalPort = "25"
' start the winsock control to listen
WSServer(0).Listen
WSSMTP(0).Listen
' load the databases
Data1.DatabaseName = App.Path & "\pippop.mdb"
Data2.DatabaseName = App.Path & "\pippop.mdb"

' setup the user database
Data1.RecordSource = "SELECT * FROM users"
Data1.Refresh
Data1.UpdateControls

'display the server hostname and ip address
Label2.Caption = WSServer(0).LocalHostName
Label3.Caption = WSServer(0).LocalIP

End Sub


Private Sub WSServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)

' we are getting a connection request on the winsock control 0 (first)
' now we set the incoming request to a new winsock control in the array
If Index = 0 Then
      intMax = intMax + 1 ' set a new value for the array
      Load WSServer(intMax) ' load the new winsock control
      WSServer(intMax).LocalPort = 110 ' set the port for the new control
      WSServer(intMax).Accept requestID ' accept the request
      
      ' now start the POP3 Session by telling the Client that
      ' the server is reay for whatever it wants to try and do
      WSServer(intMax).SendData "+OK PipPop POP3 Server Ver. 1.0.0 Ready" & vbCrLf

End If

' just display some information for the users
' this is not needed but can be used for a log file!
' I didnt setup logging in this version I will do that in the next version
Text1.Text = "CLIENT CONNECTING TO POP3" & requestID & vbCrLf

End Sub

Private Sub WSServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
'On Error GoTo errDA
    
    ' setup a string for the incoming data
    Dim strData As String
    
    ' get the incoming data from the winsock control
    WSServer(Index).GetData strData, vbString
        
    ' display the incoming data for the user
    ' again can be usefull for a log file but not needed
    Text1.Text = Text1.Text & strData & vbCrLf
        
    ' now we need to determine what the client that is connecting wants
    ' the Client will always send a string and the first 4 Chr's are
    ' captial and always in length of 4
    Select Case Left(strData, 4) ' just get the first 4 chr's
    Case "USER"
    ' the client is sending us the user name to see if the user has an account here
    ' now we check
    
    ' first we setup the user array with the user name
    ' the user name will always begin in the 6th position of the data sent
    sUser(Index) = Trim(Mid(strData, 6))
    ' we need to minus 2 off the end of the string since we also recieve
    ' CR & LF from the client
    sUser(Index) = Mid(sUser(Index), 1, (Len(sUser(Index)) - 2))
    ' user array is now setup
    
    ' now call the chkuser function to see if the user is in the database
    ChkUser sUser(Index), Index
    
    Case "PASS"
    Dim tPass As String
    ' well the user must have been validated in the database
    ' or the client would have been sent an error message and we would have disconnected
    ' so now we need to check the password against the user name
    tPass = Trim(Mid(strData, 6))
    ' we need to minus 2 off the end of the string since we also recieve
    ' CR & LF from the client
    tPass = Mid(tPass, 1, (Len(tPass) - 2))
    ' first we need to get the password from the data sent
    ' we already know which user we are checking against since
    ' we set the user in the array to the same number of the winsock
    ' control in the array
    ChkPass tPass, Index
    
    Case "STAT"
    ' lets get the statistics from the server
    GetStats Index
        
    Case "DELE"
    ' lets delete all the message from the server
    DeleMess Index, Mid(strData, 6, 1)
    'DelMess
    Case "RETR"
    ' this is only setup to retrieve up to 9 message
    ' I will set it up to except more message in the next version
    RetrMess Index, Mid(strData, 6, 1)
    'RetrMail
    Case "LIST"
    ' lets list all the mail in the server for given user
    ListMail Index
    Case "UIDL"
    ' lets list all the mail in the server for given user
    ListMail Index
    Case "QUIT"
    ' now quit and sign off from client
    WSServer(Index).SendData "+OK POP3 server signing off (maildrop empty)" & vbCrLf
    'WSServer(Index).SendData "." & vbCrLf
    'WSServer(Index).Close
    'MsgBox "quit command rec."
    'CloseSession
    Case Else
    MsgBox "Error: " & vbCrLf & strData
    End Select
    'WSServer(Index).Close
    'MsgBox strData
Exit Sub
errDA:
MsgBox "Error IN Data Avrival Module" & vbCrLf & Err.Number & vbCrLf & Err.Description
End Sub

Private Function ChkUser(sName As String, i As Integer) As Boolean
On Error GoTo errDA

' find the user
Data1.RecordSource = "SELECT * FROM users WHERE username='" & sName & "'"
Data1.Refresh
Data1.UpdateControls

' no records found
If Data1.Recordset.EOF = True And Data1.Recordset.BOF = True Then
ChkUser = False
' send error message
WSServer(i).SendData "-ERR " & sUser(i) & " is not a user here." & vbCrLf

' display text for the user
Text1.Text = Text1.Text & "user Denied" & vbCrLf
' close the conection
WSServer(i).Close
' unload the winsock control since it is not being used anymore
' this also prevents resources from being sucked up
Unload WSServer(i)
Exit Function
Else
ChkUser = True
' the user was found in the database
' since I didn't set this up to have multible users
' with the same login name if it found a record then
' the user is a valid user

' since we found a user we send out the +OK response
' to let the client know that the user is valid here
WSServer(i).SendData "+OK " & sUser(i) & " is welcome here" & vbCrLf

' just visual confirmation for the user
Text1.Text = Text1.Text & "user verified" & vbCrLf

Exit Function
End If

Exit Function
errDA:
' for testing purpose we are not going to disconnect the winsock control
' if an error is generated - we will just display a msgbox with the error
' for a completed project with error handling we would log the error
' in a log file and just disconnect the client and unload the winsock control
MsgBox "Error IN Check User Module" & vbCrLf & Err.Number & vbCrLf & Err.Description

End Function

Private Function ChkPass(sPass As String, i As Integer) As Boolean
On Error GoTo errDA

' first we need to get the user account from the database
Data1.RecordSource = "SELECT * FROM users WHERE username='" & sUser(i) & "'" '& "' AND userpassword='" & sPass & "'"
Data1.Refresh
Data1.UpdateControls

' now we check to see if the password supplied is valid
' for this user
If Data1.Recordset.Fields("userpassword") <> sPass Then
ChkPass = False
' password isn't valid - send error message to client
WSServer(i).SendData "+ERR " & sUser(i) & "'s password is incorrect" & vbCrLf

' then disconnect and unload the winsock control
WSServer(i).Close
Unload WSServer(i)
Exit Function
Else
ChkPass = True
' the password is valid so send the +OK message to the client
WSServer(i).SendData "+OK " & susers & "'s mailbox has 0 messages (0 octets)" & vbCrLf
 
Exit Function
End If



Exit Function
errDA:
' for testing purpose we are not going to disconnect the winsock control
' if an error is generated - we will just display a msgbox with the error
' for a completed project with error handling we would log the error
' in a log file and just disconnect the client and unload the winsock control
MsgBox "Error IN Check Pass Module" & vbCrLf & Err.Number & vbCrLf & Err.Description

End Function

Private Function GetStats(i As Integer) As Boolean
Dim tID As Integer

' first get the user from the database so we can get the ID
' this can be setup in the user array but for now we are just going to do it this way
' I will set that up in the next version
Data1.RecordSource = "SELECT * FROM users WHERE username='" & sUser(i) & "'"
Data1.Refresh
Data1.UpdateControls

If Data1.Recordset.EOF = True And Data1.Recordset.BOF = True Then
' enter a log file entry here
' no user found - it should not have gone this far
' send error message to the client and disconnect
WSServer(i).SendData "-ERR User was not found" & vbCrLf

WSServer(i).Close
Unload WSServer(i)
GetStats = False
Exit Function
Else
tID = Data1.Recordset.Fields("userid")
' we found the user and got his/her ID

' now get the mail from the mail table
Data1.RecordSource = "SELECT * FROM mail WHERE userid=" & tID
Data1.Refresh
Data1.UpdateControls

If Data1.Recordset.EOF = True And Data1.Recordset.BOF = True Then
' no mail found - send message to client
WSServer(i).SendData "+OK 0 0" & vbCrLf
Exit Function
Else
Dim nMessages As Integer 'counter for the number of messages
Dim nOctets As Integer ' setup the number of octets for the message

' number of bytes for each message
' An octet is an 8-bit byte
' 1 byte is usually 8 bits
' I wont go into detail, just take my word for now
' the knowledge base has info on octets
nOctets = 0
nMessages = 0

' get the total record count (message count)
Data1.Recordset.MoveLast
nMessages = Data1.Recordset.RecordCount
Data1.Recordset.MoveFirst

' add up all the octets
Do While Data1.Recordset.EOF = False
nOctets = nOctets + Data1.Recordset.Fields("octets")
Data1.Recordset.MoveNext
Loop

' send the stats to the client
WSServer(i).SendData "+OK " & nMessages & " " & nOctets & vbCrLf
GetStats = True
End If
End If
End Function


Private Function ListMail(i As Integer) As Boolean
Dim tID As Integer

' first get the user from the database so we can get the ID
' this can be setup in the user array but for now we are just going to do it this way
' I will set that up in the next version
Data1.RecordSource = "SELECT * FROM users WHERE username='" & sUser(i) & "'"
Data1.Refresh
Data1.UpdateControls

If Data1.Recordset.EOF = True And Data1.Recordset.BOF = True Then
' enter a log file entry here
' no user found - it should not have gone this far

' send error message to client
WSServer(i).SendData "-ERR User was not found" & vbCrLf

WSServer(i).Close
Unload WSServer(i)
ListMail = False

Else
tID = Data1.Recordset.Fields("userid")
' user found - get the ID

Data1.RecordSource = "SELECT * FROM mail WHERE userid=" & tID
Data1.Refresh
Data1.UpdateControls

If Data1.Recordset.EOF = True And Data1.Recordset.BOF = True Then
' some how we didnt find any messages and the client should not have got this far
' send error message to the client and disconnect
WSServer(i).SendData "+ERR no messages found" & vbCrLf

WSServer(i).Close
Unload WSServer(i)

Else
Dim nMessages As Integer ' setup total message value
Dim nOctets As Integer ' setup octet string
Dim nC As Integer '

nOctets = 0
nC = 1
Data1.Recordset.MoveLast
nMessages = Data1.Recordset.RecordCount
Data1.Recordset.MoveFirst

Do While Data1.Recordset.EOF = False
nOctets = nOctets + Data1.Recordset.Fields("octets")
Data1.Recordset.MoveNext
Loop

Data1.Recordset.MoveFirst

Dim sMess As String
' send first value to the client for the LIST command
WSServer(i).SendData "+OK " & nMessages & " messages (" & nOctets & " octets)" & vbCrLf

' send the list 1 by 1
Do While Data1.Recordset.EOF = False
WSServer(i).SendData nC & " " & Data1.Recordset.Fields("octets") & vbCrLf
' add the message to the array of messages
' I am not going to explain arrays here (sorry)
' you can find info in the knowledge base
aMessage(i, nC, 1) = Data1.Recordset.Fields("message")
Data1.Recordset.MoveNext
nC = nC + 1 ' message number
Loop

' this sent at the end to let the client know that the list
' is completed
WSServer(i).SendData "." & vbCrLf
ListMail = True
End If
End If

End Function

Private Function RetrMess(i As Integer, nM As Integer) As Boolean


If aMessage(i, nM, 1) = "" Then
' enter a log file entry here
' no message in the array - send error to client and disconnect
WSServer(i).SendData "-ERR Message not found" & vbCrLf
WSServer(i).Close
Unload WSServer
RetrMess = False
Else

' send message
WSServer(i).SendData "+OK " & Len(aMessage(i, nM, 1)) & " octets" & vbCrLf
WSServer(i).SendData aMessage(i, nM, 1) & vbCrLf
' send end of message
WSServer(i).SendData vbCrLf & "." & vbCrLf
End If
End Function

Private Function DeleMess(i As Integer, iMN As Integer) As Boolean
' we are not really deleteing this message for testing purposes
' it would get annoying to keep adding a new message everytime we
' deleted it
'
' so we just sending the message deleted string to the client
' the next version will include actually deleteing the message from the server

WSServer(i).SendData "+OK message " & iMN & " deleted" & vbCrLf

End Function

Private Sub WSSMTP_ConnectionRequest(Index As Integer, ByVal requestID As Long)
If Index = 0 Then
      intMax = intMax + 1 ' set a new value for the array
      Load WSSMTP(intMax) ' load the new winsock control
      WSSMTP(intMax).LocalPort = 25 ' set the port for the new control
      WSSMTP(intMax).Accept requestID ' accept the request
      
      ' now start the POP3 Session by telling the Client that
      ' the server is reay for whatever it wants to try and do
      WSSMTP(intMax).SendData "220" & vbCrLf ': Ready For Mail" & vbCrLf 'PipPop POP3 Server Ver. 1.0.0 Ready" & vbCrLf

End If

' just display some information for the users
' this is not needed but can be used for a log file!
' I didnt setup logging in this version I will do that in the next version
Text2.Text = "CLIENT CONNECTING TO SMTP " & requestID & vbCrLf

End Sub

Private Sub WSSMTP_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim inBuff As String
' SMTP Mail Recieve
    
        'Recieve data from server
        WSSMTP(Index).GetData inBuff, vbString
        'Show data to user
        Text2.Text = inBuff & Chr(13) & Chr(10)
        
        
    Select Case Left(inBuff, 4)
    Case "HELO"
    WSSMTP(Index).SendData "250 PipPop POP3-SMTP Server" & vbCrLf 'Server Ver1.0" & vbCrLf
    Case "MAIL" ' "FROM:"
    Dim iLen As Integer
    
    
    'ilen = len(instr(1,
    
    sMailFrom = Mid(inBuff, 13, Len(inBuff))
    sMailFrom = Left(sMailFrom, (InStr(2, sMailFrom, ">", vbTextCompare) - 1))
    WSSMTP(Index).SendData "250 OK" & vbCrLf
    Case "RCPT" ' TO:"
    sMailTo = Mid(inBuff, 11, Len(inBuff))
    sMailTo = Left(sMailTo, (InStr(1, sMailTo, ">", vbTextCompare) - 1))
    
    Data2.RecordSource = "SELECT * FROM users WHERE email='" & sMailTo & "'"
    Data2.Refresh
    Data2.UpdateControls
    
    If Data2.Recordset.EOF = True And Data2.Recordset.BOF = True Then
    WSSMTP(Index).SendData "550 User Not Here" & vbCrLf
    Else
    WSSMTP(Index).SendData "250 OK User has mailbox here" & vbCrLf
    End If
    
    Case "DATA"
    WSSMTP(Index).SendData "354" & vbCrLf ' Ready For Mail" & vbCrLf
    
    Case "QUIT"
    WSSMTP(Index).SendData "221" & vbCrLf
    WSSMTP(Index).Close
    Unload WSSMTP(Index)
    Case vbCrLf & vbCrLf & "." & vbCrLf & vbCrLf
    
    Case ""
    
    Case Else
    Dim iUid As Integer
    
    If Mid(inBuff, 3, 1) = "." Then
    WSSMTP(Index).SendData "250" & vbCrLf
    Else
    mailBody = mailBody & inBuff
    
    Data2.RecordSource = "SELECT * FROM users WHERE email='" & sMailTo & "'"
    Data2.Refresh
    Data2.UpdateControls
    
    If Data2.Recordset.EOF = True And Data2.Recordset.BOF = True Then
    ' there is an error - we cannot find the email address
    Else
    Data2.Recordset.MoveFirst
    iUid = Data2.Recordset.Fields("userid")
    
    Data2.RecordSource = "SELECT * FROM mail"
    Data2.Refresh
    Data2.UpdateControls
    
    Data2.Recordset.AddNew
    Data2.Recordset.Fields("userid") = iUid
    Data2.Recordset.Fields("message") = mailBody
    Data2.Recordset.Fields("octets") = Len(mailBody)
    Data2.Recordset.Update
    End If

    End If
    End Select
End Sub
