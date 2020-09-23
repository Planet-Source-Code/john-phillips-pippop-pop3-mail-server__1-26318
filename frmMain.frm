VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmAdmin 
   Caption         =   "PipPop POP3 Mail Server - Administration"
   ClientHeight    =   6435
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9135
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   5760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   6180
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10478
            Text            =   "Status:"
            TextSave        =   "Status:"
            Object.ToolTipText     =   "Status Of PipPop Mail Server"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "6:10 PM"
            Object.ToolTipText     =   "Current System Time"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "8/17/01"
            Object.ToolTipText     =   "Current System Date"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   10821
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Server"
      TabPicture(0)   =   "frmMain.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Mailboxs"
      TabPicture(1)   =   "frmMain.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Setup"
      TabPicture(2)   =   "frmMain.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   5655
         Left            =   -74880
         TabIndex        =   4
         Top             =   360
         Width           =   8895
         Begin VB.CheckBox Check4 
            Caption         =   "Start PipPop With Windows"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   5280
            Value           =   1  'Checked
            Width           =   5415
         End
         Begin VB.CheckBox Check3 
            Caption         =   "POP3 Protocal"
            Height          =   255
            Left            =   4440
            TabIndex        =   28
            Top             =   3480
            Value           =   1  'Checked
            Width           =   4335
         End
         Begin VB.TextBox Text4 
            Height          =   735
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   27
            Top             =   4200
            Width           =   8655
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Insert Banner Into All Out Bound Mail"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   3840
            Width           =   3615
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Require Authorization For All Out Bound Mail"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   3480
            Value           =   1  'Checked
            Width           =   3615
         End
         Begin VB.Frame Frame5 
            Caption         =   "Required Information"
            Height          =   1575
            Left            =   3240
            TabIndex        =   19
            Top             =   1680
            Width           =   5535
            Begin VB.ComboBox Combo1 
               Height          =   315
               Left            =   1440
               TabIndex        =   24
               Top             =   840
               Width           =   3615
            End
            Begin VB.CommandButton Command3 
               Caption         =   "..."
               Height          =   285
               Left            =   5160
               TabIndex        =   22
               Top             =   360
               Width           =   255
            End
            Begin VB.TextBox Text3 
               Height          =   285
               Left            =   1440
               TabIndex        =   21
               Top             =   360
               Width           =   3615
            End
            Begin VB.Label Label5 
               Caption         =   "IP Address"
               Height          =   285
               Left            =   120
               TabIndex        =   23
               Top             =   840
               Width           =   1215
            End
            Begin VB.Label Label4 
               Caption         =   "Server Directory"
               Height          =   255
               Left            =   120
               TabIndex        =   20
               Top             =   360
               Width           =   1215
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Type Of Server"
            Height          =   1575
            Left            =   120
            TabIndex        =   15
            Top             =   1680
            Width           =   2895
            Begin VB.TextBox Text2 
               Enabled         =   0   'False
               Height          =   285
               Left            =   240
               TabIndex        =   18
               Top             =   1080
               Width           =   2535
            End
            Begin VB.OptionButton Option6 
               Caption         =   "Alias Domain"
               Height          =   255
               Left            =   240
               TabIndex        =   17
               Top             =   720
               Width           =   1575
            End
            Begin VB.OptionButton Option5 
               Caption         =   "Mail Server"
               Height          =   255
               Left            =   240
               TabIndex        =   16
               Top             =   360
               Value           =   -1  'True
               Width           =   2535
            End
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Cancel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6000
            TabIndex        =   14
            Top             =   5160
            Width           =   1335
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Add New"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7440
            TabIndex        =   13
            Top             =   5160
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   2160
            TabIndex        =   11
            Top             =   960
            Width           =   6615
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Internal (LAN)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6960
            TabIndex        =   9
            Top             =   240
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton Option3 
            Caption         =   "External (internet)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4920
            TabIndex        =   8
            Top             =   240
            Width           =   2055
         End
         Begin VB.Line Line2 
            X1              =   8760
            X2              =   120
            Y1              =   5040
            Y2              =   5040
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "If  hosting on an inetrnal network the .com after the domain name is not required!"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2160
            TabIndex        =   12
            Top             =   1320
            Width           =   6615
         End
         Begin VB.Label Label2 
            Caption         =   "Enter Domain Name:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   960
            Width           =   2415
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   8760
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Label Label1 
            Caption         =   "Where will this server be hosting POP3 Mail Clients?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   4695
         End
      End
      Begin VB.Frame Frame2 
         Height          =   5655
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   8895
         Begin VB.Data Data2 
            Caption         =   "Data2"
            Connect         =   "Access 2000;"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   300
            Left            =   1680
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   5160
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Data Data1 
            Caption         =   "Data1"
            Connect         =   "Access 2000;"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   300
            Left            =   360
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   5160
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.ListBox List1 
            Height          =   4545
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   3375
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Edit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4200
            TabIndex        =   33
            Top             =   5160
            Width           =   1455
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Delete"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5760
            TabIndex        =   32
            Top             =   5160
            Width           =   1455
         End
         Begin VB.CommandButton Command4 
            Caption         =   "New"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7320
            TabIndex        =   31
            Top             =   5160
            Width           =   1455
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   4575
            Left            =   3480
            TabIndex        =   29
            Top             =   240
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   8070
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame1 
         Height          =   5655
         Left            =   -74880
         TabIndex        =   2
         Top             =   360
         Width           =   8895
         Begin VB.OptionButton Option2 
            Caption         =   "External POP3 Mail Server (ie. Internet, WAN)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4440
            TabIndex        =   6
            Top             =   240
            Value           =   -1  'True
            Width           =   4335
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Internal Network POP3 Mail Server (ie. LAN)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   4215
         End
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnunewmailbox 
         Caption         =   "New Mailbox"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnudash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnutools 
      Caption         =   "&Tools"
      Begin VB.Menu mnustartserver 
         Caption         =   "Start Server"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnucontents 
         Caption         =   "&Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnudash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
Data1.Recordset.AddNew
Data1.Recordset.Fields("type") = 1
Data1.Recordset.Fields("domainname") = Text1.Text
Data1.Recordset.Fields("ipaddress") = Combo1.Text
Data1.Recordset.Fields("dir") = Text3.Text
Data1.Recordset.Update

MsgBox "New Server Added!"

End Sub

Private Sub Command4_Click()
Load frmNewUser
frmNewUser.Show , Me

frmNewUser.Text2.Text = "@" & Text1.Text
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then
MsgBox "PipPop Is Already Open!", vbOKOnly + vbInformation, "PipPop Open"
Exit Sub
Unload Me
End If

SSTab1.Tab = 0

Frame1.Visible = True
Frame2.Visible = False
Frame3.Visible = False

Text3.Text = App.Path & "\Mail Server"

ListView1.ColumnHeaders. _
   Add , , "User Name", 1250
ListView1.ColumnHeaders. _
   Add , , "Email", 2750
ListView1.ColumnHeaders. _
   Add , , "Level", 800
ListView1.ColumnHeaders. _
   Add , , "Full Name", 2700
ListView1.View = lvwReport


Data1.DatabaseName = App.Path & "\pippop.mdb"
Data2.DatabaseName = App.Path & "\pippop.mdb"

Data1.RecordSource = "SELECT * FROM server"
Data1.Refresh
Data1.UpdateControls

Text1.Text = Winsock1.LocalHostName
Text2.Text = Winsock1.LocalHostName
Combo1.AddItem Winsock1.LocalIP
Combo1.Text = Winsock1.LocalIP

If Data1.Recordset.EOF = True And Data1.Recordset.BOF = True Then
Dim nResp As Integer
nResp = MsgBox("There are no servers setup!" & vbCrLf & "Would you like to setup a server now?", vbYesNo + vbQuestion, "No Servers Found")
If nResp = 6 Then
SSTab1.Tab = 2
Exit Sub
ElseIf nResp = 7 Then
' do nothing and keep going
End If
Exit Sub
End If

FillMailBox
End Sub

Private Sub Form_Resize()
If Me.WindowState = 1 Then Exit Sub

SSTab1.Top = 0
SSTab1.Left = 0
SSTab1.Height = Me.Height - 1000
SSTab1.Width = Me.Width - 120

Frame1.Top = 360
Frame1.Left = 120
Frame1.Width = SSTab1.Width - 240
Frame1.Height = SSTab1.Height - 480

Frame2.Top = 360
Frame2.Left = 120
Frame2.Width = SSTab1.Width - 240
Frame2.Height = SSTab1.Height - 480

Frame3.Top = 360
Frame3.Left = 120
Frame3.Width = SSTab1.Width - 240
Frame3.Height = SSTab1.Height - 480

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub List1_Click()
Data2.RecordSource = "SELECT * FROM users WHERE serverid=" & List1.ListIndex ' (List1.ListIndex)
Data2.Refresh
Data2.UpdateControls


FillUsers
End Sub

Private Sub mnuexit_Click()
Unload Me
End Sub

Private Sub mnustartserver_Click()
Load frmServer
frmServer.Show , Me
End Sub

Private Sub Option5_Click()
If Option5.Value = True Then
Text2.Enabled = False
End If
End Sub

Private Sub Option6_Click()
If Option6.Value = True Then
Text2.Enabled = True
Text2.SetFocus
End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Select Case SSTab1.Tab
Case 0
Frame1.Visible = True
Frame2.Visible = False
Frame3.Visible = False
Case 1
Frame1.Visible = False
Frame2.Visible = True
Frame3.Visible = False
Case 2
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = True
Case Else

End Select
End Sub

Private Function FillMailBox()

Data1.RecordSource = "SELECT * FROM server"
Data1.Refresh
Data1.UpdateControls

With Data1.Recordset

.MoveFirst

Do While .EOF = False
List1.AddItem .Fields("domainname")
List1.ItemData(List1.NewIndex) = .Fields("serverid")
.MoveNext
Loop

End With
End Function


Private Function FillUsers()
On Error GoTo errNoUser
Dim X As Integer

Data2.Recordset.MoveFirst
X = 1
Do While Data2.Recordset.EOF = False

ListView1.ListItems.Add(X, , Data2.Recordset.Fields("username")).SubItems(1) = Data2.Recordset.Fields("email")
ListView1.ListItems(X).SubItems(2) = Data2.Recordset.Fields("security")
ListView1.ListItems(X).SubItems(3) = Data2.Recordset.Fields("fullname")
X = X + 1
Data2.Recordset.MoveNext
Loop

Exit Function
errNoUser:
If Err.Number = 3021 Then
MsgBox "No users setup!"
Exit Function
End If
End Function

Private Sub StatusBar1_PanelClick(ByVal Panel As MSComctlLib.Panel)

End Sub
