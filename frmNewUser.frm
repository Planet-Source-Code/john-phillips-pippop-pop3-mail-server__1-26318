VERSION 5.00
Begin VB.Form frmNewUser 
   Caption         =   "New Mailbox"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   ScaleHeight     =   3285
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1200
      TabIndex        =   13
      Top             =   2280
      Width           =   4335
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2760
      Width           =   1140
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   11
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1200
      TabIndex        =   9
      Top             =   1800
      Width           =   4335
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Top             =   1320
      Width           =   4335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   960
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   480
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label4 
      Caption         =   "Server"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Full Name"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Re-Enter"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Password"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Email Address"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "User Name"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmNewUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text4.Text <> Text3.Text Then
MsgBox "The password does not match"
Text3.SetFocus
Exit Sub
End If

If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "All fields must be filled out!"
Exit Sub
End If

If Combo1.Text = "" Then
MsgBox "You must select a server!"
Exit Sub
End If

Data1.Recordset.AddNew
Data1.Recordset.Fields("username") = Text1.Text
Data1.Recordset.Fields("userpassword") = Text4.Text
Data1.Recordset.Fields("email") = Text2.Text
If Text5.Text <> "" Then
Data1.Recordset.Fields("fullname") = Text5.Text
End If
Data1.Recordset.Fields("serverid") = Combo1.ItemData(Combo1.ListIndex)
Data1.Recordset.Update

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""

Text1.SetFocus

MsgBox "User Added"
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\pippop.mdb"

Data1.RecordSource = "SELECT * FROM server"
Data1.Refresh
Data1.UpdateControls

If Data1.Recordset.EOF = True And Data1.Recordset.BOF = True Then
MsgBox "No servers to add to"
Exit Sub
Unload Me
End If

Data1.Recordset.MoveFirst

Do While Data1.Recordset.EOF = False
Combo1.AddItem Data1.Recordset.Fields("domainname")
Combo1.ItemData(Combo1.NewIndex) = Data1.Recordset.Fields("serverid")
Data1.Recordset.MoveNext
Loop

Combo1.ListIndex = 0

Data1.RecordSource = "SELECT * FROM users"
Data1.Refresh
Data1.UpdateControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmAdmin.SetFocus
End Sub
