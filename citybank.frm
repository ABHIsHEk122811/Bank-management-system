VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   Caption         =   "Sity Bank"
   ClientHeight    =   6795
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11880
   Icon            =   "citybank.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   10080
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "sitytrans.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6240
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   10080
      TabIndex        =   7
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   10080
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   10080
      TabIndex        =   5
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   10080
      TabIndex        =   4
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   10080
      TabIndex        =   3
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add Transaction"
      Height          =   495
      Left            =   10200
      TabIndex        =   2
      Top             =   5640
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   10080
      TabIndex        =   1
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "citybank.frx":0442
      Height          =   3615
      Left            =   360
      OleObjectBlob   =   "citybank.frx":0456
      TabIndex        =   0
      Top             =   960
      Width           =   7575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "sitybank.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5400
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label Label1 
      Caption         =   "Account No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   15
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   14
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   13
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Type of transaction"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   12
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Cheque No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   11
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Balance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   10
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Date of transaction"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   9
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Menu AccountD 
      Caption         =   "&Account Details"
      Begin VB.Menu New 
         Caption         =   "&New Account"
         Shortcut        =   ^N
      End
      Begin VB.Menu Delete 
         Caption         =   "&Delete Account"
         Shortcut        =   ^D
      End
      Begin VB.Menu Refresh 
         Caption         =   "&Refresh"
         Shortcut        =   ^R
      End
      Begin VB.Menu Issue 
         Caption         =   "&Issue Cheque"
         Shortcut        =   ^I
      End
      Begin VB.Menu Exit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu Trans 
      Caption         =   "Tran&saction"
      Begin VB.Menu withdrawl 
         Caption         =   "&Withdrawl"
         Shortcut        =   ^W
      End
      Begin VB.Menu Deposit 
         Caption         =   "Dep&osit"
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu Search 
      Caption         =   "Sear&ch"
      Begin VB.Menu Account 
         Caption         =   "By &Account No"
         Shortcut        =   ^A
      End
      Begin VB.Menu Name 
         Caption         =   "By Na&me"
         Shortcut        =   ^M
      End
      Begin VB.Menu list 
         Caption         =   "By acc. &type"
         Shortcut        =   ^T
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Account_Click()
a = InputBox("Enter your Account No", "Account")
Data1.RecordSource = "select * from bank where account_no='" + a + "'"
Data1.Refresh
End Sub

Private Sub Command1_Click()
If Text5.Text = "" Then
MsgBox "Please enter the cheque no", vbExclamation, "Cheque"
Else
Data2.RecordSource = "select * from transaction"
Data2.Refresh
Data2.Recordset.AddNew
Data2.Recordset.Fields(0) = Text1.Text
Data2.Recordset.Fields(1) = Text2.Text
Data2.Recordset.Fields(2) = Val(Text3.Text)
Data2.Recordset.Fields(3) = Text4.Text
Data2.Recordset.Fields(4) = Text5.Text
Data2.Recordset.Fields(5) = Val(Text6.Text)
Data2.Recordset.Fields(6) = CDate(Text7.Text)
Data2.Recordset.Update
Data2.Recordset.Close
Data2.RecordSource = "select * from transaction"
Data2.Refresh
Form1.Hide
Form3.Show
End If
End Sub


Private Sub Delete_Click()
On Error GoTo errorhandle
u = InputBox("Enter password", "Password")
If u = "Sitybank" Then
d = InputBox("Enter the account no to delete", "Delete")
On Error GoTo errorhandle
Data1.RecordSource = "select * from bank where Account_No='" + d + "'"
Data1.Refresh
Data1.Recordset.Delete
Data1.Recordset.Close
Data1.RecordSource = "select * from bank"
Data1.Refresh
Else
MsgBox "You are not a valid user", vbExclamation, "Verify"
Data1.RecordSource = "select * from bank"
Data1.Refresh
End If
Exit Sub
errorhandle:
MsgBox "Error occurred! Please fill the input properly", vbExclamation, "Error"
End Sub

Private Sub Deposit_Click()
On Error GoTo errorhandle
Text1.Visible = True
Text2.Visible = True
Text3.Visible = True
Text4.Visible = True
Text5.Visible = True
Text6.Visible = True
Text7.Visible = True
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
Label7.Visible = True
Command1.Visible = True
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
a = InputBox("Enter your Account no", "Account No")
Data1.RecordSource = "select * from bank where account_no='" + a + "'"
Data1.Refresh
d = InputBox("Enter the amount you want to deposit", "Deposit")
Data1.Recordset.Edit
Data1.Recordset.Fields(5) = Data1.Recordset.Fields(5) + Val(d)
Data1.Recordset.Fields(6) = Data1.Recordset.Fields(6) + Val(d)
Text2.Text = Data1.Recordset.Fields(1)
Text6.Text = Data1.Recordset.Fields(6)
Data1.Recordset.Update
Data1.Recordset.Close
Text1.Text = a
Text3.Text = d
Text4.Text = "Deposit"
Exit Sub
errorhandle:
MsgBox "Error occurred! Please fill the input properly", vbExclamation, "Error"

End Sub

Private Sub Exit_Click()
End
End Sub

Private Sub Form_Load()
Data1.RecordSource = "select * from bank"
Data1.Refresh
Text7.Text = Date
End Sub

Private Sub Issue_Click()
On Error GoTo errorhandle
a = InputBox("Enter your Account no", "Account No")
Data1.RecordSource = "select * from bank where account_no='" + a + "'"
Data1.Refresh
i = InputBox("Enter the no of cheque issued", "Issue Cheque")
Data1.Recordset.Edit
Data1.Recordset.Fields(7) = Data1.Recordset.Fields(7) + Val(i)
Data1.Recordset.Update
Data1.Recordset.Close
errorhandle:
MsgBox "Error occurred! Please fill the input properly", vbExclamation, "Error"
End Sub

Private Sub list_Click()
On Error GoTo errorhandle
a = InputBox("Enter the type of account", "Account Type")
Data1.RecordSource = "select * from bank where type_of_acc='" + a + "'"
Data1.Refresh
Exit Sub
errorhandle:
MsgBox "Error occurred! Please fill the input properly", vbExclamation, "Error"
End Sub

Private Sub Name_Click()
On Error GoTo errorhandle
n = InputBox("Enter your Name", "Name")
Data1.RecordSource = "select * from bank where name='" + n + "'"
Data1.Refresh
Exit Sub
errorhandle:
MsgBox "Error occurred! Please fill the input properly", vbExclamation, "Error"
End Sub

Private Sub New_Click()
Form1.Hide
Form2.Show

End Sub

Private Sub Refresh_Click()
Data1.RecordSource = "select * from bank"
Data1.Refresh
Text7.Text = Date

End Sub

Private Sub withdrawl_Click()
On Error GoTo errorhandle
Text1.Visible = True
Text2.Visible = True
Text3.Visible = True
Text4.Visible = True
Text5.Visible = True
Text6.Visible = True
Text7.Visible = True
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
Label7.Visible = True
Command1.Visible = True
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
a = InputBox("Enter your Account no", "Account No")
Data1.RecordSource = "select * from bank where account_no='" + a + "'"
Data1.Refresh
w = InputBox("Enter the amount you want to withdraw", "Withdraw")
If Val(w) > Data1.Recordset.Fields(6) Then
MsgBox " Please check your balance!", vbExclamation, "Balance"
Data1.RecordSource = "select * from bank"
Data1.Refresh
Else
Data1.Recordset.Edit
Data1.Recordset.Fields(6) = Data1.Recordset.Fields(6) - Val(w)
If Data1.Recordset.Fields(7) < 5 Then
MsgBox "Please issue the A/C holder fresh cheque", vbExclamation, "Cheque"
Else
Data1.Recordset.Fields(7) = Data1.Recordset.Fields(7) - 1
Text2.Text = Data1.Recordset.Fields(1)
Text6.Text = Data1.Recordset.Fields(6)
Data1.Recordset.Update
Data1.Recordset.Close
Text1.Text = a
Text3.Text = w
Text4.Text = "Withdrawl"
End If
End If
Exit Sub
errorhandle:
MsgBox "Error occurred! Please fill the input properly", vbExclamation, "Error"
End Sub
