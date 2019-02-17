VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Sity Bank"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   10395
   Icon            =   "siti2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8520
   ScaleWidth      =   10395
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Clear &Data"
      Height          =   495
      Left            =   8400
      TabIndex        =   12
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2520
      TabIndex        =   6
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Back"
      Height          =   495
      Left            =   7200
      TabIndex        =   10
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add Account"
      Height          =   495
      Left            =   6000
      TabIndex        =   9
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   1920
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox Text9 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2520
      TabIndex        =   7
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2520
      TabIndex        =   8
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "sitybank.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7200
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "siti2.frx":0442
      Left            =   2520
      List            =   "siti2.frx":044C
      TabIndex        =   4
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Opening Balance"
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
      Left            =   360
      TabIndex        =   20
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Deposit"
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
      Left            =   360
      TabIndex        =   19
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Type of Account"
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
      Left            =   360
      TabIndex        =   18
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Phone No"
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
      Left            =   360
      TabIndex        =   17
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Address"
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
      Left            =   360
      TabIndex        =   16
      Top             =   1920
      Width           =   1815
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
      Left            =   360
      TabIndex        =   15
      Top             =   1320
      Width           =   1815
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
      Left            =   360
      TabIndex        =   14
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "Total Cheque Issued"
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
      Left            =   360
      TabIndex        =   13
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label10 
      Caption         =   "Date"
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
      Left            =   360
      TabIndex        =   11
      Top             =   5760
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
If Combo1.Text = "CA" Then
Text9.Text = 15
Else
Text9.Text = 10
End If
End Sub

Private Sub Command1_Click()
Data1.RecordSource = "select * from bank where account_no='" + Text1.Text + "'"
Data1.Refresh
Do While Not Data1.Recordset.EOF
c = c + 1
Data1.Recordset.MoveNext
Loop
If c <> 0 Then
MsgBox " Duplicate Account no! Change it"
Else
Data1.Recordset.AddNew
Data1.Recordset.Fields(0) = Text1.Text
Data1.Recordset.Fields(1) = Text2.Text
Data1.Recordset.Fields(2) = Text3.Text
Data1.Recordset.Fields(3) = Text4.Text
Data1.Recordset.Fields(4) = Combo1.Text
Data1.Recordset.Fields(5) = Val(Text6.Text)
Data1.Recordset.Fields(6) = Val(Text7.Text)
Data1.Recordset.Fields(7) = Val(Text9.Text)
Data1.Recordset.Fields(8) = CDate(Text10.Text)
Data1.Recordset.Update
Data1.Recordset.Close
End If
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text6.Text = ""
Text7.Text = ""
Text9.Text = ""
Text10.Text = ""

End Sub

Private Sub Command2_Click()
Form2.Hide
Form1.Show
Form1.Data1.Refresh
End Sub

Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text6.Text = ""
Text7.Text = ""
Text9.Text = ""
Text10.Text = ""
End Sub

Private Sub Form_Load()
Text10.Text = Date
End Sub

Private Sub Text7_GotFocus()
Text7.Text = Text6.Text
End Sub

