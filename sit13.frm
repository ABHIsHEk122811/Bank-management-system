VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sity Bank"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   8850
   Icon            =   "sit13.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   8850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "&End"
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   6000
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "sit13.frx":0442
      Height          =   4335
      Left            =   1560
      TabIndex        =   1
      Top             =   1200
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   7646
      _Version        =   393216
      DataMember      =   "Command1_Grouping"
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Return"
      Height          =   495
      Left            =   4200
      TabIndex        =   0
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Transaction Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2640
      TabIndex        =   3
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form3.Hide
Form1.Show
Form1.Text1.Visible = False
Form1.Text2.Visible = False
Form1.Text3.Visible = False
Form1.Text4.Visible = False
Form1.Text5.Visible = False
Form1.Text6.Visible = False
Form1.Text7.Visible = False
Form1.Label1.Visible = False
Form1.Label2.Visible = False
Form1.Label3.Visible = False
Form1.Label4.Visible = False
Form1.Label5.Visible = False
Form1.Label6.Visible = False
Form1.Label7.Visible = False
Form1.Command1.Visible = False
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()

End Sub
