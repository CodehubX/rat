VERSION 5.00
Begin VB.Form parola 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Yönetim"
   ClientHeight    =   1155
   ClientLeft      =   15345
   ClientTop       =   6510
   ClientWidth     =   4440
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   4440
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Giriþ"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   240
      MousePointer    =   1  'Arrow
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Parolayý girin..."
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Server'ý kapatmak için parolayý girin..."
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "parola"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "nexus" Then
MsgBox "Server Kapatýldý !", vbCritical, "Server Yönetim"
Unload Form1
Unload Form2
Unload Form3
Unload Form5
Unload frmMain
Unload frmServer
Unload keyloger
Unload parola
Else
MsgBox "Yanlýþ þifre girdiniz.", vbCritical, "Yönetim"
End If
End Sub
