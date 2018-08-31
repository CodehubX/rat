VERSION 5.00
Begin VB.Form video_oynat 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Video Oynat"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5025
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   5025
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Video Oynatýcý Kapat"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Video Oynat"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "Serverda oynatýlacak videonun adresini yazýn."
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "video_oynat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Form1.Winsock1.SendData "video_kapat"
End Sub

Private Sub Command2_Click()
On Error Resume Next
Form1.Winsock1.SendData "video_adres|" + Text1.Text
End Sub
