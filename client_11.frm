VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ekraný Kapat"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5655
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      Caption         =   "Ekraný Kapat"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   360
      MaskColor       =   &H0000FF00&
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      Caption         =   "Ekranda Mesaj Göster"
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   3480
      MaskColor       =   &H0000FF00&
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   "Ekranýnýz Yöneticiniz Tarafýndan Kapatýldý"
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      Caption         =   "Ekranda Mesaj Göster"
      ForeColor       =   &H0000FF00&
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Kapalý Ekranda Gözükecek Mesajý Yazýn"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   3135
      End
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
Form1.Winsock1.SendData "ekran_mesaj|" + "acik"
Else
Form1.Winsock1.SendData "ekran_mesaj|" + "kapali"
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
Text1.Enabled = True
Check1.Enabled = True
Form1.Winsock1.SendData "ekran_kapat"
Else
Check1.Enabled = False
Text1.Enabled = False
Form1.Winsock1.SendData "ekran_ac"
End If
End Sub

Private Sub Form_Load()
Text1.Enabled = False
Check1.Enabled = False
End Sub

Private Sub Text1_Change()
Form1.Winsock1.SendData "ekran_mesajý|" + Text1.Text
End Sub
