VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "A��l�� Yaz�s� Ayarla"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4590
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "De�i�tir"
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Yaz�y� Sil"
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Default Yap"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Text            =   "A��l�� Mesaj� ��eri�ini Yaz�n"
      Top             =   1440
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "A��l�� Mesaj� Ba�l���n� Yaz�n"
      Top             =   720
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "A��l�� Yaz�s� Ayarla"
      ForeColor       =   &H0000FF00&
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.CommandButton Command5 
         Caption         =   "Serverdan Oku"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "De�i�tir"
         Height          =   255
         Left            =   3480
         TabIndex        =   2
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "A��l�� Yaz�s� Mesaj� : "
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "A��l�� Yaz�s� Ba�l��� : "
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Form1.Winsock1.SendData "acilis_yazisi_baslik|" + Text1.Text
End Sub

Private Sub Command2_Click()
On Error Resume Next
Form1.Winsock1.SendData "acilis_mesaj_sifirla"
End Sub

Private Sub Command3_Click()
On Error Resume Next
Form1.Winsock1.SendData "acilis_kapat"
End Sub

Private Sub Command4_Click()
On Error Resume Next
Form1.Winsock1.SendData "acilis_yazisi_mesaj|" + Text2.Text
End Sub

Private Sub Command5_Click()
MsgBox "Bu �zellik Hen�z Aktif De�ildir.", vbCritical, "Client"
'On Error Resume Next
'Form1.Winsock1.SendData "acilis_oku"
End Sub

Private Sub Form_Load()
On Error Resume Next
Form1.Winsock1.SendData "islem|" + "A�ILI� MESAJI A�ILDI"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Form1.Winsock1.SendData "islem|" + "A�ILI� MESAJI KAPATILDI"
End Sub

Private Sub Text1_Click()
Text1.Text = ""
End Sub

Private Sub Text2_Click()
Text2.Text = ""
End Sub
