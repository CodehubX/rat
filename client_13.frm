VERSION 5.00
Begin VB.Form Form13 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Yazdýr"
   ClientHeight    =   825
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5745
   LinkTopic       =   "Form13"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   5745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Yazdýr"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Yazdýrýlacak Metni Yazýn"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Form1.Winsock1.SendData "yazdir|" + Text1.Text
End Sub
