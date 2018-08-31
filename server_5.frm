VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Remote Control Server"
   ClientHeight    =   2820
   ClientLeft      =   15060
   ClientTop       =   0
   ClientWidth     =   5295
   Icon            =   "server_5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Server Kapat"
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4800
      Top             =   0
   End
   Begin VB.ListBox islem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   1590
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   5295
   End
   Begin VB.Label port 
      BackColor       =   &H80000012&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      Caption         =   "Port :"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label ip 
      BackColor       =   &H80000012&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label host 
      BackColor       =   &H80000012&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "IP :"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "Host Name : "
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "Yapýlan Ýþlem"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
parola.Show
End Sub

Private Sub Form_Load()
App.TaskVisible = False
host.Caption = Form1.Winsock1.LocalHostName
ip.Caption = Form1.Winsock1.LocalIP
port.Caption = Form1.Winsock1.LocalPort
'Form1.Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
topmost = True
End Sub
