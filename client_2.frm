VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form2 
   Caption         =   "Server Bul Durum : Aranýyor..."
   ClientHeight    =   1890
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5985
   LinkTopic       =   "Form2"
   ScaleHeight     =   1890
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3960
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   1190
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4440
      Top             =   1440
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Bulunan Server'a Baðlan"
      Height          =   495
      Left            =   4080
      TabIndex        =   6
      Top             =   480
      Width           =   1575
   End
   Begin VB.Timer Timer3 
      Interval        =   30000
      Left            =   4920
      Top             =   1440
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   5400
      Top             =   1440
   End
   Begin VB.Frame Frame1 
      Caption         =   "Bulunan Server"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin VB.Label Label4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Port Numarasý : "
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Host Name :"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "IP Numarasý : "
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
If Winsock1.State <> 7 Then
MsgBox "Server Bulunamadý!", vbCritical, "Client Hata!"
Else
Form1.Winsock1.RemoteHost = Label2.Caption
Form1.Winsock1.RemotePort = Label6.Caption
Form1.Winsock1.Connect
Unload Form2
End If
End Sub

Private Sub Form_Load()
Winsock1.Close
Winsock1.LocalPort = 1190
Winsock1.Listen
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Winsock1.Close
Winsock1.LocalPort = 1190
Winsock1.Listen
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
Form2.Height = 2460
Form2.Width = 6225
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
If Winsock1.State <> 7 Then
Me.Caption = "Server Bul : Zaman Aþýmýna Uðradý"
End If
End Sub


Private Sub Winsock1_Close()
Winsock1.Close
Winsock1.LocalPort = 1190
Winsock1.Listen
End Sub


Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Winsock1.Close
Winsock1.Accept requestID
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim receive As String
Dim Vector() As String
Winsock1.GetData receive
Vector() = Split(receive, "|")
Select Case Vector(0)
Case "ip"
Label2.Caption = Vector(1)
Case "host"
Label4.Caption = Vector(1)
Case "port"
Label6.Caption = Vector(1)
End Select
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
Winsock1.Close
Winsock1.LocalPort = 1190
Winsock1.Listen
Me.Caption = "Server Bul : Hata Meydana Geldi!"
End Sub
