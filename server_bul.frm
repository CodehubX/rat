VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form server_bul 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server Bul"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   3690
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock3 
      Left            =   4320
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3840
      TabIndex        =   14
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Temizle"
      Height          =   375
      Left            =   2400
      TabIndex        =   13
      Top             =   1680
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   3840
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Yeniden Tara"
      Height          =   375
      Left            =   2400
      TabIndex        =   10
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Baðlan"
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   720
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   3840
      Top             =   1200
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Text            =   "0"
      Top             =   720
      Width           =   375
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Height          =   2370
      Left            =   720
      TabIndex        =   6
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox ip4 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Text            =   "1"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox ip3 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Text            =   "2"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox ip2 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Text            =   "168"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox ip1 
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Text            =   "192"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tara"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   3840
      Top             =   120
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2370
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   615
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3840
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Bulunan IP(ler)"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   720
      TabIndex        =   12
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "S.No"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "IP :"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "server_bul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Winsock1.Close
Winsock2.Close
Winsock3.Close
If Command1.Caption = "Tara" Then
Command1.Caption = "Dur"
List1.Clear
List2.Clear
Text1.Text = 0
If ip4.Text = 1 Then
Timer1.Enabled = True
Timer2.Enabled = False
Else
Timer1.Enabled = False
Timer2.Enabled = True
End If

ElseIf Command1.Caption = "Dur" Then
Command1.Caption = "Tara"
Timer1.Enabled = False
Timer2.Enabled = False
End If
End Sub

Private Sub Command2_Click()
If Command2.Caption = "Baðlan" Then
Form1.Winsock1.Close
Text2.Text = List2.Text
If Text2.Text = "" Then
MsgBox "Lütfen IP Seçiniz veya IP Bulunamadý!", vbCritical, "Server Bul"
Else
Form1.Text1.Text = Text2.Text
Form1.baglan
Command2.Caption = "Baðlantýyý Kes"
End If
ElseIf Command2.Caption = "Baðlantýyý Kes" Then
Form1.baglan
Command2.Caption = "Baðlan"
End If
End Sub

Private Sub Command3_Click()
Command1_Click
End Sub

Private Sub Command4_Click()
List1.Clear
List2.Clear
ip1.Text = 192
ip2.Text = 168
ip3.Text = 2
ip4.Text = 1
Text1.Text = 0
End Sub

Private Sub Form_Load()
Text1.Visible = False
Winsock3.Close
Winsock3.LocalPort = 1170
Winsock3.Listen
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Winsock1.Close
If ip4.Text = 255 Then
Command1_Click
MsgBox "Tarama Tamamlandý.", vbInformation, "Server Bul"
End If
ip4.Text = ip4.Text + 1
'Text2.Text = Text2.Text + 1
Winsock1.RemoteHost = ip1.Text & "." & ip2.Text & "." & ip3.Text & "." & ip4.Text
Winsock1.RemotePort = 1169
Winsock1.Connect
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
Winsock1.Close
If ip4.Text = 255 Then
Command1_Click
MsgBox "Tarama Tamamlandý.", vbInformation, "Server Bul"
End If
Winsock2.RemoteHost = ip1.Text & "." & ip2.Text & "." & ip3.Text & "." & ip4.Text
Winsock2.RemotePort = 1169
Winsock2.Connect
End Sub

Private Sub Winsock1_Connect()
Text1.Text = Text1.Text + 1
List1.AddItem Text1.Text
List2.AddItem Winsock1.RemoteHostIP
Winsock1.Close
Beep
End Sub

Private Sub Winsock2_Connect()
Timer2.Enabled = False
Text1.Text = Text1.Text + 1
List1.AddItem Text1.Text
List2.AddItem Winsock2.RemoteHostIP
Command1_Click
Winsock2.Close
Beep
End Sub

Private Sub Winsock3_Close()
On Error Resume Next
Winsock3.Close
Winsock3.LocalPort = 1170
Winsock3.Listen
End Sub

Private Sub Winsock3_Connect()
On Error Resume Next
MsgBox "Uzak Serverdan Baðlantý Ýsteði Geldi", vbInformation, "Server Bul"
End Sub

Private Sub Winsock3_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next
Winsock1.Close
Winsock1.Accept requestID
End Sub

Private Sub Winsock3_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim receive As String
Dim Vector() As String
Winsock1.GetData receive
Vector() = Split(receive, "|")
Select Case Vector(0)
Case "bilgi"
Text1.Text = Text1.Text + 1
List1.AddItem Text1.Text
List2.AddItem Vector(1)
Winsock3.Close
Beep
End Select
End Sub

Private Sub Winsock3_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
Winsock3.Close
Winsock3.LocalPort = 1170
Winsock3.Listen
End Sub
