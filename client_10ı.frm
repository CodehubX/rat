VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form web_tarayýcý 
   BackColor       =   &H00000000&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Nexus's Web Browser"
   ClientHeight    =   8370
   ClientLeft      =   6435
   ClientTop       =   4305
   ClientWidth     =   14805
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   14805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "Tam Ekran"
      Height          =   375
      Left            =   9000
      TabIndex        =   5
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<Ýleri>"
      Height          =   375
      Left            =   8160
      TabIndex        =   4
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<Geri>"
      Height          =   375
      Left            =   7320
      TabIndex        =   3
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Git"
      Height          =   375
      Left            =   6480
      TabIndex        =   2
      Top             =   0
      Width           =   615
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Text            =   "http://<some-random-website>/"
      Top             =   0
      Width           =   6375
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3840
      Top             =   2040
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   8055
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   14775
      ExtentX         =   26061
      ExtentY         =   14208
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "web_tarayýcý"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
WebBrowser1.Navigate (Combo1.Text)
End Sub

Private Sub Command1_Click()
WebBrowser1.Navigate (Combo1.Text)
Combo1.AddItem (Combo1.Text)
End Sub

Public Sub git()
WebBrowser1.Navigate (Combo1.Text)
End Sub

Private Sub Command2_Click()
WebBrowser1.GoBack
End Sub

Private Sub Command3_Click()
WebBrowser1.GoForward
End Sub

Private Sub Command4_Click()
If Command4.Caption = "Tam Ekran" Then
web_tarayýcý.BorderStyle = 0
web_tarayýcý.WindowState = 2
Command4.Caption = "Tam Ekrandan Çýk"
ElseIf Command4.Caption = "Tam Ekrandan Çýk" Then
web_tarayýcý.BorderStyle = 2
web_tarayýcý.WindowState = 0
Command4.Caption = "Tam Ekran"
End If
End Sub

Private Sub Form_Load()
WebBrowser1.Navigate (Combo1.Text)
End Sub

Private Sub Timer1_Timer()
WebBrowser1.Width = web_tarayýcý.Width
WebBrowser1.Height = web_tarayýcý.Height
End Sub

