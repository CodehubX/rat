VERSION 5.00
Begin VB.Form hakkinda 
   BackColor       =   &H80000007&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Hakkýnda"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7065
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "http://<some-random-website>/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Web Sitemiz : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "<some-random-email>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   5415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1.4  ~~Yeni Bazý Özellikler Eklenmiþtir~~"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   1080
      Width           =   5415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"client_hakkinda.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6975
   End
End
Attribute VB_Name = "hakkinda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label5_Click()
web_tarayýcý.Show
web_tarayýcý.Combo1.Text = "http://<some-random-website>/"
web_tarayýcý.git
End Sub
