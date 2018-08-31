VERSION 5.00
Begin VB.Form hakkinda 
   BackColor       =   &H80000007&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Hakkýnda"
   ClientHeight    =   2895
   ClientLeft      =   1035
   ClientTop       =   1815
   ClientWidth     =   5010
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Hata anýnda ulaþabileceðiniz website adresi : "
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
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   4935
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "http://<some-random-website>/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "<some-random-email>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Hata anýnda ulaþabileceðiniz e-posta adresi : "
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
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   4935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "1.4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Yazýlým Sürümü : "
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
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bu yazýlým  Nexus Tarafýndan yapýlmýþtýr. Kopyalanamaz , Taklit Edilemez."
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
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "hakkinda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label6_Click()
web_tarayýcý.Show
web_tarayýcý.Combo1.Text = "http://<some-random-website>/"
web_tarayýcý.git
End Sub
