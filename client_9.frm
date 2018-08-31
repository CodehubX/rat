VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Programlar"
   ClientHeight    =   2160
   ClientLeft      =   1260
   ClientTop       =   2820
   ClientWidth     =   5280
   FillColor       =   &H0000FF00&
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   5280
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Caption         =   "Yararlý Programlar"
      ForeColor       =   &H0000FF00&
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.CommandButton Command1 
         Caption         =   "Web Tarayýcýsý"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
web_tarayýcý.Show
Unload Form9
End Sub
