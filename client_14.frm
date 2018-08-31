VERSION 5.00
Begin VB.Form Form14 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ýnternet"
   ClientHeight    =   870
   ClientLeft      =   9765
   ClientTop       =   7245
   ClientWidth     =   3510
   LinkTopic       =   "Form14"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   3510
   Begin VB.CommandButton Command1 
      Caption         =   "Dakika Sonra Aç"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   405
      Left            =   120
      TabIndex        =   1
      Text            =   "5"
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Ýnternetin Kaç Dakika Sonra Açýlacaðýný Yazýn"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Form1.Winsock1.SendData "internet_dakika|" + Text1.Text
End Sub
