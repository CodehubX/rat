VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Çalýþan Uygulamalar"
   ClientHeight    =   4560
   ClientLeft      =   14610
   ClientTop       =   2085
   ClientWidth     =   7575
   FillColor       =   &H0000FF00&
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   7575
   Begin VB.CommandButton Command3 
      Caption         =   "Çalýþtýr"
      Height          =   495
      Left            =   5880
      TabIndex        =   5
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Text            =   "Ýþlemin Ýsmini Yazýnýz."
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Text            =   "Ýþlemin Ýsmini Yazýnýz."
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Kapat"
      Height          =   495
      Left            =   5880
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Yenile"
      Height          =   495
      Left            =   5880
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000007&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#.##0,00 ""TL"""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1055
         SubFormatType   =   0
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   4545
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
List1.Clear
yenile
End Sub

Public Sub yenile()
Form1.Winsock1.SendData "gorev_yenile"
End Sub

Private Sub Command2_Click()
Form1.Winsock1.SendData "gorev_kapat|" + Text1.Text
End Sub

Private Sub Command3_Click()
Form1.Winsock1.SendData "gorev_calistir|" + Text2.Text
End Sub

Private Sub Form_Load()
Form1.Winsock1.SendData "gorev"
End Sub

Private Sub List1_Click()
Text1.Text = List1.Text
End Sub
