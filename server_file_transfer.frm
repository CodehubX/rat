VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmClient 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dosya Gönder"
   ClientHeight    =   1410
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Ýlk Gönderimde Servera Klasör Aç"
      Height          =   435
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Gönderilen Çalýþtýr."
      Height          =   435
      Left            =   3600
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock wsTCP 
      Index           =   0
      Left            =   120
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Gönder"
      Height          =   435
      Left            =   2280
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   600
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Dosya Seç"
      Height          =   315
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtFile 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   5175
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================
'Written by Igor Ostrovsky (igor@ostrosoft.com)
'Visual Basic 911 (http://www.ostrosoft.com/vb)
'==============================================
Option Explicit

Dim buffer() As Byte
Dim lBytes As Long
Dim temp As String

Private Sub cmdBrowse_Click()
  dlg.ShowOpen
  txtFile = dlg.FileName
End Sub

Private Sub cmdSend_Click()
  cmdSend.Enabled = False
  lBytes = 0
  ReDim buffer(FileLen(dlg.FileName) - 1)
  Open dlg.FileName For Binary As 1
  Get #1, 1, buffer
  Close #1
  Load wsTCP(1)
  wsTCP(1).RemoteHost = Form1.Text1.Text
  wsTCP(1).RemotePort = 1199
  wsTCP(1).Connect
  lblStatus = "Baðlanýyor..."
End Sub
Public Sub gonder()
cmdSend.Enabled = False
  lBytes = 0
  ReDim buffer(FileLen(dlg.FileName) - 1)
  Open dlg.FileName For Binary As 1
  Get #1, 1, buffer
  Close #1
  Load Form1.Winsock1
  
End Sub


Private Sub Form_Load()
frmClient.Hide
frmClient.Visible = False
txtFile.Text = screen_main.Text1.Text
End Sub


Private Sub wsTCP_DataArrival(Index As Integer, ByVal bytesTotal As Long)
  wsTCP(1).GetData temp
  If InStr(temp, vbCrLf) <> 0 Then temp = Left(temp, InStr(temp, vbCrLf) - 1)
  If temp = "OK" Then
    wsTCP(1).SendData buffer
  Else
    lblStatus = "Something wrong"
    Unload wsTCP(1)
    cmdSend.Enabled = True
  End If
End Sub

Private Sub wsTCP_SendComplete(Index As Integer)
  If temp = "OK" Then
    lblStatus = "Gönderme Tamamlandý"
    temp = ""
    Unload wsTCP(1)
    cmdSend.Enabled = True
  End If
End Sub

Private Sub wsTCP_SendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
  If temp = "OK" Then
    lBytes = lBytes + bytesSent
    lblStatus = lBytes & " out of " & UBound(buffer) & " byte gönderildi"
  End If
End Sub
