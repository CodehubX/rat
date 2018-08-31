VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IP BULUNAMADI - (DigiRev)"
   ClientHeight    =   1170
   ClientLeft      =   -8460
   ClientTop       =   6990
   ClientWidth     =   5265
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   5265
   Begin VB.ListBox lstIP 
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   4455
   End
   Begin VB.CheckBox chkRetAll 
      Caption         =   "Return all IP addresses found."
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   3375
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4200
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Go"
      Height          =   315
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.ComboBox cmbURL 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   720
      List            =   "frmMain.frx":0013
      Sorted          =   -1  'True
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label lblStatus 
      Caption         =   "Ready."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   4455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   120
      X2              =   4560
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   4560
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "IPs found:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "URL:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   405
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Get External IP Address
'Author: Danny Elkins/DigiRev
'Created: May 19th, 2007
'Last updated: May 19th, 2007

'This project uses the Inet control to get the HTML source code of the webpage.
'Feel free to use any method you like, I wanted to keep the source code as clean as possible.

Option Explicit

Private strHTML As String 'HTML returned from server.
Private colIPs As Collection 'Collection containing all IPs found.

Private Sub cmdGo_Click()
    'Check user input.
    cmbURL.Text = Trim$(cmbURL.Text)
    
    If Len(cmbURL.Text) > 0 Then
        lblStatus.Caption = "Getting data..."
        
        'Cancel inet if it's still doing something.
        If Inet1.StillExecuting Then Inet1.Cancel
        
        'Get the HTML source code to the webpage.
        strHTML = Inet1.OpenURL("http://showip.net/")
        lstIP.Clear
        
        'Check if the server returned any data.
        If Len(strHTML) > 0 Then
            lblStatus.Caption = "Extracting IPs..."
            'Extract the IPs.
            Set colIPs = ExtractIPs(strHTML, chkRetAll.Value)
            
            'Check if we extracted any IPs.
            If colIPs.Count > 0 Then
                DisplayIPs
            End If
            
            lblStatus.Caption = colIPs.Count & IIf(colIPs.Count, " IP found.", " IPs found.")
        Else
            lblStatus.Caption = "Error!"
            MsgBox "Unable to get HTML from webpage! Make sure you typed the URL correctly and that you are connected to the internet!", vbExclamation
        End If
        
    End If
    
End Sub

'Just adds the colIPs collection to the ListBox.
Private Sub DisplayIPs()
    Dim intLoop As Integer
    
    With lstIP
        .Clear
        
        For intLoop = 1 To colIPs.Count
            .AddItem colIPs.Item(intLoop)
            Me.Caption = colIPs.Item(intLoop)
        Next intLoop
    
    End With
    
End Sub

Private Sub Form_Load()
    'Check user input.
    cmbURL.Text = Trim$(cmbURL.Text)
    
    If Len(cmbURL.Text) > 0 Then
        lblStatus.Caption = "Getting data..."
        
        'Cancel inet if it's still doing something.
        If Inet1.StillExecuting Then Inet1.Cancel
        
        'Get the HTML source code to the webpage.
        strHTML = Inet1.OpenURL("http://showip.net/")
        lstIP.Clear
        
        'Check if the server returned any data.
        If Len(strHTML) > 0 Then
            lblStatus.Caption = "Extracting IPs..."
            'Extract the IPs.
            Set colIPs = ExtractIPs(strHTML, chkRetAll.Value)
            
            'Check if we extracted any IPs.
            If colIPs.Count > 0 Then
                DisplayIPs
            End If
            
            lblStatus.Caption = colIPs.Count & IIf(colIPs.Count, " IP found.", " IPs found.")
        Else
            lblStatus.Caption = "Error!"
            MsgBox "Unable to get HTML from webpage! Make sure you typed the URL correctly and that you are connected to the internet!", vbExclamation
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set colIPs = Nothing
End Sub
