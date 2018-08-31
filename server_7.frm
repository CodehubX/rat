VERSION 5.00
Begin VB.Form saka 
   BorderStyle     =   0  'None
   ClientHeight    =   9030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8115
   LinkTopic       =   "Form6"
   ScaleHeight     =   9030
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   4
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1055
         SubFormatType   =   8
      EndProperty
      DragMode        =   1  'Automatic
      FillStyle       =   5  'Downward Diagonal
      Height          =   8895
      Left            =   360
      Negotiate       =   -1  'True
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      Picture         =   "server_7.frx":0000
      ScaleHeight     =   593
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   473
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4320
      Top             =   2160
   End
End
Attribute VB_Name = "saka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
topmost = True
End Sub
