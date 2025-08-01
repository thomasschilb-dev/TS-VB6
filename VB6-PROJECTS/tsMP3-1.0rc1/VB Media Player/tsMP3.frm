VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H80000000&
   Caption         =   "VB Media Player 1v"
   ClientHeight    =   3165
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   3375
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD1 
      Left            =   1200
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox mp3 
      Height          =   2895
      Left            =   120
      ScaleHeight     =   2835
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuopen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuline1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "Help"
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Resize()
mp3.Move 100, 100, Me.ScaleWidth - 200, Me.ScaleHeight - 200
End Sub

Private Sub mnuabout_Click()
MsgBox "tsMP3-1.0"
End Sub

Private Sub mnuexit_Click()
End
End Sub

Private Sub mnuopen_Click()
CD1.ShowOpen
mp3.FileName = CD1.FileName
End Sub

Private Sub mp3_Click()

End Sub
