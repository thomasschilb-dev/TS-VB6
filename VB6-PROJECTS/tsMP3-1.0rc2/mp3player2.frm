VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00400040&
   Caption         =   "Wolfie MP3 Player"
   ClientHeight    =   2760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4110
   ForeColor       =   &H0000FF00&
   Icon            =   "mp3player2.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   4110
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H00400040&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   840
      TabIndex        =   5
      Text            =   "Path"
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   495
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00400040&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2040
      Left            =   1920
      Pattern         =   "*.mp3"
      TabIndex        =   2
      Top             =   360
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00400040&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1890
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00400040&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.List1.AddItem File1
Form1.List2.AddItem File1.Path & "\" & File1
Form1.Frame3.Caption = "MP3 Playlist " & Form1.List1.ListCount & " Songs"
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
Text1.text = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1
End Sub

Private Sub File1_Click()
Text1.text = Dir1.Path & "\" & File1
End Sub

Private Sub File1_DblClick()
Dim a As String
a$ = File1
Form1.List1.AddItem a$
Form1.List2.AddItem File1.Path & "\" & File1
Form1.Frame3.Caption = "MP3 Playlist " & Form1.List1.ListCount & " Songs"

End Sub

Private Sub Form_Load()
  Left = (Screen.Width - Width) \ 2
    Top = (Screen.Height - Height) \ 2
Text1.text = Dir1.Path
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Frame3.Caption = "MP3 Playlist " & Form1.List1.ListCount & " Songs"

End Sub
