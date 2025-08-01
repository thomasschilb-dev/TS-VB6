VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00400040&
   Caption         =   " Wolfie MP3 Player"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2715
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   2715
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   390
      Left            =   600
      Pattern         =   "*.mp3"
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   195
      Left            =   720
      TabIndex        =   3
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   615
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00400040&
      ForeColor       =   &H0000FF00&
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00400040&
      ForeColor       =   &H0000FF00&
      Height          =   1530
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a
Dim i
Form1.List1.clear
Form1.List2.clear
For i = 0 To File1.ListCount - 1
a = File1.list(i)
Form1.List1.AddItem a
Form1.List2.AddItem File1.Path & "\" & a
Next i
Form1.Frame3.Caption = "MP3 Playlist " & Form1.List1.ListCount & " Songs"

Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1
End Sub

Private Sub Form_Load()
    Left = (Screen.Width - Width) \ 2
    Top = (Screen.Height - Height) \ 2

End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Frame3.Caption = "MP3 Playlist " & Form1.List1.ListCount & " Songs"

End Sub
