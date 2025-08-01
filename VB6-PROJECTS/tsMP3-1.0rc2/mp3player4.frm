VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00400040&
   Caption         =   "Wolfie MP3 Player"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3735
   Icon            =   "mp3player4.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      BackColor       =   &H00400040&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   2400
      Left            =   240
      TabIndex        =   7
      Top             =   360
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
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
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Play"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   615
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00400040&
      Caption         =   "Repeat"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   150
      Left            =   2760
      TabIndex        =   4
      Top             =   2880
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00400040&
      Caption         =   "Continuous"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   150
      Left            =   240
      TabIndex        =   2
      Top             =   2880
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.ListBox List2 
      Height          =   450
      Left            =   840
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400040&
      Caption         =   "Favorites Playlist"
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
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.CheckBox Check2 
         BackColor       =   &H00400040&
         Caption         =   "Random"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   135
         Left            =   1440
         TabIndex        =   3
         Top             =   2760
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
List2.ListIndex = List1.ListIndex
On Error Resume Next
Dim cMP3 As New clsMP3time
Dim oStr As String
Dim oStr1 As String
Dim oStr2 As String
Dim tm As Date, t1 As Double
Dim Result As Boolean

Form1.MediaPlayer1.filename = List2

cMP3.filename = List2
Result = cMP3.refresh

oStr = oStr & "" & Format$(cMP3.Time, "hh:mm:ss") 'hh:
oStr1 = oStr1 & "" & (cMP3.BitRate / 1000) & "kb " '& cMP3.Frames & " frames" & vbCrLf
oStr2 = oStr2 & "" & cMP3.Frequency & "hz " '& cMP3.ChannelModeText & vbCrLf
Form1.Text1.text = "Bitrate: " & oStr1
Form1.Text2.text = "Freq. : " & oStr2
Form1.Text3.text = "Time: " & oStr
If Form1.MediaPlayer1.IsDurationValid = False Then
MsgBox "Make sure the song is valid.", , "Little MP3 Player"
Exit Sub
End If
Form1.HScroll1.Max = Form1.MediaPlayer1.Duration
Form1.MediaPlayer1.play
Form1.Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
Dim e
e = DoEvents()
Call SaveListBox(List1, "C:\fav.lst")
Call SaveListBox(List2, "C:\favd.lst")
Unload Me
Form5.Hide
End Sub

Private Sub Form_Load()
Dim b
On Error Resume Next
b = DoEvents()
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2
If FileExists("c:\Favorite.m3u") = False Then Exit Sub
Call LoadPlayList(List2, "C:\Favorite.m3u")
Call LoadListView(List1, List2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim d
d = DoEvents()
Call SaveListBox(List2, "C:\Favorite.m3u")
Unload Me
End Sub

Private Sub List1_Click()
On Error Resume Next

List2.ListIndex = List1.ListIndex
Frame1.Caption = "Favorite MP3's " & List1.ListCount & " Songs"

End Sub

Private Sub List1_DblClick()
List2.ListIndex = List1.ListIndex
On Error Resume Next
Dim cMP3 As New clsMP3time
Dim oStr As String
Dim oStr1 As String
Dim oStr2 As String
Dim tm As Date, t1 As Double
Dim Result As Boolean

Form1.MediaPlayer1.filename = List2

cMP3.filename = List2
Result = cMP3.refresh

oStr = oStr & "" & Format$(cMP3.Time, "hh:mm:ss") 'hh:
oStr1 = oStr1 & "" & (cMP3.BitRate / 1000) & "kb " '& cMP3.Frames & " frames" & vbCrLf
oStr2 = oStr2 & "" & cMP3.Frequency & "hz " '& cMP3.ChannelModeText & vbCrLf
Form1.Text1.text = "Bitrate: " & oStr1
Form1.Text2.text = "Freq. : " & oStr2
Form1.Text3.text = "Time: " & oStr
If Form1.MediaPlayer1.IsDurationValid = False Then
MsgBox "Make sure the song is valid.", , "Little MP3 Player"
Exit Sub
End If
Form1.HScroll1.Max = Form1.MediaPlayer1.Duration
Form1.MediaPlayer1.play
Form1.Timer1.Enabled = True
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If List1 = "" Then Exit Sub
If Button = 2 Then
PopupMenu Form4.other
End If
End Sub
