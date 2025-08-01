VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00400040&
   Caption         =   "MP3 Tag Editer"
   ClientHeight    =   2640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6210
   Icon            =   "mp3player7.frx":0000
   LinkTopic       =   "Form8"
   ScaleHeight     =   2640
   ScaleWidth      =   6210
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      BackColor       =   &H00400040&
      ForeColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   4200
      TabIndex        =   16
      Top             =   2160
      Width           =   1815
   End
   Begin VB.FileListBox File1 
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
      Height          =   1740
      Left            =   4200
      Pattern         =   "*.mp3"
      TabIndex        =   15
      Top             =   360
      Width           =   1815
   End
   Begin VB.DirListBox Dir1 
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
      Height          =   1650
      Left            =   2880
      TabIndex        =   14
      Top             =   720
      Width           =   1335
   End
   Begin VB.DriveListBox Drive1 
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
      Left            =   2880
      TabIndex        =   13
      Top             =   360
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00400040&
      Caption         =   "MP3 Files"
      ForeColor       =   &H0000FF00&
      Height          =   2415
      Left            =   2760
      TabIndex        =   12
      Top             =   120
      Width           =   3375
   End
   Begin VB.TextBox Text4 
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
      ForeColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   840
      TabIndex        =   11
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox Text3 
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
      ForeColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   840
      TabIndex        =   10
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox Text2 
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
      ForeColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   840
      TabIndex        =   9
      Top             =   840
      Width           =   1815
   End
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
      ForeColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   840
      TabIndex        =   8
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Write"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
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
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Read"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
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
      Top             =   2160
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400040&
      Caption         =   "Control Panel"
      ForeColor       =   &H0000FF00&
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin VB.Label Label4 
         BackColor       =   &H00400040&
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
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
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00400040&
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
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
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00400040&
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
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
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00400040&
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
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
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

   Dim Filename   As String
   Dim HasTag     As Boolean
   Dim Tag        As String * 3
   Dim Songname   As String * 30
   Dim Artist     As String * 30
   Dim Album      As String * 30
   Dim Year       As String * 4
   Dim Comment    As String * 30
   Dim Genre      As String * 1
   
   Filename = Dir1.Path & File1
   Open Filename For Binary As #1
   Get #1, FileLen(Filename) - 127, Tag
   If Not Tag = "TAG" Then
      Close #1
      HasTag = False
      Exit Sub
   End If
   HasTag = True
   Get #1, , Songname
   Get #1, , Artist
   Get #1, , Album
   Get #1, , Year
   Get #1, , Comment
   Get #1, , Genre
   Close #1

End Sub

Private Sub Command2_Click()

   Dim Filename   As String
   Dim hostage    As Boolean
   Dim Tag        As String * 3
   Dim Songname   As String * 30
   Dim Artist     As String * 30
   Dim Album      As String * 30
   Dim Year       As String * 4
   Dim Comment    As String * 30
   Dim Genre      As String * 1
   
   Filename = Dir1.Path & File1
   Tag = "TAG"
   Songname = "My Song"
   Artist = "My Artist"
   Album = "My Album"
   Year = "1970"
   Comment = "This is my favourite"
   Genre = Chr(12)
   
   Open Filename For Binary Access Write As #1
   Seek #1, FileLen(Filename) - 127
   Put #1, , Tag
   Put #1, , Songname
   Put #1, , Artist
   Put #1, , Album
   Put #1, , Year
   Put #1, , Comment
   Put #1, , Genre
   Close #1
   

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1
End Sub

Private Sub File1_Click()
Text5.text = Dir1.Path & File1
End Sub
