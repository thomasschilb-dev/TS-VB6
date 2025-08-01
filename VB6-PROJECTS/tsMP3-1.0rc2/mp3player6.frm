VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00400040&
   Caption         =   "Harddrive Search"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4455
   Icon            =   "mp3player6.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   2160
      Top             =   4680
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Exit"
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Minimize"
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
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Add All"
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
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0C0&
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Stop"
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
      TabIndex        =   7
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Search"
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
      TabIndex        =   6
      Top             =   2280
      Width           =   735
   End
   Begin VB.ListBox List1 
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
      Height          =   885
      Left            =   240
      TabIndex        =   5
      Top             =   3000
      Width           =   3975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00400040&
      Caption         =   "MP3's Found"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   1695
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   4215
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
      ForeColor       =   &H00FFC0C0&
      Height          =   2190
      Left            =   1920
      Pattern         =   "*.mp3"
      TabIndex        =   3
      Top             =   360
      Width           =   2295
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
      ForeColor       =   &H00FFC0C0&
      Height          =   1350
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1695
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
      ForeColor       =   &H00FFC0C0&
      Height          =   270
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400040&
      Caption         =   "Search"
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
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SearchFlag As Integer

Private Function DirDiver(NewPath As String, DirCount As Integer, BackUp As String) As Integer

Static FirstErr As Integer
Dim DirsToPeek As Integer, AbandonSearch As Integer, ind As Integer
Dim OldPath As String, ThePath As String, entry As String
Dim retval As Integer
    SearchFlag = True
    DirDiver = False
    retval = DoEvents()
    If SearchFlag = False Then
        DirDiver = True
        Exit Function
    End If
    On Local Error GoTo DirDriverHandler
    DirsToPeek = Dir1.ListCount                  '
    Do While DirsToPeek > 0 And SearchFlag = True
        OldPath = Dir1.Path
        Dir1.Path = NewPath
        If Dir1.ListCount > 0 Then
            ' Get to the
            Dir1.Path = Dir1.list(DirsToPeek - 1)
            AbandonSearch = DirDiver((Dir1.Path), DirCount%, OldPath)
        End If
        
        DirsToPeek = DirsToPeek - 1
        If AbandonSearch = True Then Exit Function
    Loop
    
    If File1.ListCount Then
        If Len(Dir1.Path) <= 3 Then
            ThePath = Dir1.Path
        Else
            ThePath = Dir1.Path + "\"
        End If
        For ind = 0 To File1.ListCount - 1
            entry = ThePath + File1.list(ind)
            List1.AddItem entry
            Frame2.Caption = "Mp3's Found: " & List1.ListCount
        Next ind
    End If
    If BackUp <> "" Then
        Dir1.Path = BackUp
    End If
    Exit Function
DirDriverHandler:
    If Err = 7 Then
        MsgBox "You've filled the list box. Abandoning search..."
        Exit Function
    Else
        MsgBox Error
        End
    End If
End Function

Private Sub Command1_Click()
Dim FirstPath As String, DirCount As Integer, NumFiles As Integer
Dim result As Integer
  
    If Dir1.Path <> Dir1.list(Dir1.ListIndex) Then
        Dir1.Path = Dir1.list(Dir1.ListIndex)
        Exit Sub
    End If

    FirstPath = Dir1.Path
    DirCount = Dir1.ListCount

    result = DirDiver(FirstPath, DirCount, "")
    File1.Path = Dir1.Path
   
End Sub

Private Sub Command4_Click()
Dim a As String
If List1 = "" Then Exit Sub
a$ = GetLastBackSlash(List1)
Form1.List1.AddItem a$
Form1.List2.AddItem List1
Form1.Frame3.Caption = "MP3 Playlist " & Form1.List1.ListCount & " Songs"
End Sub

Private Sub Command5_Click()
Dim b
Dim c As String
Dim d As String
On Error Resume Next
For b = 0 To List1.ListCount - 1
c$ = GetLastBackSlash(List1.list(b))
Form1.List1.AddItem c$
Form1.List2.AddItem List1.list(b)
Next b
Form1.Frame3.Caption = "MP3 Playlist " & Form1.List1.ListCount & " Songs"
End Sub

Private Sub Command6_Click()
Form7.WindowState = 1
End Sub

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2

End Sub

Private Sub List1_DblClick()
Dim a As String
If List1 = "" Then Exit Sub
a$ = GetLastBackSlash(List1)
Form1.List1.AddItem a$
Form1.List2.AddItem List1
Form1.Frame3.Caption = "MP3 Playlist " & Form1.List1.ListCount & " Songs"

End Sub

