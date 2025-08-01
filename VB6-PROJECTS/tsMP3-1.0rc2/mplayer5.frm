VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00400040&
   Caption         =   " Create Text List"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3885
   Icon            =   "mplayer5.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   3885
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
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
      Left            =   2520
      TabIndex        =   8
      Text            =   "7"
      Top             =   3120
      Width           =   1215
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
      Left            =   120
      TabIndex        =   7
      Text            =   "MP3 Playlist"
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Close"
      Height          =   255
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Print"
      Height          =   255
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Save"
      Height          =   255
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Create"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00400040&
      Caption         =   "Control Panel"
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
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   3615
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
      Height          =   2295
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   360
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400040&
      Caption         =   "Create a Text List"
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
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00400040&
      BackStyle       =   0  'Transparent
      Caption         =   "Playlist Font Size"
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
      Left            =   2520
      TabIndex        =   10
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400040&
      BackStyle       =   0  'Transparent
      Caption         =   "Name of Playlist"
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
      TabIndex        =   9
      Top             =   2880
      Width           =   1215
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
Text1.FontSize = Combo1.text
End Sub

Private Sub Combo1_Click()
Text1.FontSize = Combo1.text

End Sub

Private Sub Command1_Click()
Dim a
Dim b
For a = 0 To Form1.List1.ListCount - 1
Text1.SelText = a & ")" & Form1.List1.list(a)
Text1.SelText = vbNewLine
Next a
End Sub

Private Sub Command2_Click()
On Error GoTo error
Open Text2 & ".txt" For Append As 1
' example C:\demo.txt
Print #1, Text1.text
Close 1
MsgBox "Your list has been saved.", , " Wolfie MP3 Player"
Exit Sub
error:  MsgBox Err.Description, , "Error saving playlist"
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2
Combo1.AddItem "6"
Combo1.AddItem "7"
Combo1.AddItem "8"
Combo1.AddItem "9"
Combo1.AddItem "10"
Combo1.AddItem "11"
Combo1.AddItem "12"
End Sub

