VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00400040&
   Caption         =   " Wolfie MP3 Player"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
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
   Icon            =   "mp3player3.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu file 
      Caption         =   "file"
      Begin VB.Menu az 
         Caption         =   "-"
      End
      Begin VB.Menu list 
         Caption         =   "List Options"
         Begin VB.Menu x9 
            Caption         =   "-"
         End
         Begin VB.Menu play 
            Caption         =   "Play"
         End
         Begin VB.Menu a 
            Caption         =   "-"
         End
         Begin VB.Menu clear 
            Caption         =   "Clear List"
         End
         Begin VB.Menu savelist 
            Caption         =   "Save List"
         End
         Begin VB.Menu x8 
            Caption         =   "-"
         End
         Begin VB.Menu remove 
            Caption         =   "Remove"
         End
         Begin VB.Menu dupes 
            Caption         =   "Remove Dupes"
         End
         Begin VB.Menu a3 
            Caption         =   "-"
         End
         Begin VB.Menu rename 
            Caption         =   "Rename Song"
         End
         Begin VB.Menu a1 
            Caption         =   "-"
         End
      End
      Begin VB.Menu a2 
         Caption         =   "-"
      End
      Begin VB.Menu favorite 
         Caption         =   "Favorites"
         Begin VB.Menu add 
            Caption         =   "Add Favorite"
         End
         Begin VB.Menu as 
            Caption         =   "-"
         End
         Begin VB.Menu show 
            Caption         =   "Show Favorites"
         End
      End
      Begin VB.Menu x2 
         Caption         =   "-"
      End
   End
   Begin VB.Menu other 
      Caption         =   "fav"
      Begin VB.Menu z1 
         Caption         =   "-"
      End
      Begin VB.Menu play5 
         Caption         =   "Play"
      End
      Begin VB.Menu z 
         Caption         =   "-"
      End
      Begin VB.Menu rem 
         Caption         =   "Remove"
      End
      Begin VB.Menu clear5 
         Caption         =   "Clear List"
      End
      Begin VB.Menu z3 
         Caption         =   "-"
      End
      Begin VB.Menu rename5 
         Caption         =   "Rename Song"
      End
      Begin VB.Menu v 
         Caption         =   "-"
      End
   End
   Begin VB.Menu option 
      Caption         =   "options"
      Begin VB.Menu v1 
         Caption         =   "-"
      End
      Begin VB.Menu autosave 
         Caption         =   "Auto Save Playlist on Exit"
         Checked         =   -1  'True
      End
      Begin VB.Menu autostart 
         Caption         =   "Auto Load Playlist on Start-Up"
         Checked         =   -1  'True
      End
      Begin VB.Menu x 
         Caption         =   "-"
      End
      Begin VB.Menu createlist 
         Caption         =   "Create Textlist of Playlist"
      End
      Begin VB.Menu x1 
         Caption         =   "-"
      End
      Begin VB.Menu makedefault 
         Caption         =   "Make Default MP3 Player"
      End
      Begin VB.Menu zq 
         Caption         =   "-"
      End
      Begin VB.Menu mini 
         Caption         =   "Minimize to System Tray"
      End
      Begin VB.Menu d 
         Caption         =   "-"
      End
      Begin VB.Menu tag 
         Caption         =   "MP3 Tag Editer"
      End
      Begin VB.Menu q 
         Caption         =   "-"
      End
      Begin VB.Menu search 
         Caption         =   "Search Harddrive for MP3's"
      End
      Begin VB.Menu v2 
         Caption         =   "-"
      End
   End
   Begin VB.Menu color 
      Caption         =   "color"
      Begin VB.Menu v3 
         Caption         =   "-"
      End
      Begin VB.Menu scroll 
         Caption         =   "Backwards Scroll"
      End
      Begin VB.Menu v4 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
    Private Const LB_FINDSTRINGEXACT = &H1A2
Private Function LBDupe(lpBox As ListBox) As Integer
    Dim nCount As Integer, nPos1 As Integer, nPos2 As Integer, nDelete As Integer
    Dim sText As String


    If lpBox.ListCount < 3 Then
        LBDupe = 0
        Exit Function
    End If


    For nCount = 0 To lpBox.ListCount - 1


        Do: DoEvents
                sText = lpBox.list(nCount) 'had To update this line, sorry
                nPos1 = SendMessageByString(lpBox.hwnd, LB_FINDSTRINGEXACT, nCount, sText)
                nPos2 = SendMessageByString(lpBox.hwnd, LB_FINDSTRINGEXACT, nPos1 + 1, sText)
                If nPos2 = -1 Or nPos2 = nPos1 Then Exit Do
                lpBox.RemoveItem nPos2
                nDelete = nDelete + 1
            Loop
        Next nCount
        LBDupe = nDelete
        Form1.Frame3.Caption = "MP3 Playlist " & Form1.List1.ListCount & " Songs"

    End Function
Private Sub add_Click()
On Error Resume Next
Form5.show
If Form1.List1 = "" Then Exit Sub
If Form1.List2 = "" Then Exit Sub
Form5.List1.AddItem Form1.List1
Form5.List2.AddItem Form1.List2
'Form1.List3.AddItem Form1.List1
'Form1.List4.AddItem Form1.List2
End Sub

Private Sub autosave_Click()
Dim autoexit As String
If autosave.Checked = True Then
autosave.Checked = False
Else
autosave.Checked = True
End If
autoexit$ = autosave.Checked
Call PutValue("OnExit", "Autosave", autoexit$, "options.ini")
End Sub

Private Sub autostart_Click()
Dim autoload As String
If autostart.Checked = True Then
autostart.Checked = False
Else
autostart.Checked = True
End If
autoload$ = autostart.Checked
Call PutValue("OnLoad", "Autoload", autoload$, "options.ini")
End Sub

Private Sub back_Click()
Form1.Label1.Caption = StrReverse(Form1.Label1.Caption)
End Sub

Private Sub clear_Click()
Form1.List1.clear
Form1.List2.clear
End Sub

Private Sub clear5_Click()
Form5.List1.clear
Form5.List2.clear
End Sub

Private Sub createlist_Click()
Form6.show
End Sub



Private Sub dupes_Click()
Dim ab
Dim abc
MsgBox "If you have alot of files this could take a minute.", , " Wolfie MP3 Player"
ab = LBDupe(Form1.List1)
abc = LBDupe(Form1.List2)
MsgBox "Removed: " & ab & " dupes", , " Wolfie MP3 Player"
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Width) \ 2
    Top = (Screen.Height - Height) \ 2

End Sub

Private Sub makedefault_Click()
'associatemyapp "File Descrip.", Location of new default application %1", ".file extension"
AssociateMyApp "My Mp3 Files", "c:\Wolfie MP3.exe %1", ".mp3"
End Sub



Private Sub mini_Click()
'Call AddToTray(Form1.Icon, Form1.Caption, Form1)
End Sub

Private Sub PLay_Click()
On Error Resume Next
Dim cMP3 As New clsMP3time
Dim oStr As String
Dim oStr1 As String
Dim oStr2 As String
Dim tm As Date, t1 As Double
Dim result As Boolean

Form1.MediaPlayer1.FileName = Form1.List2
cMP3.FileName = Form1.List2
result = cMP3.refresh

oStr = oStr & "" & Format$(cMP3.Time, "hh:mm:ss")
oStr1 = oStr1 & "" & (cMP3.BitRate / 1000) & "kb " '& cMP3.Frames & " frames" & vbCrLf
oStr2 = oStr2 & "" & cMP3.Frequency & "hz " '& cMP3.ChannelModeText & vbCrLf
Form1.Text1.text = oStr1
Form1.Text2.text = oStr2
Form1.Text3.text = oStr
If Form1.MediaPlayer1.IsDurationValid = False Then MsgBox "Make sure the song is valid.", , "Little MP3 Player": Exit Sub
Form1.HScroll1.Max = Form1.MediaPlayer1.Duration
Form1.MediaPlayer1.play
Form1.Timer1.Enabled = True
End Sub

Private Sub play5_Click()
List2.ListIndex = List1.ListIndex
On Error Resume Next
Dim cMP3 As New clsMP3time
Dim oStr As String
Dim oStr1 As String
Dim oStr2 As String
Dim tm As Date, t1 As Double
Dim result As Boolean

Form1.MediaPlayer1.FileName = Form5.List2

cMP3.FileName = Form5.List2
result = cMP3.refresh

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

Private Sub rem_Click()
Dim i
Dim h
Form5.List2.ListIndex = Form5.List1.ListIndex
i = Form5.List1.ListIndex
h = Form5.List2.ListIndex
Form5.List1.RemoveItem i
Form5.List2.RemoveItem h
End Sub

Private Sub remove_Click()
Dim i
Dim h
Form1.List2.ListIndex = Form1.List1.ListIndex
i = Form1.List1.ListIndex
h = Form1.List2.ListIndex
Form1.List1.RemoveItem i
Form1.List2.RemoveItem h
End Sub

Private Sub rename_Click()
On Error Resume Next
Dim OldName As String
Dim NewName As String
'Dim a As String
OldName$ = Form1.List2
NewName$ = InputBox("Make changes to filename. DO NOT change the path.", " Little MP3 Player", Form1.List2)
If NewName$ = OldName$ Then MsgBox "No changes made.", "Wolfie MP3 Player": Exit Sub
Name OldName$ As NewName$
End Sub

Private Sub rename5_Click()
On Error Resume Next
Dim OldName As String
Dim NewName As String
'Dim a As String
OldName$ = Form5.List2
NewName$ = InputBox("Make changes to filename. DO NOT change the path.", " Wolfie MP3 Player", Form5.List2)
If NewName$ = OldName$ Then MsgBox "No changes made.", "Little MP3 Player": Exit Sub
Name OldName$ As NewName$
End Sub

Private Sub savelist_Click()
c = DoEvents()
Call SavePlayList(Form1.List2, "C:\MP3Playlist.m3u")

End Sub

Private Sub scroll_Click()
If Form4.scroll.Caption = "Backwards Scroll" Then
Form1.lbltext.Caption = StrReverse(Form1.lbltext.Caption)
Form4.scroll.Caption = "Normal Scroll"
Else
Form1.lbltext.Caption = StrReverse(Form1.lbltext.Caption)
Form4.scroll.Caption = "Backwards Scroll"
End If
End Sub

Private Sub search_Click()
Form7.show
End Sub

Private Sub show_Click()
Form5.show
End Sub

Private Sub tag_Click()
Form8.show
End Sub
