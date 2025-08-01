VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00400040&
   Caption         =   "Wolfie MP3 Player"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3840
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   3840
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check3 
      BackColor       =   &H00400040&
      Caption         =   "Random"
      ForeColor       =   &H00FFC0C0&
      Height          =   150
      Left            =   1440
      TabIndex        =   32
      Top             =   1800
      Width           =   735
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00400040&
      Caption         =   "Continuous"
      ForeColor       =   &H00FFC0C0&
      Height          =   150
      Left            =   2640
      TabIndex        =   31
      Top             =   1800
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Options"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   1560
      Width           =   975
   End
   Begin VB.PictureBox pbscrollbox 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00400040&
      FillColor       =   &H00FFC0C0&
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
      Height          =   255
      Left            =   1440
      ScaleHeight     =   195
      ScaleWidth      =   915
      TabIndex        =   28
      Top             =   1440
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1680
      Top             =   2400
   End
   Begin VB.ListBox List4 
      Height          =   360
      Left            =   1680
      TabIndex        =   27
      Top             =   5640
      Width           =   1215
   End
   Begin VB.ListBox List3 
      Height          =   360
      Left            =   1320
      TabIndex        =   26
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox txt1 
      BackColor       =   &H00400040&
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   1440
      TabIndex        =   25
      Text            =   "Search List"
      Top             =   1080
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00400040&
      Caption         =   "Mute"
      ForeColor       =   &H00FFC0C0&
      Height          =   150
      Left            =   2640
      TabIndex        =   23
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BackColor       =   &H00400040&
      ForeColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "Balance"
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BackColor       =   &H00400040&
      ForeColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "Volume"
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H00400040&
      ForeColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "Time Elapsed"
      Top             =   360
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3000
      Top             =   2400
   End
   Begin VB.ListBox List2 
      Height          =   210
      Left            =   480
      TabIndex        =   19
      Top             =   5760
      Width           =   3135
   End
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
      Height          =   1815
      Left            =   240
      TabIndex        =   18
      Top             =   2280
      Width           =   3375
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   135
      LargeChange     =   10
      Left            =   2640
      Max             =   5000
      Min             =   -5000
      TabIndex        =   17
      Top             =   1320
      Width           =   975
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   135
      LargeChange     =   5
      Left            =   2640
      Max             =   2500
      TabIndex        =   16
      Top             =   960
      Value           =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   14
      Text            =   "Text4"
      Top             =   5520
      Width           =   3135
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   135
      LargeChange     =   15
      Left            =   2640
      Max             =   0
      SmallChange     =   5
      TabIndex        =   13
      Top             =   600
      Width           =   975
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00400040&
      Caption         =   "MP3 Playlist"
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
      Height          =   2175
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   3615
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00400040&
      Caption         =   "Controls"
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
      Height          =   1935
      Left            =   2520
      TabIndex        =   11
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00400040&
      ForeColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Duration"
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00400040&
      ForeColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Freq."
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00400040&
      ForeColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Bitrate:"
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Add Directory"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Add Mp3"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Pause Song"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Stop Song"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Play Song"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00400040&
      Caption         =   "MP3 Info"
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
      Height          =   1935
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
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
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lbltext 
      Height          =   375
      Left            =   1320
      TabIndex        =   29
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   360
      TabIndex        =   24
      Top             =   5640
      Width           =   1215
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   135
      Left            =   1080
      TabIndex        =   0
      Top             =   5520
      Width           =   2295
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -230
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private TheX As Long
Private TheY As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Sub MP3Info()
On Error Resume Next
Dim cMP3 As New clsMP3time
Dim oStr As String
Dim oStr1 As String
Dim oStr2 As String
Dim tm As Date, t1 As Double
Dim result As Boolean

MediaPlayer1.FileName = FileName

cMP3.FileName = FileName
result = cMP3.refresh

oStr = oStr & "" & Format$(cMP3.Time, "hh:mm:ss") 'hh:
oStr1 = oStr1 & "" & (cMP3.BitRate / 1000) & "kb " '& cMP3.Frames & " frames" & vbCrLf
oStr2 = oStr2 & "" & cMP3.Frequency & "hz " '& cMP3.ChannelModeText & vbCrLf
Text1.text = "Bitrate: " & oStr1
Text2.text = "Freq. : " & oStr2
Text3.text = "Time: " & oStr
End Sub
Function percent(Complete As Integer, Total As Variant, TotalOutput As Integer) As Integer
    On Error Resume Next
    percent = Int(Complete / Total * TotalOutput)
End Function
Private Sub Check1_Click()
If MediaPlayer1.Mute = False Then
MediaPlayer1.Mute = True
Check1.Value = Checked
Else
MediaPlayer1.Mute = False
Check1.Value = Unchecked
End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim cMP3 As New clsMP3time
Dim oStr As String
Dim oStr1 As String
Dim oStr2 As String
Dim tm As Date, t1 As Double
Dim result As Boolean
MediaPlayer1.FileName = List2

cMP3.FileName = List2
result = cMP3.refresh

oStr = oStr & "" & Format$(cMP3.Time, "hh:mm:ss") 'hh:
oStr1 = oStr1 & "" & (cMP3.BitRate / 1000) & "kb " '& cMP3.Frames & " frames" & vbCrLf
oStr2 = oStr2 & "" & cMP3.Frequency & "hz " '& cMP3.ChannelModeText & vbCrLf
Text1.text = "Bitrate: " & oStr1
Text2.text = "Freq. : " & oStr2
Text3.text = "Time: " & oStr
If MediaPlayer1.IsDurationValid = False Then
MsgBox "Make sure the song is valid.", , "Wolfie MP3 Player"
Exit Sub
End If
HScroll1.Max = MediaPlayer1.Duration
MediaPlayer1.play
lbltext.Caption = List1
Timer1.Enabled = True
Timer2.Enabled = True
End Sub
Private Sub AssociateMyApp(ByVal sAppName As String, ByVal sEXE As String, ByVal sExt As String)
Dim lRegKey As Long
Call RegCreateKey(HKEY_CLASSES_ROOT, sExt, lRegKey)
Call RegSetValueEx(lRegKey, "", 0&, 1, ByVal sAppName, Len(sAppName))
Call RegCloseKey(lRegKey)
Call RegCreateKey(HKEY_CLASSES_ROOT, sAppName & "\Shell\Open\Command", lRegKey) ' adds info into the shell open command
Call RegSetValueEx(lRegKey, "", 0&, 1, ByVal sEXE, Len(sEXE))
Call RegCloseKey(lRegKey)
End Sub

Private Sub Command2_Click()
MediaPlayer1.Stop
Timer1.Enabled = False
Timer2.Enabled = False
End Sub

Private Sub Command3_Click()
On Error Resume Next
If Command3.Caption = "Pause Song" Then
MediaPlayer1.Pause
Timer1.Enabled = False
Timer2.Enabled = False
Command3.Caption = "Resume Song"
Exit Sub
End If
If Command3.Caption = "Resume Song" Then
MediaPlayer1.play
Timer1.Enabled = True
Timer2.Enabled = True
Command3.Caption = "Pause Song"
Exit Sub
End If

End Sub

Private Sub Command4_Click()
Form3.show
End Sub

Private Sub Command5_Click()
Form2.show
'File1.Path = "c:\"
End Sub

Private Sub Command6_Click()
'If File1.ListIndex < 0 Then Exit Sub

'Dim cMP3 As New clsMP3time
'Dim oStr As String
'Dim oStr1 As String
'Dim oStr2 As String
'Dim tm As Date, t1 As Double
'Dim Result As Boolean

'cMP3.FileName = File1.Path & "\" & File1.list(File1.ListIndex)
'Result = cMP3.refresh

'If Result Then
'oStr = oStr & "" & Format$(cMP3.Time, "hh:mm:ss")
'oStr1 = oStr1 & "" & (cMP3.BitRate / 1000) & "kb " '& cMP3.Frames & " frames" & vbCrLf
'oStr2 = oStr2 & "" & cMP3.Frequency & "hz " '& cMP3.ChannelModeText & vbCrLf
'Text1.Text = oStr1
'Text2.Text = oStr2
'Text3.Text = oStr
'Else
'Exit Sub
'End If
End Sub

Private Sub Command7_Click()
PopupMenu Form4.option
End Sub

Private Sub Form_Load()
Dim a
Dim listload As String
Dim listload2 As String
On Error Resume Next
a = DoEvents()
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2
'
'
If Command = "" Then
GoTo starting
Else
FileName = LCase(Command) ' gets the file location
List2.AddItem FileName
Call LoadListView(List1, List2)
MediaPlayer1.FileName = FileName
MediaPlayer1.play
Call MP3Info
lbltext.Caption = FileName
Timer1.Enabled = True
Timer2.Enabled = True
Exit Sub
End If
'
'
starting:
listload$ = GetValue("OnLoad", "Autoload", "options.ini")
listload2$ = GetValue("OnExit", "Autosave", "options.ini")
Form4.autostart.Checked = listload$
Form4.autosave.Checked = listload2$
If listload$ = "true" Then
If FileExists("mp3playlist.m3u") = False Then Exit Sub
Call LoadPlayList(List2, "MP3Playlist.m3u")
Call LoadListView(List1, List2)
End If
Frame3.Caption = "MP3 Playlist " & List1.ListCount & " Songs"
Call AddScroll(List1)
End Sub

Private Sub Form_Resize()
If Form1.WindowState = 1 Then Exit Sub
    ResizeAll Form1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim listsave As String
Dim c
listsave$ = GetValue("OnExit", "Autosave", "options.ini")
If listsave$ = "true" Then
c = DoEvents()
Call SavePlayList(List2, "MP3Playlist.m3u")
Else
Unload Me
End
End If
Unload Me
End
End Sub

Private Sub HScroll1_Change()
'MediaPlayer1.CurrentPosition = HScroll1
End Sub

Private Sub HScroll1_Scroll()
MediaPlayer1.CurrentPosition = HScroll1
End Sub

Private Sub HScroll2_Change()
Dim ult As Integer, inc As Integer
Dim a As Integer, b As Integer
Dim d, c
ult = HScroll2.Min
inc = HScroll2.Value
c = HScroll2.Value - 2500
MediaPlayer1.Volume = c
b = HScroll2.Min
a = HScroll2.Value
Text6.text = "Volume: " & inc \ 25 & " %"
'Timer2.Enabled = True
End Sub

Private Sub HScroll3_Change()
On Error GoTo hello
If HScroll3.Value > -500 And HScroll3.Value < 500 Then
Text7.text = "Center"
End If
If HScroll3.Value < -500 Then
Text7.text = "Left"
End If
If HScroll3.Value > 500 Then
Text7.text = "Right"
End If
MediaPlayer1.Balance = HScroll3.Value
Exit Sub
hello:
'MsgBox "Err"
Exit Sub
End Sub

Private Sub HScroll3_Scroll()
On Error Resume Next
If HScroll3.Value > -500 And HScroll3.Value < 500 Then
Text7.text = "Center"
End If
If HScroll3.Value < -500 Then
Text7.text = "Left"
End If
If HScroll3.Value > 500 Then
Text7.text = "Right"
End If
MediaPlayer1.Balance = HScroll3.Value
Exit Sub
End Sub

Private Sub List1_Click()
On Error Resume Next
List2.ListIndex = List1.ListIndex
Frame3.Caption = "MP3 Playlist " & List1.ListCount & " Songs"
End Sub

Private Sub List1_DblClick()
List2.ListIndex = List1.ListIndex
lbltext.Caption = List1
Command1_Click
End Sub

Private Sub List1_GotFocus()
On Error Resume Next

List2.ListIndex = List1.ListIndex
Frame3.Caption = "MP3 Playlist " & List1.ListCount & " Songs"
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
List2.ListIndex = List1.ListIndex
Command1_Click
End If
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If List1 = "" Then Exit Sub
If Button = 2 Then
PopupMenu Form4.file
End If
End Sub

Private Sub List2_Click()
Label1.Caption = List2.ListCount
End Sub

Private Sub MediaPlayer1_EndOfStream(ByVal result As Long)
 Timer1.Enabled = False
 Timer2.Enabled = False
 On Error Resume Next
    Dim x As Integer
    Dim NoFreeze
    If Check3.Value = 1 Then
    NoFreeze = DoEvents()
    Randomize
    x = Int((List1.ListCount * Rnd) + 1)
    List1.ListIndex = x
    Command1_Click
    End If
If Check3.Value = 1 Then Exit Sub
If Check2.Value = 1 Then
On Error Resume Next
Dim a
Dim b
Dim c
a = List1.ListIndex
c = Val(List1.ListIndex) + 1
b = List1.list(c)
List1 = b
Command1_Click
'Next a
End If
If Form5.Check3.Value = 1 Then
Form1.MediaPlayer1.FileName = Form5.List2
Form1.MediaPlayer1.play
End If
End Sub

Private Sub MediaPlayer1_PositionChange(ByVal oldPosition As Double, ByVal newPosition As Double)
HScroll1.Value = Int(MediaPlayer1.CurrentPosition)
End Sub




Private Sub pbscrollbox_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 Then
PopupMenu Form4.color
End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Dim tinseconden
HScroll1.Value = Int(MediaPlayer1.CurrentPosition)
tinseconden = MediaPlayer1.CurrentPosition
Dim Min As Integer
Dim sec As Integer
Min = tinseconden \ 60
sec = tinseconden - (Min * 60)
If sec = "-1" Then sec = "0"
Text5.text = Min & ":" & sec
If Timer1.Enabled = False Then Exit Sub
End Sub

Private Sub Timer2_Timer()
If Timer2.Enabled = False Then Exit Sub
pbscrollbox.Cls ' so we don't Get text trails
    ' Scroll from right to left


    If TheX <= 0 - lbltext.Width Then
        TheX = pbscrollbox.ScaleWidth
    Else
        TheX = TheX - 10 ' larger number means faster scrolling
    End If
    pbscrollbox.CurrentX = TheX
    pbscrollbox.CurrentY = TheY
    pbscrollbox.Print lbltext.Caption
pbscrollbox.ForeColor = QBColor(Rnd * 15)
If Timer2.Enabled = False Then Exit Sub
End Sub

Private Sub txt1_Change()
listsearch
On Error Resume Next

List2.ListIndex = List1.ListIndex
Frame3.Caption = "MP3 Playlist " & List1.ListCount & " Songs"

End Sub

