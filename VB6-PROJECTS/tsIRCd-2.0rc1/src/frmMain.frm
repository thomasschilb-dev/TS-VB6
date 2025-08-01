VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "tsIRCd"
   ClientHeight    =   1770
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   5910
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   5910
   StartUpPosition =   2  'CenterScreen
   Begin tsIRCd.fbTrayIcon fbTrayIcon1 
      Height          =   1155
      Left            =   3000
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   2037
   End
   Begin VB.Timer tmrLinkPing 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   65535
      Left            =   3180
      Top             =   1260
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Shutdown"
      Height          =   315
      Left            =   4560
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdRestart 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Restart"
      Height          =   315
      Left            =   4560
      MaskColor       =   &H000000FF&
      TabIndex        =   3
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Link 
      Index           =   0
      Left            =   360
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "dennis"
      RemotePort      =   6669
      LocalPort       =   6668
   End
   Begin VB.Timer tmrTimeOut 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   60000
      Left            =   3660
      Top             =   1260
   End
   Begin VB.Timer tmrKlined 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   10000
      Left            =   2220
      Top             =   2640
   End
   Begin VB.Timer tmrKill 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   200
      Left            =   2700
      Top             =   1260
   End
   Begin VB.Timer tmrFloodProt 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   250
      Left            =   1740
      Top             =   2640
   End
   Begin VB.Timer tmrSend 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   200
      Left            =   1260
      Top             =   2640
   End
   Begin MSWinsockLib.Winsock wsock 
      Index           =   0
      Left            =   780
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   6667
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "2.0rc1"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   5280
      TabIndex        =   7
      Top             =   1440
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   5760
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblServer 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   1
      Left            =   240
      LinkTimeout     =   60
      TabIndex        =   6
      Top             =   1320
      Width           =   4275
   End
   Begin VB.Label lblServer 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Index           =   0
      Left            =   960
      LinkTimeout     =   60
      TabIndex        =   2
      Top             =   1320
      Width           =   1035
   End
   Begin VB.Label lblServerSocket 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   4155
   End
   Begin VB.Label lblClientSocket 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4155
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Options"
      Begin VB.Menu mnuMainStartServer 
         Caption         =   "Restart"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuMainCloseServer 
         Caption         =   "Shutdown"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuMainExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
   Begin VB.Menu mnuTray 
      Caption         =   "mnuTray"
      Visible         =   0   'False
      Begin VB.Menu mnuTrayStartServer 
         Caption         =   "Restart"
      End
      Begin VB.Menu mnuTrayCloseServer 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu mnuTrayShow 
         Caption         =   "Show/Hide"
      End
      Begin VB.Menu mnuTrayExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1

Private Sub cmdClose_Click()
Dim X As Long
For X = LBound(Users) To UBound(Users)
    If Not Users(X) Is Nothing Then SendNotice "", "*** Global -- " & "Recieved DIE command from CONSOLE", "GLOBAL", , CInt(X)
Next X
Wait 2000
Unload Me
End Sub

Private Sub cmdRestart_Click()
Dim X As Long
For X = LBound(Users) To UBound(Users)
    If Not Users(X) Is Nothing Then SendNotice "", "*** Global -- " & "Recieved RESTART command from CONSOLE", "GLOBAL", , CInt(X)
Next X
Wait 2000
Restart
End Sub

Private Sub fbTrayIcon1_MouseClick(ByVal FBButton As EnumFBButtonConstants)
Select Case FBButton
    Case &H203
        frmMain.Visible = Not frmMain.Visible
    Case &H205
        frmMain.PopupMenu mnuTray, , , , mnuTrayShow
End Select
End Sub

Private Sub Form_Load()
Dim i As Long, FS As New FileSystemObject
ReDim Users(4)
UserCount = UserCount + 4
Rehash
Link(0).LocalPort = LinkPort
Link(0).Listen
Me.Caption = "tsIRCd"
For i = LBound(Users) To UBound(Users): Set Users(i) = Nothing: Next i
For i = LBound(Channels) To UBound(Channels): Set Channels(i) = Nothing: Next i
fbTrayIcon1.AddTrayIcon App.Path & "\tsircd2.ico", "tsIRCd"
Started = Now
If LogLevel > 3 Then WriteHeader
MaxGlobalUsers = 0
CurGlobalUsers = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
If LogLevel <> 0 Then WriteFooter
fbTrayIcon1.RemoveTrayIcon
End
End Sub

Private Sub Link_Close(Index As Integer)
SendSvrMsg ServerName & " -- link closed -- " & Link(Index).Tag, False
SendLinks "DeadLink" & vbLf & ServerName & vbLf & Link(Index).Tag
If LogLevel = 1 Or LogLevel = 3 Then
    If LogFormat = 0 Then
        LogText "[LINK]<" & Now & " (LINK CLOSED " & ServerName & " -- " & Link(Index).Tag & ")> "
    Else
        LogHTML ServerName, "LINK CLOSED " & ServerName & " -- " & Link(Index).Tag
    End If
End If
Dim i As Long
On Error Resume Next
For i = 1 To UBound(Users)
    DoEvents
    If Not Users(i) Is Nothing Then
        If Users(i).IsOnLink(Link(Index).Tag) Then
            SendQuit i, ServerName & Link(Index).Tag
            Set Users(i) = Nothing
        End If
    End If
Next i
CurLinkCount = CurLinkCount - 1
Unload Link(Index)
Unload tmrLinkPing(Index)
End Sub

Private Sub Link_Connect(Index As Integer)
On Error Resume Next
Wait 100
Dim i As Long, X As Long
For i = 1 To UBound(Users)
    If Not Users(i) Is Nothing Then
        SendLinks "NewUser" & vbLf & Users(i).Nick & vbLf & Users(i).Name & vbLf & Users(i).DNS & vbLf & Users(i).Ident & vbLf & Users(i).Server & vbLf & Users(i).ServerDescritption & vbLf & Users(i).SignOn & vbLf & Users(i).GID & vbLf & Users(i).GetModes & vbLf & ServerName & "-", , Index
        Wait 10
        For X = 1 To Users(i).Onchannels.Count
            SendLinks "JoinChan" & vbLf & Users(i).Nick & vbLf & Users(i).Onchannels(X), , Index '1 = Command, 2 = Nick, 3 = Channel
        Next X
    End If
Next i
For i = 1 To UBound(Channels)
    If Not Channels(i) Is Nothing Then
        SendLinks "ChanMode" & vbLf & "ChanServ" & vbLf & "+" & vbLf & Channels(i).GetModesForFile & vbLf & Channels(i).Name, , Index
        If Channels(i).Key <> "" Then SendLinks "Key" & vbLf & "ChanServ" & vbLf & Channels(i).Name & vbLf & Channels(i).Key, , Index
        If Channels(i).Limit <> 0 Then SendLinks "Limit" & vbLf & "ChanServ" & vbLf & Channels(i).Name & vbLf & Channels(i).Limit, , Index
        Dim Y As Long
        SendLinks "SetTopic" & vbLf & "ChanServ" & vbLf & Channels(i).Name & vbLf & "" & vbLf & Channels(i).Topic, , Index
        For Y = 1 To Channels(i).Bans.Count
            SendLinks "BanUser" & vbLf & "ChanServ" & vbLf & Channels(i).Name & vbLf & "" & vbLf & Channels(i).Bans(Y), , Index
        Next Y
        For Y = 1 To Channels(i).Invites.Count
            SendLinks "InviteUser" & vbLf & "ChanServ" & vbLf & Channels(i).Name & vbLf & "" & vbLf & Channels(i).Invites(Y), , Index
        Next Y
        For Y = 1 To Channels(i).Exceptions.Count
            SendLinks "ExceptUser" & vbLf & "ChanServ" & vbLf & Channels(i).Name & vbLf & "" & vbLf & Channels(i).Exceptions(Y), , Index
        Next Y
    End If
Next i
SendLinks "Info" & vbLf & ServerName
On Local Error GoTo 0
Load tmrLinkPing(Index)
tmrLinkPing(Index).Enabled = True
tmrLinkPing(Index).Tag = 1
End Sub

Private Sub Link_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim LinkCount As Long
LinkCount = Link.Count + 1
CurLinkCount = CurLinkCount + 1
MaxLinkCount = MaxLinkCount + 1
Index = LinkCount
Load Link(LinkCount)
Link(LinkCount).Close
Link(LinkCount).LocalPort = 30000 + LinkCount
Link(LinkCount).Accept requestID
Wait 100
Dim i As Long, X As Long
For i = 1 To UBound(Users)
    If Not Users(i) Is Nothing Then
        SendLinks "NewUser" & vbLf & Users(i).Nick & vbLf & Users(i).Name & vbLf & Users(i).DNS & vbLf & Users(i).Ident & vbLf & Users(i).Server & vbLf & Users(i).ServerDescritption & vbLf & Users(i).SignOn & vbLf & Users(i).GID & vbLf & Users(i).GetModes & vbLf & ServerName & " ", , Index
        For X = 1 To Users(i).Onchannels.Count
            SendLinks "JoinChan" & vbLf & Users(i).Nick & vbLf & Users(i).Onchannels(X), , Index '1 = Command, 2 = Nick, 3 = Channel
        Next X
    End If
Next i
For i = 1 To UBound(Channels)
    If Not Channels(i) Is Nothing Then
        SendLinks "ChanMode" & vbLf & "ChanServ" & vbLf & "+" & vbLf & Channels(i).GetModesForFile & vbLf & Channels(i).Name, , Index
        If Channels(i).Key <> "" Then SendLinks "Key" & vbLf & "ChanServ" & vbLf & Channels(i).Name & vbLf & Channels(i).Key, , Index
        If Channels(i).Limit <> 0 Then SendLinks "Limit" & vbLf & "ChanServ" & vbLf & Channels(i).Name & vbLf & Channels(i).Limit, , Index
        Dim Y As Long
        SendLinks "SetTopic" & vbLf & "ChanServ" & vbLf & Channels(i).Name & vbLf & "" & vbLf & Channels(i).Topic, , Index
        For Y = 1 To Channels(i).Bans.Count
            SendLinks "BanUser" & vbLf & "ChanServ" & vbLf & Channels(i).Name & vbLf & "" & vbLf & Channels(i).Bans(Y), , Index
        Next Y
        For Y = 1 To Channels(i).Invites.Count
            SendLinks "InviteUser" & vbLf & "ChanServ" & vbLf & Channels(i).Name & vbLf & "" & vbLf & Channels(i).Invites(Y), , Index
        Next Y
        For Y = 1 To Channels(i).Exceptions.Count
            SendLinks "ExceptUser" & vbLf & "ChanServ" & vbLf & Channels(i).Name & vbLf & "" & vbLf & Channels(i).Exceptions(Y), , Index
        Next Y
    End If
Next i
SendLinks "Info" & vbLf & ServerName
Load tmrLinkPing(Index)
tmrLinkPing(Index).Enabled = True
tmrLinkPing(Index).Tag = 1
End Sub

Private Sub Link_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
Dim strData As String, strcmd() As String, cmdArray() As String, i As Long, User As clsUser, DontSendLink As Boolean, X As Long, NewUser As clsUser, strRoute() As String
Link(Index).GetData strData, 8
If LogLevel = 1 Then
    If LogFormat = 0 Then
        LogText "[LINK]<" & Now & " (INCOMING " & Link(Index).Tag & ")> " & strData
    Else
        LogHTML Link(Index).Tag & "(Link) INCOMING", strData
    End If
End If
cmdArray = Split(strData, vbCrLf)
For X = LBound(cmdArray) To UBound(cmdArray)
    If cmdArray(X) = "" Then GoTo NextCmd
    strcmd = Split(cmdArray(X), vbLf)
    Select Case strcmd(0)
        Case "Info"
            Link(Index).Tag = strcmd(1)
            SendSvrMsg ServerName & " -- linked -- " & Link(Index).Tag, False
            If LogLevel = 1 Or LogLevel = 3 Then
                If LogFormat = 0 Then
                    LogText "[LINK]<" & Now & " (LINKED " & ServerName & " -- " & strcmd(1) & ")> " & strData
                Else
                    LogHTML ServerName, "LINKED " & ServerName & " -- " & strcmd(1)
                End If
            End If
        Case "NewUser"
            If strcmd(8) = "" Then GoTo NextCmd
            '1 = Command, 2 = Nick, 3 = Name, 4 = DNS, 5 = Ident, 6 = Server, 7 = ServerDescription, 8 = SignOn
            Set NewUser = NickToObject(strcmd(1))
            If NewUser Is Nothing Then
                Set User = GetFreeSlot
                UserCount = UserCount - 1
                User.DNS = strcmd(3)
                User.Nick = strcmd(1)
                User.Name = strcmd(2)
                User.Ident = strcmd(4)
                User.Server = strcmd(5)
                User.ServerDescritption = strcmd(6)
                User.SignOn = strcmd(7)
                User.GID = strcmd(8)
                If InStr(1, strcmd(10), " ") = 0 Then strcmd(10) = strcmd(10) & " "
                User.Route = strcmd(10)
                User.Hops = CountSpaces(strcmd(10)) - 1
                User.AddModes strcmd(9)
            Else
                If strcmd(8) = NewUser.GID Then GoTo NextCmd
                If NewUser.SignOn < strcmd(7) Then
                    SendLinks "KillUser" & vbLf & strcmd(1) & vbLf & "Nick Collision, other nick signed on earlier"
                Else
                    SendWsock Index, ":" & Users(NewUser.Index).Nick & "!" & Users(NewUser.Index).Ident & "@" & Users(NewUser.Index).DNS & " KILL :" & "Nick Collision, other nick signed on earlier"
                    SendWsock Index, "ERROR :Closing Link: " & "Nick Collision, other nick signed on earlier" & vbCrLf
                    Set User = GetFreeSlot
                    UserCount = UserCount - 1
                    User.DNS = strcmd(3)
                    User.Nick = strcmd(1)
                    User.Name = strcmd(2)
                    User.Ident = strcmd(4)
                    User.Server = strcmd(5)
                    User.ServerDescritption = strcmd(6)
                    User.SignOn = strcmd(7)
                    User.GID = strcmd(8)
                    If InStr(1, strcmd(10), " ") = 0 Then strcmd(10) = strcmd(10) & " "
                    User.Route = strcmd(10)
                    User.Hops = CountSpaces(strcmd(10)) - 1
                    User.AddModes strcmd(9)
                End If
            End If
            SendLinks cmdArray(X) & " " & ServerName & " ", CLng(Index)
            DontSendLink = False
        Case "QuitUser"
            '1 = Command, 2 = Nick, 3 = QuitMsg
            Set User = NickToObject(strcmd(1))
            SendQuit User.Index, strcmd(2), , False
            Set Users(User.Index) = Nothing
        Case "KillUser"
            '1 = Command, 2 = Nick, 3 = Reason
            Set User = NickToObject(strcmd(1))
            SendQuit User.Index, strcmd(2), True, False
            Set Users(User.Index) = Nothing
        Case "JoinChan"
            '1 = Command, 2 = Nick, 3 = Channel
            Set User = NickToObject(strcmd(1))
            If Not ChanExists(strcmd(2)) Then
                Dim NewChannel As clsChannel
                Set NewChannel = GetFreeChan
                NewChannel.Name = strcmd(2)
                NewChannel.Modes.Add "t", "t"
                NewChannel.Modes.Add "n", "n"
                NewChannel.Topic = DefTopic
                NotifyJoin User.Index, strcmd(2), False
                NewChannel.NormUsers.Add User.Nick, User.Nick
                NewChannel.All.Add User.Nick, User.Nick
                Users(Index).Onchannels.Add strcmd(2), strcmd(2)
            Else
                NotifyJoin User.Index, strcmd(2), False
                ChanToObject(strcmd(2)).All.Add User.Nick, User.Nick
                ChanToObject(strcmd(2)).NormUsers.Add User.Nick, User.Nick
            End If
            User.Onchannels.Add strcmd(2), strcmd(2)
        Case "PartUser"
            '1 = Command, 2 = Nick, 3 = Channel, 4 = Reason
            Set User = NickToObject(strcmd(1))
            SendPart User.Index, strcmd(2), strcmd(3), False
        Case "ModeUser"
            '1 = Command, 2 = Nick, 3 = +/-, 4 = Modes
            Set User = NickToObject(strcmd(1))
            Select Case strcmd(2)
                Case "+"
                    AddUserMode User.Index, strcmd(3), , False
                Case "-"
                    RemoveUsermode User.Index, strcmd(3), , False
            End Select
        Case "ChanMode"
            '1 = Command, 2 = Nick, 3 = +/-, 4 = Modes, 5 = Channel
            Set User = NickToObject(strcmd(1))
            Select Case strcmd(2)
                Case "+"
                    AddChanModes strcmd(3), strcmd(4), User, False
                Case "-"
                    RemoveChanModes strcmd(3), strcmd(4), User, False
            End Select
        Case "KickUser"
            '1 = Command, 2 = Nick, 3 = Channel, 4 = Reason, 5 = Target
            Set User = NickToObject(strcmd(1))
            If strcmd(3) = "" Then strcmd(3) = strcmd(1)
            KickUser User.Nick, strcmd(2), strcmd(4), strcmd(3), True, False
        Case "KLine"
            '1 = Command, 2 = Mask
            Klines.Add strcmd(1), strcmd(1)
        Case "ServerMsg"
            '1 = Command, 2 = Msg
            SendSvrMsg strcmd(1), , Link(Index).Tag
        Case "Global"
            '1 = Command, 2 = Msg
            For i = LBound(Users) To UBound(Users)
                If Not Users(i) Is Nothing Then SendNotice "", "*** Global -- " & strcmd(1), ServerName, , CInt(i), False
            Next i
        Case "PrivMsgChan"
            '1 = Command, 2 = Nick, 3 = Channel, 4 = Msg
            SendMsg strcmd(2), strcmd(3), strcmd(1), True, False
        Case "PrivMsgUser"
            '1 = Command, 2 = Nick, 3 = Target, 4 = Msg
            SendMsg strcmd(2), strcmd(3), strcmd(1), False, False
        Case "NoticeUser"
            '1 = Command, 2 = Nick, 3 = Target, 4 = Msg
            SendNotice strcmd(2), strcmd(3), strcmd(1), , , False
        Case "NoticeChan"
            '1 = Command, 2 = Nick, 3 = Channel, 4 = Msg
            SendNotice strcmd(2), strcmd(3), strcmd(1), True, , False
        Case "Nick"
            '1 = Command, 2 = Nick, 3 = NewNick
            Set User = NickToObject(strcmd(1))
            ChangeNick User.Index, strcmd(2), False
        Case "OpUser"
            '1 = Command, 2 = Nick, 3 = Channel, 4 = Target
            OpUser ChanToObject(strcmd(2)), strcmd(4), strcmd(1), True, False
        Case "DeOpUser"
            '1 = Command, 2 = Nick, 4 = Channel, 4 = Target
            DeOpUser ChanToObject(strcmd(2)), strcmd(4), strcmd(1), False
        Case "VoiceUser"
            '1 = Command, 2 = Nick, 4 = Channel, 4 = Target
            VoiceUser ChanToObject(strcmd(2)), strcmd(4), strcmd(1), True, False
        Case "DeVoiceUser"
            '1 = Command, 2 = Nick, 4 = Channel, 4 = Target
            DeVoiceUser ChanToObject(strcmd(2)), strcmd(4), strcmd(1), False
        Case "BanUser"
            '1 = Command, 2 = Nick, 4 = Channel, 4 = Mask
            BanUser ChanToObject(strcmd(2)), strcmd(4), strcmd(1), False
        Case "UnBanUser"
            '1 = Command, 2 = Nick, 4 = Channel, 4 = Mask
            UnBanUser ChanToObject(strcmd(2)), strcmd(4), strcmd(1), False
        Case "ExceptUser"
            '1 = Command, 2 = Nick, 4 = Channel, 4 = Mask
            ExceptionUser ChanToObject(strcmd(2)), strcmd(4), strcmd(1), False
        Case "UnExceptUser"
            '1 = Command, 2 = Nick, 4 = Channel, 4 = Mask
            UnExceptionUser ChanToObject(strcmd(2)), strcmd(4), strcmd(1), False
        Case "InviteUser"
            '1 = Command, 2 = Nick, 4 = Channel, 4 = Mask
            InviteUser ChanToObject(strcmd(2)), strcmd(4), strcmd(1), False
        Case "UnInviteUser"
            '1 = Command, 2 = Nick, 4 = Channel, 4 = Mask
            UnInviteUser ChanToObject(strcmd(2)), strcmd(4), strcmd(1), False
        Case "Limit"
            '1 = Command, 2 = Nick, 4 = Channel, 4 = Mask
            AddChanModes "l " & strcmd(3), strcmd(2), NickToObject(strcmd(1)), False
        Case "Key"
            '1 = Command, 2 = Nick, 4 = Channel, 4 = Mask
            AddChanModes "k " & strcmd(3), strcmd(2), NickToObject(strcmd(1)), False
        Case "AddInvite"
            '1 = Command, 2 = Nick, 4 = Channel, 4 = Target
            ChanToObject(strcmd(2)).Invited.Add strcmd(4), strcmd(4)
            If NickToObject(strcmd(4)).LocalUser Then SendWsock NickToObject(strcmd(4)).Index, ":" & strcmd(1) & " INVITE " & strcmd(4) & " " & strcmd(2)
        Case "SetTopic"
            '1 = Command, 2 = Nick, 4 = Channel, 4 = Target
            SetTopic strcmd(2), strcmd(4), strcmd(1), False
        Case "DeadLink"
            '1 = Command, 2 = Server1, 3 = Server2
            SendSvrMsg strcmd(1) & " -- Link Closed -- " & strcmd(2)
            For i = 5 To UBound(Users)
                If Users(i).IsOnLink(strcmd(2)) Then
                    SendQuit i, strcmd(1) & " -- " & strcmd(2)
                    Set Users(i) = Nothing
                End If
            Next i
            SendLinks cmdArray(X), CLng(Index)
            DontSendLink = False
        Case "PING"
            SendLinks "PONG!" & vbLf, , Index
            DontSendLink = True
        Case "PONG"
            tmrLinkPing(Index).Tag = 1
            DontSendLink = True
        Case Else
            DontSendLink = True
    End Select
    If Not DontSendLink Then SendLinks cmdArray(X), CLng(Index)
    DontSendLink = False
NextCmd:
Next X
End Sub

Private Sub Link_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
SendSvrMsg ServerName & " -- Link Closed -- " & Link(Index).Tag, True
If LogLevel = 1 Or LogLevel = 3 Then
    If LogFormat = 0 Then
        LogText "[LINK]<" & Now & " (LINK CLOSED " & ServerName & " -- " & Link(Index).Tag & ")> "
    Else
        LogHTML ServerName, "LINK CLOSED " & ServerName & " -- " & Link(Index).Tag
    End If
End If
Link_Close (Index)
End Sub

Private Sub mnuAbout_Click()
MsgBox "tsIRCd-2.0rc1", vbOKOnly, "About"
End Sub

Private Sub mnuMainCloseServer_Click()
wsock(0).Close
Unload Me
End Sub

Private Sub mnuMainExit_Click()
'wsock(0).Close
Unload Me
End Sub

Private Sub mnuMainStartServer_Click()
If Not wsock(0).State = 2 Then
    wsock(0).Close
    wsock(0).Listen
End If
End Sub

Private Sub mnuTrayCloseServer_Click()
mnuMainCloseServer_Click
End Sub

Private Sub mnuTrayExit_Click()
Unload Me
End Sub

Private Sub mnuTrayShow_Click()
frmMain.Visible = Not frmMain.Visible
End Sub

Private Sub mnuTrayStartServer_Click()
mnuMainStartServer_Click
End Sub

Private Sub tmrFloodProt_Timer(Index As Integer)
On Error Resume Next
If Users(Index).HasRegistered = False Then
    Users(Index).HasRegistered = True
    Users(Index).MsgsSent = 0
    Exit Sub
End If
If Users(Index).IRCOp Then Exit Sub
If Users(Index).MsgsSent > 3000 Then
    SendQuit CLng(Index), "killed by sysadmin (excess flooding)", True
    SendWsock Users(Index).Index, ":Server!IRCd@" & ServerName & " KILL " & Users(Index).Nick & " :Excess flooding", True
    SendWsock Users(Index).Index, "ERROR :Closing Link: " & Users(Index).Nick & "[" & frmMain.wsock(Index).RemoteHostIP & ".] " & ServerName & " (excess flooding)", True
    If LogLevel = 1 Or LogLevel = 3 Then
        If LogFormat = 0 Then
            LogText "[LINK]<" & Now & " (FLOOD PROTECTION " & Users(Index).Nick & ")> "
        Else
            LogHTML ServerName, "FLOOD PROTECTION " & Users(Index).Nick
        End If
    End If
    Users(Index).Killed = True
    SendSvrMsg "Recieved Kill message for " & Users(Index).Nick & "!" & Users(Index).Ident & "@" & Users(Index).DNS & " Path: " & Users(Index).Nick & " (excess flooding)", True
End If
Users(Index).FloodProt = Users(Index).FloodProt + 1
If Users(Index).FloodProt = 5 Then
    If GetPercent(3000, Users(Index).MsgsSent) >= 80 Then SendSvrMsg "Flooding Alert for user " & Users(Index).Nick & "! (" & Users(Index).MsgsSent & "/3000 (" & GetPercent(3000, Users(Index).MsgsSent) & "%)", True
    Users(Index).MsgsSent = 0
    Users(Index).FloodProt = 1
End If
End Sub

Private Sub tmrKill_Timer(Index As Integer)
If Index = 0 Then
    wsock_Close tmrKill(Index).Tag
    wsock(0).Listen
    tmrKill(0).Enabled = False
    tmrKill(0).Interval = 200
Else
    wsock_Close tmrKill(Index).Tag
    Unload tmrKill(Index)
End If
End Sub

Private Sub tmrKlined_Timer(Index As Integer)
On Error Resume Next
Klines.Remove tmrKlined(Index).Tag
Unload tmrKlined(Index)
End Sub

Private Sub tmrLinkPing_Timer(Index As Integer)
If Not CLng(tmrLinkPing(Index).Tag) = 1 Then
    SendQuit CLng(Index), "Ping Timeout"
    SendSvrMsg "No response from " & Link(Index).Tag & ", Closing Link", True, ServerName
    Link_Close (Index)
    Exit Sub
End If
tmrLinkPing(Index).Tag = 0
Link(Index).SendData "PING" & vbLf
End Sub

Private Sub tmrSend_Timer(Index As Integer)
On Error Resume Next
If Not wsock(Index).Tag = "" Then
    wsock(Index).Tag = Replace(wsock(Index).Tag, vbCrLf, "")
    Dim i As Long
    For i = 1 To Len(wsock(Index).Tag) Step MaxChunkSize
        wsock(Index).SendData Mid(wsock(Index).Tag, i, MaxChunkSize)
    Next i
    wsock(Index).SendData vbCrLf
    wsock(Index).Tag = ""
End If
End Sub

Private Sub tmrTimeOut_Timer(Index As Integer)
If Not Users(Index).Ponged Then
    SendQuit CLng(Index), "Ping Timeout"
    wsock_Close (Index)
    Exit Sub
End If
Users(Index).Ponged = False
SendPing CLng(Index)
End Sub

Public Sub wsock_Close(Index As Integer)
On Error Resume Next
If Not Users(Index).SentQuit Then SendQuit CLng(Index), "Client exited"
Dim i As Long, CurChan As clsChannel
For i = 1 To Users(Index).Onchannels.Count
        Set CurChan = ChanToObject(Users(Index).Onchannels(i))
        CurChan.All.Remove Users(Index).Nick
        If CurChan.IsNorm(Users(Index).Nick) Then
            CurChan.NormUsers.Remove Users(Index).Nick
        ElseIf CurChan.IsVoice(Users(Index).Nick) Then
            CurChan.Voices.Remove Users(Index).Nick
        ElseIf CurChan.IsOp(Users(Index).Nick) Then
            CurChan.Ops.Remove Users(Index).Nick
        End If
Next i
Set Users(Index) = Nothing
For i = 1 To CloneControl.Count
    If CloneControl(i) = wsock(Index).RemoteHostIP Then: CloneControl.Remove (i): Exit For
Next i
wsock(Index).Close
Unload wsock(Index)
Unload tmrTimeOut(Index)
Unload tmrFloodProt(Index)
Unload tmrSend(Index)
End Sub

Private Sub wsock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error Resume Next
Dim FS As clsUser, WelcomeStr As String
If ClientsFromIP(wsock(0).RemoteHostIP) >= SessionLimit Then
    SendSvrMsg "Session limit exceeded: " & wsock(0).RemoteHostIP & " [" & ServerName & "]", True
    wsock(0).Close
    wsock(0).Accept requestID
    wsock(0).SendData ":Server!Server@" & ServerName & " KILL You :You have exceeded your session limit" & vbCrLf
    wsock(0).SendData "ERROR :Closing Link: Session limit exceeded" & vbCrLf
    Wait 100
    wsock(0).Close
    wsock(0).Listen
    Exit Sub
End If
Dim Killine As String
Killine = IsKlined(AddressToName(wsock(0).RemoteHostIP))
If Killine <> "" Then
    SendSvrMsg "K-Line active for: " & wsock(0).RemoteHostIP & " (" & Killine & ") [" & ServerName & "]", True
    wsock(0).Close
    wsock(0).Accept requestID
    SendWsock 0, ":" & ServerName & " NOTICE AUTH :***  " & ServerName & " -- Your Site (IP, Country, ISP..etc...) has been banned from this Server", , True
    SendWsock 0, ":" & ServerName & " NOTICE AUTH :***  " & ServerName & " -- This is not necessarily your fault, if you think you have been banned without any reason please send an email to the admin: " & AdminEmail, , True
    SendWsock 0, ":Server!Server@" & ServerName & " KILL You :Your Site (IP, Country, ISP..etc...) has been banned from this Server", , True
    SendWsock 0, "ERROR :Closing Link: Your Site (IP, Country, ISP..etc...) has been banned from this Server", , True
    Wait 150
    wsock(0).Close
    wsock(0).Listen
    Exit Sub
End If
Set FS = GetFreeSlot
If FS Is Nothing Then
    SendSvrMsg "Server is Full: " & wsock(0).RemoteHostIP & " [" & ServerName & "]", True
    wsock(0).Close
    wsock(0).Accept requestID
    wsock(0).SendData ":Server!Server@" & ServerName & " KILL You :Server is full, try again later" & vbCrLf
    wsock(0).SendData "ERROR :Closing Link: Server is Full" & vbCrLf
    Wait 100
    wsock(0).Close
    wsock(0).Listen
    Exit Sub
End If
If ((UserCount - 4) > MaxUser) Then MaxUser = MaxUser + 1
Load wsock(FS.Index)
wsock(FS.Index).Accept requestID
wsock(0).Close: wsock(0).Listen
Load tmrTimeOut(FS.Index)
Load tmrFloodProt(FS.Index)
Load tmrSend(FS.Index)
tmrTimeOut(FS.Index).Enabled = True
tmrSend(FS.Index).Enabled = True
On Local Error Resume Next
Dim strDNS As String
strDNS = modDNS.AddressToName(wsock(FS.Index).RemoteHostIP)
FS.DNS = IIf(strDNS = "", wsock(FS.Index).RemoteHostIP, strDNS)
FS.Server = ServerName
FS.NewUser = True
FS.SignOn = UnixTime
FS.Idle = UnixTime
FS.ServerDescritption = ServerDesc
FS.LocalUser = True
FS.GID = CreateGUID
'Welcome User
WelcomeStr = ":" & ServerName & " NOTICE AUTH :*** Welcome to " & ServerName & "!" & vbCrLf & _
                           ":" & ServerName & " NOTICE AUTH :*** Looking up your Hostname...." & vbCrLf
'SendWsock FS.Index, WelcomeStr
End Sub

Private Sub wsock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error GoTo parseerr
If Users(Index).Killed Then Exit Sub
Dim strMsg As String, strcmd() As String, LB As Long, UB As Long, i As Long
wsock(Index).GetData strMsg
strMsg = Replace(strMsg, vbCrLf, vbLf)
Debug.Print strMsg
If Index <= 0 Then Exit Sub
Users(Index).MsgsSent = Users(Index).MsgsSent + (bytesTotal * 1.5)
If LogLevel = 1 Or LogLevel = 2 Then
    If LogFormat = 0 Then
        LogText "[Client]<" & Now & " (from " & Users(Index).Nick & ")> " & strMsg
    Else
        LogHTML Users(Index).Nick, strMsg
    End If
End If
ServerTraffic = ServerTraffic + bytesTotal
strcmd = Split(strMsg, vbLf)
LB = LBound(strcmd)
UB = UBound(strcmd)
For i = LB To UB
    If Users(Index) Is Nothing Then Exit Sub
    If Users(Index).Killed Or strcmd(i) = "" Then Exit Sub
'*****************************
'|      Client Commands     ||
'*****************************
'Nick
    If strcmd(i) Like "NICK*" Then
        Dim NewNick As String
        If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 150
        If InStr(1, strcmd(i), ":") <> 0 Then
            NewNick = Replace(strcmd(i), "NICK :", "")
        Else
            NewNick = Replace(strcmd(i), "NICK ", "")
        End If
        If InStr(1, NewNick, " ") <> 0 Then NewNick = Mid(NewNick, 1, InStr(1, NewNick, " ") - 1)
        If NewNick = Users(Index).Nick Then GoTo NextCmd
        If Len(NewNick) > Nicklen Then NewNick = Mid(NewNick, 1, Nicklen)
        If Not IsValidString(NewNick) Then SendWsock Index, ":" & ServerName & " 432 * " & NewNick & " :Erroneus nickname, Nickname has been cut"
        If Len(NewNick) > Nicklen Then NewNick = Left(NewNick, Nicklen)
        If Not (ChangeNick(CLng(Index), NewNick, (Not Users(Index).NewUser))) Then
            SendWsock Index, ":" & ServerName & " 433 * " & NewNick & " :Nickname is already in use"
            GoTo NextCmd
        Else
            SendWsock Index, "PING " & GetRand
            Users(Index).Identified = False
            Users(Index).NR = False
            Users(Index).ClearOwnerShip
        End If
'UserHost
    ElseIf strcmd(i) Like "USERHOST*" Then
        Dim User As clsUser
        Set User = NickToObject(Replace(strcmd(i), "USERHOST ", ""))
        If User Is Nothing Then
            SendWsock Index, ":" & ServerName & " 302 " & Users(Index).Nick & " :"
        Else
            SendWsock Index, ":" & ServerName & " 302 " & Users(Index).Nick & " :" & Replace(strcmd(i), "USERHOST ", "") & "=+" & User.ID
        End If
    ElseIf strcmd(i) Like "USER*" Then
'User
        Dim Ident As String, Email As String, Name As String, NewIdent As String * 10
        Ident = Replace(strcmd(i), "USER ", "")
        Ident = Mid(Ident, 1, InStr(1, Ident, " ") - 1)
        Email = Replace(strcmd(i), "USER " & Ident, "")
        Email = Mid(Email, 3)
        Email = Mid(Email, 1, InStr(1, Email, " "))
        Email = Replace(Email, Chr(34), "")
        Email = Mid(Email, 1, Len(Email) - 1)
        Email = Ident & "@" & Email
        Name = Mid(strcmd(i), InStr(1, strcmd(i), ":") + 1)
        Users(Index).Ident = Mid(Ident, 1, 10)
        Ident = Mid(Ident, 1, 10) & "@" & wsock(Index).RemoteHostIP
        Users(Index).Email = Email
        Users(Index).ID = Ident
        Users(Index).Name = Name
'Quit
    ElseIf strcmd(i) Like "QUIT*" Then
        Dim Quit As String
        Quit = Mid(strcmd(i), InStr(1, strcmd(i), " :") + 2)
        SendQuit CLng(Index), Quit
        wsock_Close (Index)
'Join
    ElseIf strcmd(i) Like "JOIN*" Then
        If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 75
        Dim Chan As String, ck As String, X As Long
        If Replace(strcmd(i), "JOIN ", "") = "0" Or Replace(strcmd(i), "JOIN ", "") = "#0" Then
            For X = 1 To Users(Index).Onchannels.Count
                SendPart CLng(Index), Users(Index).Onchannels.Item(1), ""
            Next X
            GoTo NextCmd
        End If
        If CountSpaces(strcmd(i)) = 3 Then
            Chan = Replace(strcmd(i), "JOIN ", "")
            Chan = Mid(Chan, 1, InStr(1, Chan, " ") - 1)
            ck = Replace(strcmd(i), "JOIN " & Chan & " ", "")
        Else
            Chan = Replace(strcmd(i), "JOIN ", "")
        End If
        Dim Chans() As String
        Chan = Replace(Chan, " ", "")
        Chans = Split(Chan, ",")
        For X = 0 To UBound(Chans)
            Chan = Chans(X)
            IsValidString Mid(Chan, 2)
            If Users(Index).Onchannels.Count >= MaxJoinChannels Then
                SendWsock Index, ":" & ServerName & " 432 * " & NewNick & ":You have joined too many Channels"
                GoTo NextCmd
            End If
            If Not Users(Index).IsOnChan(Chan) Then
                'If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 75
                If Not ChanExists(Chan) Then
                    Dim NewChannel As clsChannel
                    Set NewChannel = GetFreeChan
                    NewChannel.Name = Chan
                    NewChannel.Modes.Add "t", "t"
                    NewChannel.Modes.Add "n", "n"
                    NewChannel.Topic = DefTopic
                    NewChannel.Ops.Add Users(Index).Nick, Users(Index).Nick
                    NewChannel.All.Add Users(Index).Nick, Users(Index).Nick
                    Users(Index).Onchannels.Add Chan, Chan
                    SendWsock Index, ":" & Users(Index).Nick & " JOIN " & Chan, True
                    SendWsock Index, ":" & ServerName & " 353 " & Users(Index).Nick & " = " & Chan & " :" & Replace(NewChannel.GetOps & " " & NewChannel.GetVoices & " " & NewChannel.GetNorms, "  ", " "), True
                    SendWsock Index, ":" & ServerName & " 366 " & Users(Index).Nick & " " & Chan & " :End of /NAMES list.", True
                    NotifyJoin CLng(Index), Chan, False
                    SendLinks "JoinChan" & vbLf & Users(Index).Nick & vbLf & Chan
                Else
                    On Local Error Resume Next
                    Dim JoinChan As clsChannel
                    Set JoinChan = ChanToObject(Chan)
                    If (Not JoinChan.Key = "") Then
                        If Not JoinChan.Key = ck And Not Users(Index).IRCOp And (Not Users(Index).IsOwner(JoinChan.Name)) Then
                            SendWsock Index, ":" & ServerName & " 475 " & Users(Index).Nick & " " & Chan & " :Cannot join channel (+b)"
                            GoTo NextCmd
                        End If
                    End If
                    If (JoinChan.All.Count >= JoinChan.Limit And JoinChan.Limit <> 0) And Not Users(Index).IRCOp And (Not Users(Index).IsOwner(JoinChan.Name)) Then
                        SendWsock Index, ":" & ServerName & " 471 " & Users(Index).Nick & " " & Chan & " :Cannot join channel (+l)"
                        GoTo NextCmd
                    End If
                    If JoinChan.IsBanned(Users(Index)) And (Users(Index).IRCOp = False) And (JoinChan.IsException(Users(Index)) = False) And (Not Users(Index).IsOwner(JoinChan.Name)) Then
                        SendWsock Index, ":" & ServerName & " 474 " & Users(Index).Nick & " " & Chan & " :Cannot join channel (+b)"
                        GoTo NextCmd
                    End If
                    If JoinChan.IsMode("i") And (Users(Index).IRCOp = False) And (JoinChan.IsInvited2(Users(Index)) = False) And (JoinChan.IsInvited(Users(Index).Nick) = False) And (Not Users(Index).IsOwner(JoinChan.Name)) Then
                        SendWsock Index, ":" & ServerName & " 473 " & Users(Index).Nick & " " & Chan & " :Cannot join channel (+i)"
                        GoTo NextCmd
                    End If
                    NotifyJoin CLng(Index), Chan
                    JoinChan.NormUsers.Add Users(Index).Nick, Users(Index).Nick
                    JoinChan.All.Add Users(Index).Nick, Users(Index).Nick
                    Users(Index).Onchannels.Add Chan, Chan
                    SendWsock Index, ":" & Users(Index).Nick & " JOIN " & Chan, True
                    SendWsock Index, ":" & ServerName & " 353 " & Users(Index).Nick & " = " & Chan & " :" & FixNickList((Replace(JoinChan.GetOps & " " & JoinChan.GetVoices & " " & JoinChan.GetNorms, "  ", " "))), True
                    SendWsock Index, ":" & ServerName & " 366 " & Users(Index).Nick & " " & Chan & " :End of /NAMES list.", True
                    SendWsock Index, ":" & ServerName & " 332 " & Users(Index).Nick & " " & Chan & " :" & JoinChan.Topic, True
                    SendWsock Index, ":" & ServerName & " 333 " & JoinChan.TopicSetBy & " " & Chan & " " & JoinChan.TopicSetBy & " " & JoinChan.TopicSetOn, True
                End If
            End If
        Next X
'Part
    ElseIf strcmd(i) Like "PART*" Then
        Chan = Replace(strcmd(i), "PART ", "")
        If InStr(1, Chan, " ") Then Chan = Mid(Chan, 1, InStr(1, Chan, " ") - 1)
        Dim Reason As String
        Reason = Replace(strcmd(i), "PART " & Chan & " :", "")
        If Reason = strcmd(i) Then Reason = ""
        SendPart CLng(Index), Chan, Reason
'Mode
    ElseIf (strcmd(i) Like "MODE*") = True Then
        If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 250
        Dim Mode As String, ToUser() As String, Modes() As String, Op As String, ToUsers As String, Channel As clsChannel, Y As Long
        Chan = Replace(strcmd(i), "MODE ", "")
        If InStr(1, Chan, " ") <> 0 Then
            Chan = Mid(Chan, 1, InStr(1, Chan, " ") - 1)
        End If
        Set Channel = ChanToObject(Chan)
        Set User = NickToObject(Chan)
        If Not User Is Nothing Then
            Dim UM As String
            UM = Mid(Replace(strcmd(i), "MODE " & User.Nick & " ", "", , , vbTextCompare), 1, 1)
            Select Case UM
                Case "+"
                    AddUserMode User.Index, Mid(Replace(strcmd(i), "MODE " & User.Nick & " ", "", , , vbTextCompare), 2)
                Case "-"
                    RemoveUsermode User.Index, Mid(Replace(strcmd(i), "MODE " & User.Nick & " ", "", , , vbTextCompare), 2)
            End Select
            GoTo NextCmd
        End If
        Dim cmdline() As String, UserMode As Boolean
        cmdline = Split(strcmd(i), " ")
        For X = LBound(cmdline) To UBound(cmdline)
            If Not NickToObject(cmdline(X)) Is Nothing Then UserMode = True
        Next X
        If InStr(1, strcmd(i), "*") <> 0 Then UserMode = True
        If strcmd(i) Like "MODE * +?" Then UserMode = False
        If UserMode Then
            Mode = Replace(strcmd(i), "MODE " & Chan & " ", "")
            Op = Mid(Mode, 1, 1)
            Mode = Mid(Mode, 2, InStr(1, Mode, " ") - 2)
            ToUsers = Mid(strcmd(i), InStr(1, strcmd(i), Op) + Len(Mode) + 2)
            If InStr(1, ToUsers, " ") <> 0 Then
                ToUser = Split(" " & ToUsers, " ")
            Else
                ReDim ToUser(1)
                ToUser(1) = ToUsers
            End If
            If Channel.IsOp(Users(Index).Nick) = False And (Not Users(Index).IsOwner(Channel.Name)) And Not Users(Index).IRCOp Then
                SendWsock Index, ":" & ServerName & " 482 " & Users(Index).Nick & " " & Chan & " :You're not channel operator"
                GoTo NextCmd
            End If
            For X = 1 To Len(Mode)
                ReDim Preserve Modes(X)
                Modes(X) = Mid(Mode, X, 1)
            Next X
            ReDim Preserve Modes(UBound(ToUser))
            For Y = LBound(Modes) To UBound(Modes)
                If Not ToUser(Y) = "" Then
                    Select Case Modes(IIf((Y = 0), Y + 1, Y))
                        Case "o"
                            Select Case Op
                                Case "+"
                                    OpUser Channel, ToUser(Y), Users(Index).Nick
                                Case "-"
                                    DeOpUser Channel, ToUser(Y), Users(Index).Nick
                            End Select
                        Case "v"
                            Select Case Op
                                Case "+"
                                    VoiceUser Channel, ToUser(Y), Users(Index).Nick
                                Case "-"
                                    DeVoiceUser Channel, ToUser(Y), Users(Index).Nick
                            End Select
                        Case "b"
                            Select Case Op
                                Case "+"
                                    BanUser Channel, ToUser(Y), Users(Index).Nick
                                Case "-"
                                    UnBanUser Channel, ToUser(Y), Users(Index).Nick
                            End Select
                        Case "e"
                            Select Case Op
                                Case "+"
                                    ExceptionUser Channel, ToUser(Y), Users(Index).Nick
                                Case "-"
                                    UnExceptionUser Channel, ToUser(Y), Users(Index).Nick
                            End Select
                        Case "I"
                            Select Case Op
                                Case "+"
                                    InviteUser Channel, ToUser(Y), Users(Index).Nick
                                Case "-"
                                    UnInviteUser Channel, ToUser(Y), Users(Index).Nick
                            End Select
                    End Select
                End If
            Next Y
        Else
            If InStr(1, strcmd(i), " +b", vbBinaryCompare) <> 0 Then
                For X = 1 To Channel.Bans.Count
                    SendWsock Index, ":" & ServerName & " 367 " & Users(Index).Nick & " " & Channel.Name & " " & Channel.Bans(X)
                Next X
                SendWsock Index, ":" & ServerName & " 368 " & Users(Index).Nick & " " & Channel.Name & " :End of Channel Ban List"
            ElseIf InStr(1, strcmd(i), " +e", vbBinaryCompare) <> 0 Then
                For X = 1 To Channel.Exceptions.Count
                    SendWsock Index, ":" & ServerName & " 348 " & Users(Index).Nick & " " & Channel.Name & " " & Channel.Exceptions(X)
                Next X
                SendWsock Index, ":" & ServerName & " 349 " & Users(Index).Nick & " " & Channel.Name & " :End of Channel Exceptions List"
            ElseIf InStr(1, strcmd(i), " +I", vbBinaryCompare) <> 0 Then
                For X = 1 To Channel.Invites.Count
                    SendWsock Index, ":" & ServerName & " 346 " & Users(Index).Nick & " " & Channel.Name & " " & Channel.Invites(X)
                Next X
                SendWsock Index, ":" & ServerName & " 347 " & Users(Index).Nick & " " & Channel.Name & " :End of Channel Invites List"
            ElseIf InStr(1, strcmd(i), " +w", vbBinaryCompare) <> 0 Then
                SendWsock Index, ":" & ServerName & " 472 " & Users(Index).Nick & " w :is unknown mode char to me"
            ElseIf InStr(1, strcmd(i), "+") <> 0 Then
                AddChanModes Mid(strcmd(i), InStr(1, strcmd(i), "+") + 1), Chan, Users(Index)
            ElseIf InStr(1, strcmd(i), "-") <> 0 Then
                RemoveChanModes Mid(strcmd(i), InStr(1, strcmd(i), "-") + 1), Chan, Users(Index)
            Else
                SendWsock Index, ":" & ServerName & " 324 " & Users(Index).Nick & " " & Channel.Name & " " & Channel.GetModes
            End If
        End If
'Topic
    ElseIf (strcmd(i) Like "TOPIC*") = True Then
        If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 250
        If InStr(1, strcmd(i), " :") <> 0 Then
            Dim NewTopic As String
            Chan = Replace(strcmd(i), "TOPIC ", "")
            Chan = Mid(Chan, 1, InStr(1, Chan, " ") - 1)
            Set Channel = ChanToObject(Chan)
            If Channel Is Nothing Then GoTo NextCmd
            NewTopic = strcmd(i)
            NewTopic = Mid(NewTopic, InStr(1, NewTopic, ":") + 1)
            If Len(NewTopic) > TopicLen Then NewTopic = Left(NewTopic, TopicLen)
            If Channel.IsOp(Users(Index).Nick) = False And (Not Users(Index).IsOwner(Channel.Name)) And Channel.IsMode("t") Then
                SendWsock Index, ":" & ServerName & " 482 " & Users(Index).Nick & " :You're not channel operator"
                GoTo NextCmd
            End If
            SetTopic Chan, NewTopic, Users(Index).Nick
        Else
            Chan = Replace(strcmd(i), "TOPIC ", "")
            Set Channel = ChanToObject(Chan)
            If Channel Is Nothing Then GoTo NextCmd
            SendWsock Index, ":" & ServerName & " 332 " & Users(Index).Nick & " " & Chan & " :" & Channel.Topic
            SendWsock Index, ":" & ServerName & " 333 " & Channel.TopicSetBy & " " & Chan & " " & Users(Index).Nick & " " & Channel.TopicSetOn
        End If
'Invite
    ElseIf (strcmd(i) Like "INVITE*") = True Then
        Dim Target As String
        Target = Replace(strcmd(i), "INVITE ", "")
        Target = Mid(Target, 1, InStr(1, Target, " ") - 1)
        Chan = Mid(strcmd(i), Len("INVITE " & Target & " ") + 1)
        Set Channel = ChanToObject(Chan)
        If Channel.IsMode("i") And Channel.IsOp(Users(Index).Nick) = False Then
            SendWsock Index, ":" & ServerName & " 482 " & Users(Index).Nick & " " & Chan & " :You're not channel operator"
            GoTo NextCmd
        End If
        On Local Error Resume Next
        Channel.Invited.Add Target, Target
        SendWsock NickToObject(Target).Index, ":" & Users(Index).Nick & " INVITE " & Target & " " & Chan
        '1 = Command, 2 = Nick, 3 = Channel, 4 = Target
        SendLinks "AddInvite" & vbLf & Users(Index).Nick & vbLf & Channel.Name & vbLf & "" & vbLf & Target
'Kick
    ElseIf (strcmd(i) Like "KICK*") = True Then
        Dim Source As String
        Chan = Mid(strcmd(i), 6)
        Chan = Mid(Chan, 1, InStr(1, Chan, " ") - 1)
        Set Channel = ChanToObject(Chan)
        Source = Users(Index).Nick
        If Channel.IsOp(Source) = False And (Not Users(Index).IsOwner(Channel.Name)) Then
            SendWsock Index, ":" & ServerName & " 482 " & Users(Index).Nick & " " & Chan & " :You're not channel operator"
            GoTo NextCmd
        End If
        If InStr(1, strcmd(i), ":") <> 0 Then
            Reason = Mid(strcmd(i), InStr(1, strcmd(i), " :") + 2)
            If Len(Reason) > KickLen Then Reason = Left(Reason, KickLen)
            Target = Replace(strcmd(i), "KICK", "")
            Target = Mid(Target, 2)
            Target = Mid(Target, 1, InStr(1, Target, ":") - 2)
            Target = Replace(Target, Chan & " ", "")
            If Target = "ChanServ" Then
                SendSvrMsg Source & " tried to kick services[" & Chan & "]", True, ServerName
                SendWsock Index, ":" & ServerName & " 404 " & Source & " " & Chan & " :Cannot kick Services"
                GoTo NextCmd
            End If
            KickUser Source, Chan, Target, Reason, True
            GoTo NextCmd
        End If
        Target = Mid(strcmd(i), InStrRev(strcmd(i), " ", InStrRev(strcmd(i), " ")) + 1)
        If Target = "ChanServ" Then
            SendSvrMsg Source & " tried to kick services[" & Chan & "]", True, ServerName
            SendWsock Index, ":" & ServerName & " 404 " & Source & " " & Chan & " :Cannot kick Services"
            GoTo NextCmd
        End If
        KickUser Source, Chan, Target
'Pong
    ElseIf (strcmd(i) Like "PONG*") = True Then
        Users(Index).Ponged = True
        If Users(Index).NewUser Then
            SendWsock Index, GetWelcome(CLng(Index))
            SendWsock Index, ReadMotd(Users(Index).Nick)
            Users(Index).NewUser = False
            SendLogonNews Index
            SendLinks "NewUser" & vbLf & Users(Index).Nick & vbLf & Users(Index).Name & vbLf & Users(Index).DNS & vbLf & Users(Index).Ident & vbLf & Users(Index).Server & vbLf & Users(Index).ServerDescritption & vbLf & Users(Index).SignOn & vbLf & Users(Index).GID & vbLf & Users(Index).GetModes & vbLf & ServerName & " "
            CloneControl.Add wsock(Index).RemoteHostIP
            Users(Index).MsgsSent = 0
            tmrFloodProt(Index).Enabled = True
            If DefUserModes <> "" Then AddUserMode CLng(Index), DefUserModes
            '1 = Command, 2 = Nick, 3 = Name, 4 = DNS, 5 = Ident, 6 = Server, 7 = ServerDescription
        End If
'Ping
        ElseIf (strcmd(i) Like "PING*") = True Then
            SendWsock Index, "PONG " & Replace(strcmd(i), "PING ", ""), True
'PrivMsg
    ElseIf (strcmd(i) Like "PRIVMSG*") = True Then
        cmdline = Split(Mid(strcmd(i), 1, InStr(1, strcmd(i), " :")), " ")
        For X = LBound(cmdline) To UBound(cmdline)
            If Not NickToObject(cmdline(X)) Is Nothing Then UserMode = True
        Next X
        Target = Replace(strcmd(i), "PRIVMSG ", "")
        Target = Strings.Left(Target, InStr(1, Target, ":") - 2)
        cmdline = Split(Target, ",")
        For X = LBound(cmdline) To UBound(cmdline)
            Target = cmdline(X)
            If Not NickToObject(Target) Is Nothing Then UserMode = True
            Select Case LCase(Target)
                Case "chanserv"
                    UserMode = True
                Case "nickserv"
                   UserMode = True
                Case "memoserv"
                   UserMode = True
                Case "operserv"
                   UserMode = True
            End Select
            If (Not UserMode) Then
                Dim msgstr As String, Msg As String
                Chan = Target
                Set Channel = ChanToObject(Chan)
                If Channel Is Nothing Then
                    SendWsock Index, ":" & ServerName & " 404 " & Users(Index).Nick & " " & Chan & " :Cannot send to channel"
                    GoTo NextCmd
                End If
                If Channel.IsMode("n") Then
                    If (Not Channel.IsOnChan(Users(Index).Nick) And (Not Users(Index).IsOwner(Channel.Name)) And Not Users(Index).IRCOp) Then
                        SendWsock Index, ":" & ServerName & " 404 " & Users(Index).Nick & " " & Chan & " :Cannot send to channel"
                        GoTo NextCmd
                    End If
                End If
                If Channel.IsBanned(Users(Index)) Then
                   If ((Not Users(Index).IsOwner(Channel.Name)) And Not Users(Index).IRCOp) Then
                        SendWsock Index, ":" & ServerName & " 404 " & Users(Index).Nick & " " & Chan & " :Cannot send to channel"
                        GoTo NextCmd
                    End If
                End If
                If Channel.IsMode("m") Then
                    If Channel.IsOp(Users(Index).Nick) Then
                    ElseIf Channel.IsVoice(Users(Index).Nick) Or (Users(Index).IsOwner(Channel.Name)) And Not Users(Index).IRCOp Then
                    Else
                        SendWsock Index, ":" & ServerName & " 404 " & Users(Index).Nick & " " & Chan & " :Cannot send to channel"
                        GoTo NextCmd
                    End If
                End If
                Msg = strcmd(i)
                Msg = Mid(Msg, InStr(1, Msg, ":") + 1)
                If Len(Msg) > Msglen Then Msg = Left(Msg, Msglen)
                SendMsg Chan, Msg, Users(Index).Nick
            Else
                Target = Replace(strcmd(i), "PRIVMSG ", "")
                Target = Strings.Left(Target, InStr(1, Target, ":") - 2)
                Msg = strcmd(i)
                Msg = Mid(Msg, InStr(1, Msg, ":") + 1)
                If Len(Msg) > Msglen Then Msg = Left(Msg, Msglen)
                Set User = NickToObject(Target)
                If User Is Nothing Then
                    SendWsock Index, ":" & ServerName & " 401 " & Users(Index).Nick & " " & Target & " :No such nick/channel"
                    GoTo NextCmd
                End If
                SendMsg Target, Msg, Users(Index).Nick, False
            End If
        Next X
'Notice
    ElseIf (strcmd(i) Like "NOTICE*") = True Then
        Target = Replace(strcmd(i), "NOTICE ", "")
        Target = Replace(Target, ":*", " ")
        Target = Left(Target, InStr(1, Target, ":") - 2)
        Dim Targets() As String
        Targets = Split(Target, ",")
        Msg = strcmd(i)
        Msg = Mid(Msg, InStr(1, Msg, ":") + 1)
        If Len(Msg) > Msglen Then Msg = Left(Msg, Msglen)
        For Y = LBound(Targets) To UBound(Targets)
            Target = Targets(Y)
            If InStr(1, Target, "#") = 0 Then
                If NickToObject(Target) Is Nothing Then
                    SendWsock Index, ":" & ServerName & " 401 " & Users(Index).Nick & " :No such nick/channel"
                    GoTo NextCmd
                End If
                SendNotice Target, Msg, Users(Index).Nick
            Else
                Dim CurChan As clsChannel
                Set CurChan = ChanToObject(Target)
                If CurChan Is Nothing Then
                    SendWsock Index, ":" & ServerName & " 401 " & Users(Index).Nick & " :No such nick/channel"
                    GoTo NextCmd
                End If
                    If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + (10 * CurChan.All.Count)
                    SendNotice Target, Msg, Users(Index).Nick, True
            End If
        Next Y
'Motd
    ElseIf (strcmd(i) Like "MOTD") = True Then
        If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 1200
        SendWsock Index, ReadMotd(Users(Index).Nick)
'Whois
    ElseIf (strcmd(i) Like "WHOIS*") = True Then
        Dim WhoisStr As String, Nick As String
        Set User = NickToObject(Replace(strcmd(i), "WHOIS ", ""))
        If Not User Is Nothing Then
            SendWsock Index, User.GetWhois(Users(Index).Nick)
        Else
            SendWsock Index, ":" & ServerName & " 401 " & Users(Index).Nick & " :No such nick/channel"
        End If
'Away
    ElseIf (strcmd(i) Like "AWAY*") = True Then
        If Not Users(Index).Away Then
            Users(Index).AwayMsg = Replace(strcmd(i), "AWAY :", "")
            Users(Index).Away = True
            SendWsock Index, ":" & ServerName & " 306 " & Users(Index).Nick & " :You have been marked as being away"
            Users(Index).Modes.Add "a", "a"
        Else
            Users(Index).Away = False
            Users(Index).AwayMsg = ""
            RemoveUsermode CLng(Index), "a", True
        End If
'WallOps
    ElseIf (strcmd(i) Like "WALLOPS*") = True Then
        If Users(Index).IRCOp Then
            WallOps Replace(strcmd(i), "WALLOPS ", ""), Index
        Else
            SendWsock Index, ":" & ServerName & " 481 " & Users(Index).Nick & " :Permission Denied- You're not an IRC operator"
        End If
'Wall
    ElseIf (strcmd(i) Like "WALL*") = True Then
        If Users(Index).IRCOp Then
            Wall Replace(strcmd(i), "WALL ", ""), Index
        Else
            SendWsock Index, ":" & ServerName & " 481 " & Users(Index).Nick & " :Permission Denied- You're not an IRC operator"
        End If
'*****************************
'|      Client Queries             ||
'*****************************
'Version
    ElseIf (strcmd(i) Like "VERSION") = True Then
        SendWsock Index, GetWelcome(CLng(Index))
'Time
    ElseIf (strcmd(i) Like "TIME") = True Then
        SendWsock Index, ":" & ServerName & " 391" & Users(Index).Nick & " " & ServerName & " :" & Now
'Info
    ElseIf (strcmd(i) Like "VERSION") = True Then
        SendWsock Index, GetWelcome(CLng(Index))
'IsOn
    ElseIf (strcmd(i) Like "ISON*") = True Then
        Dim strIsOn As String, LoggedIn() As String, IsOnArr() As String
        ReDim LoggedIn(1)
        strIsOn = Replace(strcmd(i), "ISON ", "")
        IsOnArr = Split(strIsOn, " ")
        For X = LBound(IsOnArr) To UBound(IsOnArr)
            If Not NickToObject(IsOnArr(X)) Is Nothing Then
                ReDim Preserve LoggedIn(UBound(LoggedIn) + 1)
                LoggedIn(UBound(LoggedIn)) = IsOnArr(X)
            End If
        Next X
        strIsOn = Join(LoggedIn, " ")
        SendWsock Index, (":" & ServerName & " 303 " & Users(Index).Nick & " :" & strIsOn) 'True
'Lusers
    ElseIf (strcmd(i) Like "LUSERS*") = True Then
        If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 400
        SendWsock Index, ":" & ServerName & " 252 " & Users(Index).Nick & " :" & Operators & " Operator(s) online" & vbCrLf & _
                                           ":" & ServerName & " 254 " & Users(Index).Nick & " :channels formed = " & ChanCount & vbCrLf & _
                                           ":" & ServerName & " 255 " & Users(Index).Nick & " :I have " & UserCount - 4 & " clients and " & (CurLinkCount + 1) & " Servers" & vbCrLf & _
                                           ":" & ServerName & " 265 " & Users(Index).Nick & " :Current Local Users : " & (UserCount - 4) & " Max Local Users : " & MaxUser & vbCrLf & _
                                           ":" & ServerName & " 266 " & Users(Index).Nick & " :Current Global Users: " & CurGlobalUsers & " Max Global Users: " & MaxGlobalUsers & vbCrLf
'Stats
    ElseIf (strcmd(i) Like "STATS*") = True Then
        If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 250
        Dim StatsParam As String
        StatsParam = Replace(strcmd(i), "STATS ", "")
        Select Case StatsParam
            Case "u"
                SendWsock Index, ":" & ServerName & " 242 " & Users(Index).Nick & " :" & CStr(Started)
                SendWsock Index, ":" & ServerName & " 250 " & Users(Index).Nick & " :Highest Connection Count: " & MaxUser
                SendWsock Index, ":" & ServerName & " 219 " & Users(Index).Nick & " u :End of /STATS report"
        End Select
'Info
    ElseIf (strcmd(i) Like "INFO*") = True Then
        If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 250
        SendWsock Index, ":" & ServerName & " 371 " & Users(Index).Nick & " :" & ServerName & " running tsIRCd-2.0rc1"
        SendWsock Index, ":" & ServerName & " 371 " & Users(Index).Nick & " :This server was created 2014-10-02 by Thomas Schilb (thomasschilb@gmx.net)"
        SendWsock Index, ":" & ServerName & " 374 " & Users(Index).Nick & " :End of INFO list"
'Links
    ElseIf (strcmd(i) Like "LINKS*") = True Then
        If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 200
        SendWsock Index, ":" & ServerName & " 364 " & Users(Index).Nick & " " & ServerName & " " & ServerName & " :0 " & ServerDesc
        SendWsock Index, ":" & ServerName & " 365 " & Users(Index).Nick & " * :End of /LINKS list"
'Names
    ElseIf strcmd(i) Like "NAMES*" Then
        If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 150
        Chan = Replace(strcmd(i), "NAMES ", "")
        Set Channel = ChanToObject(Chan)
        SendWsock Index, ":" & ServerName & " 353 " & Users(Index).Nick & " = " & Chan & " :" & FixNickList((Replace(Channel.GetOps & " " & Channel.GetVoices & " " & Channel.GetNorms, "  ", " ")))
        SendWsock Index, ":" & ServerName & " 366 " & Users(Index).Nick & " " & Chan & " :End of /NAMES list."
'List
    ElseIf strcmd(i) Like "LIST*" Then
        If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 1000
        SendWsock Index, ":" & ServerName & " 321 " & Users(Index).Nick & " Channel :Users  Name"
        SendWsock Index, GetChanList(Users(Index).Nick)
        SendWsock Index, ":" & ServerName & " 323 " & Users(Index).Nick & " :End of /LIST"
'Admin
    ElseIf strcmd(i) Like "ADMIN*" Then
        If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 250
        SendWsock Index, ":" & ServerName & " 256 " & Users(Index).Nick & " :Administrative info about " & ServerName
        SendWsock Index, ":" & ServerName & " 257 " & Users(Index).Nick & " :" & ServerDesc
        SendWsock Index, ":" & ServerName & " 258 " & Users(Index).Nick & " :" & AdminName
        SendWsock Index, ":" & ServerName & " 259 " & Users(Index).Nick & " :" & AdminEmail
'Who
    ElseIf strcmd(i) Like "WHO*" Then
        Dim strSearch As String
        strSearch = Replace(strcmd(i), "WHO ", "")
        For X = 1 To UBound(Users)
            DoEvents
            If Not Users(X) Is Nothing Then
                If (Users(X).Nick Like strSearch) Then
                    SendWsock Index, ":" & ServerName & " 352 * " & LCase(Users(Index).Nick) & " " & Users(X).Nick & " " & Users(X).DNS & " " & ServerName & " " & Users(X).Nick & " H :" & Users(X).Hops & " " & Users(X).Name
                End If
            End If
        Next X
        SendWsock Index, ":" & ServerName & " " & 315 & " " & Users(Index).Nick & " " & strSearch & " :END of /WHO list."
'352 Dilligent #darkmyst ~dilligent p508F2CEA.dip.t-dialin.net *.quakenet.org Dilligent H@ :0 Dennis Fisch
'352 Dilligent * ~dilligent p508F2CEA.dip.t-dialin.net *.quakenet.org Dilligent H :0 Dennis Fisch
'315 Dilligent Dilligent :End of /WHO list.
'*****************************
'|      Operator Queries        ||
'*****************************
'Oper
    ElseIf (strcmd(i) Like "OPER *") = True Then
        Dim PW As String, UserName As String
        UserName = Replace(strcmd(i), "OPER ", "")
        PW = Mid(UserName, InStr(1, UserName, " ") + 1)
        UserName = Mid(UserName, 1, InStr(1, UserName, " ") - 1)
        PW = Replace(PW, ":", "")
        If Not HasOline(Users(Index).Nick, Users(Index).GetMask) Then
            SendWsock Index, ":" & ServerName & " 491 " & Users(Index).Nick & " :No O-lines for your host"
            GoTo NextCmd
        End If
        With Olines(GetOline(Users(Index).DNS))
            If Not Users(Index).Nick = .UserName Then
                SendWsock Index, ":" & ServerName & " 491 " & Users(Index).Nick & " :your nickname must match the nickname with which the O-Line has been created"
                GoTo NextCmd
            End If
            If Not PW = .Password Then
                SendWsock Index, ":" & ServerName & " 464 " & Users(Index).Nick & " :Password incorrect"
                GoTo NextCmd
            End If
            SendWsock Index, ":" & ServerName & " 381 " & Users(Index).Nick & " :You are now an IRC operator"
            SendLinks "ModeUser" & vbLf & Users(Index).Nick & vbLf & "+" & vbLf & "o"
            SendWsock Index, ":" & Users(Index).Nick & " MODE " & Users(Index).Nick & " +o"
            On Local Error Resume Next
            AddUserMode CLng(Index), "o"
            Users(Index).AddModes "o"
            Users(Index).IRCOp = True
            SendSvrMsg Users(Index).Nick & " is now Operator", True
            Users(Index).RealDNS = Users(Index).DNS
            Users(Index).DNS = ServerName
            Operators = Operators + 1
            'WallOps " is now Operator", 1
        End With
'Restart
    ElseIf (strcmd(i) = "RESTART") = True Then
        If Users(Index).IRCOp Then
            For X = LBound(Users) To UBound(Users)
                If Not Users(X) Is Nothing Then SendNotice "", "*** Global -- " & "Recieved RESTART command from " & Users(Index).Nick, "GLOBAL", , CInt(X)
            Next X
            Wait 2000
            Restart
        Else
            SendWsock Index, ":" & ServerName & " 481 " & Users(Index).Nick & " :Permission Denied- You're not an IRC operator"
        End If
'Die
    ElseIf strcmd(i) Like "DIE*" Then
        If Users(Index).IRCOp Then
            For X = LBound(Users) To UBound(Users)
                If Not Users(X) Is Nothing Then SendNotice "", "*** Global -- " & "Recieved DIE command from " & Users(Index).Nick, "GLOBAL", , CInt(X)
            Next X
            Wait 2000
            Unload Me
        Else
            SendWsock Index, ":" & ServerName & " 481 " & Users(Index).Nick & " :Permission Denied- You're not an IRC operator"
        End If
'K-line
    ElseIf (strcmd(i) Like "KLINE*") = True Then
        If Not Users(Index).IRCOp Then
            SendWsock Index, ":" & ServerName & " 481 " & Users(Index).Nick & " :Permission Denied- You're not an IRC operator"
            GoTo NextCmd
        End If
        Klines.Add Replace(strcmd(i), "KLINE ", "")
'Kill
    ElseIf (strcmd(i) Like "KILL*") = True Then
        If Users(Index).IRCOp Then
            Dim NickName As String, Comment As String
            NickName = Replace(strcmd(i), "KILL ", "")
            If InStr(1, NickName, " ") = 0 Then
                Comment = Users(Index).Nick
            Else
                NickName = Mid(NickName, 1, InStr(1, NickName, " :") - 1)
                Comment = Replace(strcmd(i), "KILL " & NickName & " :", "")
            End If
            Set User = NickToObject(NickName, , True)
            If Not User Is Nothing Then
                SendLinks "KillUser" & vbLf & User.Nick & vbLf & Comment
                SendQuit User.Index, "Killed by " & Users(Index).Nick & " (" & Comment & ")", True
                If Not User.LocalUser Then GoTo NextCmd
                User.Killed = True
                SendWsock User.Index, ":" & Users(Index).Nick & "!" & Users(Index).Ident & "@" & ServerName & " KILL " & User.Nick & " :" & Comment, True
                SendWsock User.Index, "ERROR :Closing Link: " & User.Nick & "[" & frmMain.wsock(User.Index).RemoteHostIP & ".] " & ServerName & " (" & Comment & ")", True
                'K-Line (Ban) User from Network for 10 seconds
                Dim Kline As Long
                Kline = GetRand
                Load tmrKlined(Kline)
                tmrKlined(Kline).Tag = wsock(User.Index).RemoteHostIP
                tmrKlined(Kline).Enabled = True
                On Local Error Resume Next
                Klines.Add wsock(User.Index).RemoteHostIP, wsock(User.Index).RemoteHostIP
                SendNotice Users(Index).Nick, User.Nick & " has been removed from the network", "" & ServerName & ""
                SendSvrMsg "Recieved Kill message for " & User.Nick & "!" & User.Ident & "@" & User.DNS & " Path: " & Users(Index).Nick & " (" & Comment & ")", True
            Else
                SendWsock Index, ":" & ServerName & " 401 " & Users(Index).Nick & " :No such nick/channel, using wildcards instead"
                For X = 1 To UBound(Users)
                    If Not Users(X) Is Nothing Then
                        If (Users(X).Nick & "!" & Users(X).Ident & "@" & Users(X).DNS) Like NickName Then
                            Set User = Users(X)
                            SendLinks "KillUser" & vbLf & User.Nick & vbLf & Comment
                            SendQuit User.Index, "Killed by " & Users(Index).Nick & " (" & Comment & ")", True
                            If Not User.LocalUser Then GoTo NextCmd
                            User.Killed = True
                            SendWsock User.Index, ":" & Users(Index).Nick & "!" & Users(Index).Ident & "@" & ServerName & " KILL " & User.Nick & " :" & Comment, True
                            SendWsock User.Index, "ERROR :Closing Link: " & User.Nick & "[" & frmMain.wsock(User.Index).RemoteHostIP & ".] " & ServerName & " (" & Comment & ")", True
                            SendQuit User.Index, "Killed by " & Users(Index).Nick & " (" & Comment & ")", True
                            SendSvrMsg "Recieved Kill message for " & User.Nick & "!" & User.Ident & "@" & User.DNS & " Path: " & Users(Index).Nick & " (" & Comment & ")", True
                        End If
                    End If
                Next X
            End If
        Else
            SendWsock Index, ":" & ServerName & " 481 " & Users(Index).Nick & " :Permission Denied- You're not an IRC operator"
        End If
'AKill
    ElseIf (strcmd(i) Like "AKILL*") = True Then
        If Users(Index).IRCOp Then
            NickName = Replace(strcmd(i), "AKILL ", "")
            If InStr(1, NickName, " ") = 0 Then
                Comment = Users(Index).Nick
            Else
                NickName = Mid(NickName, 1, InStr(1, NickName, " ") - 1)
                Comment = Replace(strcmd(i), "AKILL " & NickName & " ", "")
            End If
            Set User = NickToObject(NickName, , True)
            If Not User Is Nothing Then
                SendLinks "KillUser" & vbLf & User.Nick & vbLf & Comment
                SendQuit User.Index, "AKilled by " & Users(Index).Nick & " (" & Comment & ")", True
                If Not User.LocalUser Then GoTo NextCmd
                User.Killed = True
                SendWsock User.Index, ":" & Users(Index).Nick & "!" & Users(Index).Ident & "@" & ServerName & " KILL " & User.Nick & " :" & Comment, True
                SendWsock User.Index, "ERROR :Closing Link: " & User.Nick & "[" & frmMain.wsock(User.Index).RemoteHostIP & ".] " & ServerName & " (" & Comment & ")", True
                Klines.Add wsock(User.Index).RemoteHostIP, wsock(User.Index).RemoteHostIP
'                Wait 300
'                wsock_Close (User.Index)
                SendNotice Users(Index).Nick, User.Nick & " has been removed from the network", "" & ServerName & ""
                SendSvrMsg "Recieved AKill message for " & User.Nick & "!" & User.Ident & "@" & User.DNS & " Path: " & Users(Index).Nick & " (" & Comment & ")", True
            Else
                SendWsock Index, ":" & ServerName & " 401 " & Users(Index).Nick & " :No such nick/channel, using wildcards instead"
                For X = 5 To UBound(Users)
                    If Not Users(X) Is Nothing Then
                        If (Users(X).Nick & "!" & Users(X).Ident & "@" & Users(X).DNS) Like NickName Then
                            Set User = Users(X)
                            SendLinks "KillUser" & vbLf & User.Nick & vbLf & Comment
                            SendQuit User.Index, "AKilled by " & Users(Index).Nick & " (" & Comment & ")", True
                            If Not User.LocalUser Then GoTo NextCmd
                            User.Killed = True
                            SendWsock User.Index, ":" & Users(Index).Nick & "!" & Users(Index).Ident & "@" & ServerName & " KILL " & User.Nick & " :" & Comment, True
                            SendWsock User.Index, "ERROR :Closing Link: " & User.Nick & "[" & frmMain.wsock(User.Index).RemoteHostIP & ".] " & ServerName & " (" & Comment & ")", True
                            On Local Error Resume Next
                            Klines.Add wsock(User.Index).RemoteHostIP, wsock(User.Index).RemoteHostIP
                            SendQuit User.Index, "AKilled by " & Users(Index).Nick & " (" & Comment & ")", True
                            SendSvrMsg "Recieved AKill message for " & User.Nick & "!" & User.Ident & "@" & User.DNS & " Path: " & Users(Index).Nick & " (" & Comment & ")", True
                        End If
                    End If
                Next X
            End If
        Else
            SendWsock Index, ":" & ServerName & " 481 " & Users(Index).Nick & " :Permission Denied- You're not an IRC operator"
        End If
'Rehash
    ElseIf strcmd(i) Like "REHASH*" Then
        If Users(Index).IRCOp Then
            Rehash Users(Index).Nick
        Else
            SendWsock Index, ":" & ServerName & " 481 " & Users(Index).Nick & " :Permission Denied- You're not an IRC operator"
        End If
'ClientInfo
    ElseIf strcmd(i) Like "CLIENTINFO*" Then
        If Users(Index).IRCOp Then
            Set User = NickToObject(Replace(strcmd(i), "CLIENTINFO ", ""))
            SendWsock Index, ":NickName = " & User.Nick
            SendWsock Index, ":Ident = " & User.Ident
            SendWsock Index, ":Name = " & User.Name
            SendWsock Index, ":Email = " & User.Email
            SendWsock Index, ":Modes = " & User.GetModes
        Else
            SendWsock Index, ":" & ServerName & " 481 " & Users(Index).Nick & " :Permission Denied- You're not an IRC operator"
        End If
'DeleteUserEntry
    ElseIf strcmd(i) Like "DELETE*" Then
        If Users(Index).IRCOp Then
            Set Users(NickToObject(Replace(strcmd(i), "DELETE ", "")).Index) = Nothing
        Else
            SendWsock Index, ":" & ServerName & " 481 " & Users(Index).Nick & " :Permission Denied- You're not an IRC operator"
        End If
'Connect
    ElseIf strcmd(i) Like "CONNECT*" Then
        Dim SName As String, SPort As String
        SName = Replace(strcmd(i), "CONNECT ", "")
        SPort = Mid(SName, InStr(1, SName, " ") + 1)
        SName = Replace(SName, " " & SPort, "")
        If SPort = 0 Then SPort = 6668
        If Users(Index).IRCOp Then
            Dim LinkCount As Long
            LinkCount = Link.Count + 1
            CurLinkCount = CurLinkCount + 1
            MaxLinkCount = MaxLinkCount + 1
            Load Link(LinkCount)
            Link(LinkCount).LocalPort = 0
            Link(LinkCount).Connect SName, SPort
            Link(LinkCount).Tag = SName
        Else
            SendWsock Index, ":" & ServerName & " 481 " & Users(Index).Nick & " :Permission Denied- You're not an IRC operator"
        End If
'SQuit
    ElseIf strcmd(i) Like "SQUIT*" Then
        If Users(Index).IRCOp Then
            Dim CloseLink As Long
            SName = Replace(strcmd(i), "SQUIT ", "")
            For X = 2 To Link.UBound
                On Error Resume Next
                If Link(X).Tag = SName Then
                    SendSvrMsg "Link closed by " & Users(Index).Nick & "[" & ServerName & " -- " & Link(X).Tag & "]", True
                    Link(X).Close
                    Link_Close (X)
                    Unload Link(X)
                    Exit For
                End If
                On Error GoTo 0
            Next X
        Else
            SendWsock Index, ":" & ServerName & " 481 " & Users(Index).Nick & " :Permission Denied- You're not an IRC operator"
        End If
'*****************************
'|     Service Commands    ||
'*****************************
'NickServ
    ElseIf (strcmd(i) Like "NS*") = True Then
        SendMsg "NickServ", Replace(strcmd(i), "NS ", ""), Users(Index).Nick, False
    ElseIf (strcmd(i) Like "NICKSERV*") = True Then
        SendMsg "NickServ", Replace(strcmd(i), "NickServ ", ""), Users(Index).Nick, False
'MemoServ
    ElseIf (strcmd(i) Like "MS*") = True Then
        SendMsg "MemoServ", Replace(strcmd(i), "MS ", ""), Users(Index).Nick, False
    ElseIf (strcmd(i) Like "MEMOSERV*") = True Then
        SendMsg "MemoServ", Replace(strcmd(i), "MemoServ ", ""), Users(Index).Nick, False
'ChanServ
    ElseIf (strcmd(i) Like "CS*") = True Then
        SendMsg "ChanServ", Replace(strcmd(i), "CS ", ""), Users(Index).Nick, False
    ElseIf (strcmd(i) Like "CHANSERV*") = True Then
        SendMsg "ChanServ", Replace(strcmd(i), "ChanServ ", ""), Users(Index).Nick, False
'OperServ
    ElseIf (strcmd(i) Like "OS*") = True Then
        ParseOSCmd strcmd(i), CLng(Index)
    ElseIf (strcmd(i) Like "OPERSERV*") = True Then
        ParseOSCmd Replace(strcmd(i), "OPERSERV", "OS"), CLng(Index)
    ElseIf strcmd(i) = "" Then
    Else
        If InStr(1, strcmd(i), " ") <> 0 Then strcmd(i) = Mid(strcmd(i), 1, InStr(1, strcmd(i), " ") - 1)
        SendWsock Index, ":" & ServerName & " 421 " & Users(Index).Nick & " :" & strcmd(i) & " Unknown command"
    End If
NextCmd:
Next i
parseerr:
On Error Resume Next
If Index < 5 Then Exit Sub
If Not Users(Index) Is Nothing Then SendWsock Index, ":" & ServerName & " 421 " & Users(Index).Nick & " :Parsing error | Need more Parameters or wrong order of Parameters" & Err.Description
End Sub

Private Sub wsock_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
SendQuit CLng(Index), "Connection error: " & Description, False
wsock_Close (Index)
End Sub

Private Sub wsock_SendComplete(Index As Integer)
On Error Resume Next
If Users(Index).Killed Then wsock_Close (Index)
End Sub
