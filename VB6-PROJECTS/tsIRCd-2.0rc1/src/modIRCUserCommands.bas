Attribute VB_Name = "modIRCUserCommands"
Option Explicit
Dim NickName As String, level As Long

Public Function ChangeNick(Index As Long, NewNick As String, Optional SendLink As Boolean = True) As Boolean
On Error Resume Next
If NickInUse(NewNick) Then
    ChangeNick = False
    Exit Function
End If
If UCase(Users(Index).Nick) = UCase(NewNick) Then
    ChangeNick = True
    Exit Function
End If
If Users(Index).Nick = "" Then
    SendWsock Index, ":" & NewNick & " NICK " & NewNick
    ChangeNick = True
Else
    Dim i As Long, X As Long, CurChan As clsChannel, CurNick As clsUser
    For i = 1 To Users(Index).Onchannels.Count
        Set CurChan = ChanToObject(Users(Index).Onchannels(i))
        If CurChan.IsNorm(Users(Index).Nick) Then
            CurChan.NormUsers.Remove Users(Index).Nick
            CurChan.NormUsers.Add NewNick, NewNick
        ElseIf CurChan.IsVoice(Users(Index).Nick) Then
            CurChan.Voices.Remove Users(Index).Nick
            CurChan.Voices.Add NewNick, NewNick
        ElseIf CurChan.IsOp(Users(Index).Nick) Then
            CurChan.Ops.Remove Users(Index).Nick
            CurChan.Ops.Add NewNick, NewNick
        End If
        CurChan.All.Remove Users(Index).Nick
        CurChan.All.Add NewNick, NewNick
        For X = 1 To (CurChan.All.Count)
            If Not NewNick = CurChan.All(X) Then SendWsock NickToObject(CurChan.All(X)).Index, ": " & Users(Index).Nick & " NICK " & NewNick
        Next X
    Next i
    SendWsock Index, ":" & Users(Index).Nick & " NICK " & NewNick
    ChangeNick = True
End If
'1 = Command, 2 = Nick, 3 = NewNick
If SendLink Then SendLinks "Nick" & vbLf & Users(Index).Nick & vbLf & NewNick
Users(Index).Nick = NewNick
ChangeNick = True
End Function

Public Sub SendMsg(Target As String, Message As String, User As String, Optional SendToChan As Boolean = True, Optional SendLink As Boolean = True)
On Error Resume Next
Dim Index As Long
Index = NickToObject(User, 5).Index
NickToObject(User, 5).Idle = UnixTime
If SendToChan Then
    Dim i As Long, X As Long, CurChan As clsChannel
    Set CurChan = ChanToObject(Target)
    If CurChan.IsMode("c") Then StripColorCodes Message
    For X = 1 To (CurChan.All.Count)
        If Not User = CurChan.All(X) Then SendWsock NickToObject(CurChan.All(X), 5, True).Index, ":" & Users(Index).Nick & "!" & Users(Index).Ident & "@" & Users(Index).DNS & " PRIVMSG " & Target & " :" & Message
    Next X
    '1 = Command, 2 = Nick, 3 = Channel, 4 = Msg
    If SendLink Then SendLinks "PrivMsgChan" & vbLf & User & vbLf & Target & vbLf & Message
Else
    Dim cUser As clsUser
    Set cUser = NickToObject(Target, 5)
    If cUser.LocalUser Then SendLink = False
    SendWsock cUser.Index, ":" & Users(Index).Nick & "!" & Users(Index).Ident & "@" & Users(Index).DNS & " PRIVMSG " & Target & " :" & Message & vbCrLf
    '1 = Command, 2 = Nick, 3 = Target, 4 = Msg
    If SendLink Then SendLinks "PrivMsgUser" & vbLf & User & vbLf & Target & vbLf & Message
End If
End Sub

Public Sub SendQuit(Index As Long, QuitMsg As String, Optional Kill As Boolean = False, Optional SendLink As Boolean = True)
On Error Resume Next
Dim i As Long, X As Long, CurChan As clsChannel
If QuitMsg = "" Then QuitMsg = DefQuit
Users(Index).SentQuit = True
For i = 1 To Users(Index).Onchannels.Count
    Set CurChan = ChanToObject(Users(Index).Onchannels(i))
    For X = 1 To (CurChan.All.Count)
        If Not Users(Index).Nick = CurChan.All(X) Then SendWsock NickToObject(CurChan.All(X)).Index, ": " & Users(Index).Nick & "!" & Users(Index).Ident & "@" & Users(Index).DNS & " QUIT :" & QuitMsg
    Next X
    CurChan.All.Remove Users(Index).Nick
    If CurChan.IsNorm(Users(Index).Nick) Then
        CurChan.NormUsers.Remove Users(Index).Nick
    ElseIf CurChan.IsVoice(Users(Index).Nick) Then
        CurChan.Voices.Remove Users(Index).Nick
    ElseIf CurChan.IsOp(Users(Index).Nick) Then
        CurChan.Ops.Remove Users(Index).Nick
    End If
Next i
If Kill Then
    SendWsock Index, ":" & Users(Index).Nick & "!" & Users(Index).Ident & "@" & Users(Index).DNS & " KILL :" & QuitMsg
    SendWsock Index, "ERROR :Closing Link: " & "You have been killed" & vbCrLf
    '1 = Command, 2 = Nick, 3 = Reason
    If SendLink Then SendLinks "KillUser" & vbLf & Users(Index).Nick & vbLf & QuitMsg
Else
    '1 = Command, 2 = Nick, 3 = QuitMsg
    If SendLink Then SendLinks "QuitUser" & vbLf & Users(Index).Nick & vbLf & QuitMsg
End If
End Sub

Public Sub SendPart(Index As Long, Chan As String, Reason As String, Optional SendLink As Boolean = True)
On Error Resume Next
Dim i As Long, X As Long, CurChan As clsChannel, Found As Boolean
Users(Index).Idle = UnixTime
For i = 1 To Users(Index).Onchannels.Count
    If Chan = Users(Index).Onchannels(i) Then Found = True
Next i
If Found = False Then Exit Sub
'1 = Command, 2 = Nick, 3 = Channel, 4 = Reason
If SendLink Then SendLinks "PartUser" & vbLf & Users(Index).Nick & vbLf & Chan & vbLf & Reason
Set CurChan = ChanToObject(Chan)
For X = 1 To (CurChan.All.Count)
    If Not Users(Index).Nick = CurChan.All(X) Then SendWsock NickToObject(CurChan.All(X)).Index, ": " & Users(Index).Nick & "!" & Users(Index).Ident & "@" & Users(Index).DNS & " PART " & Chan & " :" & Reason
Next X
CurChan.All.Remove Users(Index).Nick
If CurChan.IsNorm(Users(Index).Nick) Then
    CurChan.NormUsers.Remove Users(Index).Nick
ElseIf CurChan.IsVoice(Users(Index).Nick) Then
    CurChan.Voices.Remove Users(Index).Nick
ElseIf CurChan.IsOp(Users(Index).Nick) Then
    CurChan.Ops.Remove Users(Index).Nick
End If
Users(Index).Onchannels.Remove Chan
SendWsock Index, ":" & Users(Index).Nick & " PART " & Chan
If CurChan.All.Count = 0 And (CurChan.IsMode("r")) Then
    Do Until CurChan.Invited.Count = 0
        CurChan.Invited.Remove 1
    Loop
    CurChan.TopicSetBy = ""
    CurChan.TopicSetBy = 0
    CurChan.Topic = ""
ElseIf CurChan.All.Count = 0 And (Not CurChan.IsMode("r")) Then
    Set Channels(CurChan.Index) = Nothing
End If
End Sub

Public Sub SendPing(Index As Long)
SendWsock Index, "PING " & GetRandom
End Sub

Public Sub NotifyJoin(Index As Long, Chan As String, Optional SendLink As Boolean = True)
Dim i As Long, Channel As clsChannel
Set Channel = ChanToObject(Chan)
If Channel.IsOnChan(Users(Index).Nick) Then Exit Sub
For i = 1 To Channel.All.Count
    If Not Channel.All(i) = Users(Index).Nick Or Channel.All(i) = "" Then SendWsock NickToObject(Channel.All(i)).Index, ":" & Users(Index).Nick & "!" & Users(Index).Ident & "@" & Users(Index).DNS & " JOIN " & Chan
Next i
'1 = Command, 2 = Nick, 3 = Channel
If SendLink Then SendLinks "JoinChan" & vbLf & Users(Index).Nick & vbLf & Chan
End Sub

Public Sub SendNotice(Target As String, Message As String, User As String, Optional ToChannel As Boolean = False, Optional Index As Integer, Optional SendLink As Boolean = True)
On Error Resume Next
Dim TargetIndex As Long, i As Long
If Index <> 0 Then
    TargetIndex = Index
Else
    TargetIndex = NickToObject(Target).Index
End If
If Not ToChannel Then
    If Users(TargetIndex).LocalUser Then SendLink = False
    SendWsock TargetIndex, ":" & User & " NOTICE " & Target & " :" & Message
    '1 = Command, 2 = Nick, 3 = Target, 4 = Msg
    If SendLink Then SendLinks "NoticeUser" & vbLf & User & vbLf & Target & vbLf & Message
Else
    Dim Chan As clsChannel
    Set Chan = ChanToObject(Target)
    For i = 1 To Chan.All.Count
        SendWsock NickToObject(Chan.All(i)).Index, ":" & User & " NOTICE " & Target & " :" & Message
    Next i
    '1 = Command, 2 = Nick, 3 = Target, 4 = Msg
    If SendLink Then SendLinks "NoticeChan" & vbLf & User & vbLf & Target & vbLf & Message
End If
End Sub

Public Sub KickUser(Source As String, Chan As String, Target As String, Optional Reason As String, Optional Reasoning As Boolean = False, Optional SendLink As Boolean = True)
On Error Resume Next
Dim i As Long, Channel As clsChannel, KickMsg As String
Set Channel = ChanToObject(Chan)
If Reasoning = True Then
    KickMsg = ":" & Source & " KICK " & Chan & " " & Target & " :" & Reason
Else
    KickMsg = ":" & Source & " KICK " & Chan & " " & Target
End If
For i = 1 To Channel.All.Count
     SendWsock NickToObject(Channel.All(i)).Index, KickMsg
Next i
Channel.All.Remove Target
If Channel.IsNorm(Target) Then
    Channel.NormUsers.Remove Target
ElseIf Channel.IsVoice(Target) Then
    Channel.Voices.Remove Target
ElseIf Channel.IsOp(Target) Then
    Channel.Ops.Remove Target
End If
NickToObject(Target).Onchannels.Remove Chan
'1 = Command, 2 = Nick, 3 = Channel, 4 = Reason, 5 = Target
If SendLink Then SendLinks "KickUser" & vbLf & Source & vbLf & Channel.Name & vbLf & Reason & vbLf & Target
End Sub

Public Sub SetTopic(Chan As String, NewTopic As String, User As String, Optional SendLink As Boolean = True)
Dim i As Long, Channel As clsChannel
Set Channel = ChanToObject(Chan)
If Channel.Topic = NewTopic Then Exit Sub
For i = 1 To Channel.All.Count
     SendWsock NickToObject(Channel.All(i)).Index, ":" & User & " TOPIC " & Chan & " :" & NewTopic
Next i
Channel.Topic = NewTopic
Channel.TopicSetOn = UnixTime
Channel.TopicSetBy = User
'1 = Command, 2 = Nick, 3 = Channel, 4 = Target
If SendLink Then SendLinks "SetTopic" & vbLf & User & vbLf & Channel.Name & vbLf & "" & vbLf & NewTopic
End Sub

Public Sub OpUser(Channel As clsChannel, Target As String, User As String, Optional OpAnyway As Boolean = False, Optional SendLink As Boolean = True)
On Error Resume Next
Dim i As Long, Chan As String
Chan = Channel.Name
If Channel.IsOp(Target) And OpAnyway = False Then Exit Sub
For i = 1 To Channel.All.Count
     SendWsock NickToObject(Channel.All(i)).Index, (":" & User & " MODE " & Chan & " +o " & Target)
Next i
If Channel.IsNorm(Target) Then Channel.NormUsers.Remove Target
Channel.Ops.Add Target, Target
'1 = Command, 2 = Nick, 3 = Channel, 4 = Target
If SendLink Then SendLinks "OpUser" & vbLf & User & vbLf & Channel.Name & vbLf & "" & vbLf & Target
End Sub

Public Sub DeOpUser(Channel As clsChannel, Target As String, User As String, Optional SendLink As Boolean = True)
On Error Resume Next
Dim i As Long, Chan As String
If Not Channel.IsOp(Target) Then Exit Sub
Chan = Channel.Name
For i = 1 To Channel.All.Count
     SendWsock NickToObject(Channel.All(i)).Index, (":" & User & " MODE " & Chan & " -o " & Target)
Next i
Channel.Ops.Remove Target
If Channel.IsVoice(Target) Then
Else
    Channel.NormUsers.Add Target, Target
End If
'1 = Command, 2 = Nick, 3 = Channel, 4 = Target
If SendLink Then SendLinks "DeOpUser" & vbLf & User & vbLf & Channel.Name & vbLf & "" & vbLf & Target
End Sub
Public Sub VoiceUser(Channel As clsChannel, Target As String, User As String, Optional VoiceAnyway As Boolean = False, Optional SendLink As Boolean = True)
On Error Resume Next
Dim i As Long, Chan As String
Chan = Channel.Name
If Channel.IsVoice(Target) And VoiceAnyway = False Then Exit Sub
For i = 1 To Channel.All.Count
     SendWsock NickToObject(Channel.All(i)).Index, (":" & User & " MODE " & Chan & " +v " & Target)
Next i
Channel.Voices.Add Target, Target
If Channel.IsNorm(Target) Then Channel.NormUsers.Remove Target
'1 = Command, 2 = Nick, 3 = Channel, 4 = Target
If SendLink Then SendLinks "VoiceUser" & vbLf & User & vbLf & Channel.Name & vbLf & "" & vbLf & Target
End Sub
Public Sub DeVoiceUser(Channel As clsChannel, Target As String, User As String, Optional SendLink As Boolean = True)
On Error Resume Next
Dim i As Long, Chan As String
If Not Channel.IsVoice(Target) Then Exit Sub
Chan = Channel.Name
For i = 1 To Channel.All.Count
     SendWsock NickToObject(Channel.All(i)).Index, (":" & User & " MODE " & Chan & " -v " & Target)
Next i
Channel.Voices.Remove Target
If Channel.IsOp(Target) Then
Else
    Channel.NormUsers.Add Target, Target
End If
'1 = Command, 2 = Nick, 3 = Channel, 4 = Target
If SendLink Then SendLinks "DeVoiceUser" & vbLf & User & vbLf & Channel.Name & vbLf & "" & vbLf & Target
End Sub

Public Sub BanUser(Channel As clsChannel, Target As String, User As String, Optional SendLink As Boolean = True)
Dim i As Long
If Not Channel.IsBanned2(Target) Then
    Channel.Bans.Add Target, Target
    For i = 1 To Channel.All.Count
        SendWsock NickToObject(Channel.All(i)).Index, ":" & User & "!" & NickToObject(User).ID & " MODE " & Channel.Name & " +b " & Target
    Next i
End If
'1 = Command, 2 = Nick, 3 = Channel, 4 = Target
If SendLink Then SendLinks "BanUser" & vbLf & User & vbLf & Channel.Name & vbLf & "" & vbLf & Target
End Sub

Public Sub UnBanUser(Channel As clsChannel, Target As String, User As String, Optional SendLink As Boolean = True)
Dim i As Long
For i = 1 To Channel.All.Count
    SendWsock NickToObject(Channel.All(i)).Index, ":" & User & "!" & NickToObject(User).ID & " MODE " & Channel.Name & " -b " & Target
Next i
Channel.Bans.Remove Target
'1 = Command, 2 = Nick, 3 = Channel, 4 = Target
If SendLink Then SendLinks "UnBanUser" & vbLf & User & vbLf & Channel.Name & vbLf & "" & vbLf & Target
End Sub

Public Sub RemoveChanModes(NewModes As String, Chan As String, User As clsUser, Optional SendLink As Boolean = True)
Dim Found As Boolean, X As Long, Channel As clsChannel, Modes As String
Set Channel = ChanToObject(Chan)
If InStr(1, NewModes, "r") And User.Nick <> "ChanServ" Then NewModes = Replace(NewModes, "r", "")
If Not (Mid(NewModes, 1, 1) = "k" Or Mid(NewModes, 1, 1) = "l") Then
    For X = 1 To Len(NewModes)
        If Channel.IsMode(Mid(NewModes, X, 1)) Then
            Channel.Modes.Remove Mid(NewModes, X, 1)
            Modes = Modes & Mid(NewModes, X, 1)
        End If
    Next X
    If Modes = "" Then Exit Sub
    Dim i As Long
    For i = 1 To Channel.All.Count
        SendWsock NickToObject(Channel.All(i)).Index, ":" & User.Nick & "!" & User.ID & " MODE " & Channel.Name & " -" & Modes
    Next i
ElseIf Mid(NewModes, 1, 2) = "lk" Then
    Dim Key As String, Limit As String
    Limit = Replace(NewModes, "lk ", "")
    If Channel.Key = Limit Then Channel.Key = ""
    Channel.Limit = 0
    For i = 1 To Channel.All.Count
        SendWsock NickToObject(Channel.All(i)).Index, ":" & User.Nick & "!" & User.ID & " MODE " & Channel.Name & " -" & NewModes
    Next i
ElseIf Mid(NewModes, 1, 1) = "k" Then
    If Channel.Key = Replace(NewModes, "k ", "") Then
        If SendLink Then SendLinks "UnKey" & vbLf & User.Nick & vbLf & Channel.Name & vbLf & Channel.Key
        Channel.Key = ""
        For i = 1 To Channel.All.Count
            SendWsock NickToObject(Channel.All(i)).Index, ":" & User.Nick & "!" & User.ID & " MODE " & Channel.Name & " -" & NewModes
        Next i
    End If
ElseIf Mid(NewModes, 1, 1) = "l" Then
    Channel.Limit = 0
    If SendLink Then SendLinks "UnLimit" & vbLf & User.Nick & vbLf & Channel.Name & vbLf & Channel.Limit
    For i = 1 To Channel.All.Count
        SendWsock NickToObject(Channel.All(i)).Index, ":" & User.Nick & "!" & User.ID & " MODE " & Channel.Name & " -" & NewModes
    Next i
End If
'1 = Command, 2 = Nick, 3 = +/-, 4 = Modes, 5 = Channel
If SendLink Then SendLinks "ChanMode" & vbLf & User.Nick & vbLf & "-" & vbLf & Modes & vbLf & Channel.Name
End Sub

Public Sub AddChanModes(NewModes As String, Chan As String, User As clsUser, Optional SendLink As Boolean = True)
Dim Found As Boolean, X As Long, Channel As clsChannel, Modes As String
On Error Resume Next
If InStr(1, NewModes, "r") And User.Nick <> "ChanServ" Then NewModes = Replace(NewModes, "r", "")
Set Channel = ChanToObject(Chan)
If Not (Mid(NewModes, 1, 1) = "k" Or Mid(NewModes, 1, 1) = "l") Then
    For X = 1 To Len(NewModes)
        If Not Channel.IsMode(Mid(NewModes, X, 1)) Then
            Channel.Modes.Add Mid(NewModes, X, 1), Mid(NewModes, X, 1)
            Modes = Modes & Mid(NewModes, X, 1)
        End If
    Next X
    If Modes = "" Then Exit Sub
    Dim i As Long
    For i = 1 To Channel.All.Count
        SendWsock NickToObject(Channel.All(i)).Index, ":" & User.Nick & "!" & User.ID & " MODE " & Channel.Name & " +" & Modes
    Next i
ElseIf Mid(NewModes, 1, 2) = "lk" Then
    Dim Key As String, Limit As String
    Limit = Replace(NewModes, "lk ", "")
    Key = Mid(Limit, 1, InStr(1, Limit, " ") - 1)
    Limit = Replace(Limit & " ", Key, "")
    Limit = Trim(Limit)
    Channel.Key = Limit
    Channel.Limit = Key
    For i = 1 To Channel.All.Count
        SendWsock NickToObject(Channel.All(i)).Index, ":" & User.Nick & "!" & User.ID & " MODE " & Channel.Name & " +" & NewModes
    Next i
ElseIf Mid(NewModes, 1, 1) = "k" Then
    Channel.Key = Replace(NewModes, "k ", "")
    If SendLink Then SendLinks "Key" & vbLf & User.Nick & vbLf & Channel.Name & vbLf & Channel.Key
    For i = 1 To Channel.All.Count
        SendWsock NickToObject(Channel.All(i)).Index, ":" & User.Nick & "!" & User.ID & " MODE " & Channel.Name & " +" & NewModes
    Next i
ElseIf Mid(NewModes, 1, 1) = "l" Then
    Channel.Limit = Replace(NewModes, "l ", "")
    If SendLink Then SendLinks "Limit" & vbLf & User.Nick & vbLf & Channel.Name & vbLf & Channel.Limit
    For i = 1 To Channel.All.Count
        SendWsock NickToObject(Channel.All(i)).Index, ":" & User.Nick & "!" & User.ID & " MODE " & Channel.Name & " +" & NewModes
    Next i
ElseIf Mid(NewModes, 1, 1) = "C" Then
    Channel.Limit = Replace(NewModes, "C ", "")
    For i = 1 To Channel.All.Count
        SendWsock NickToObject(Channel.All(i)).Index, ":" & User.Nick & "!" & User.ID & " MODE " & Channel.Name & " +" & NewModes
    Next i
ElseIf Mid(NewModes, 1, 1) = "c" Then
    Channel.Limit = Replace(NewModes, "c ", "")
    For i = 1 To Channel.All.Count
        SendWsock NickToObject(Channel.All(i)).Index, ":" & User.Nick & "!" & User.ID & " MODE " & Channel.Name & " +" & NewModes
    Next i
End If
'1 = Command, 2 = Nick, 3 = +/-, 4 = Modes, 5 = Channel
If SendLink Then SendLinks "ChanMode" & vbLf & User.Nick & vbLf & "+" & vbLf & Modes & vbLf & Channel.Name
End Sub

Public Function GetChanList(User As String)
Dim i As Long, Chan As clsChannel
For i = 1 To UBound(Channels)
    If Not Channels(i) Is Nothing Then
        GetChanList = GetChanList & ":" & ServerName & " 322 " & User & " " & Channels(i).Name & " " & Channels(i).All.Count & " :[+" & Channels(i).GetModes & "] " & Channels(i).Topic & vbCrLf
    End If
Next i
End Function

Public Sub InviteUser(Channel As clsChannel, Target As String, User As String, Optional SendLink As Boolean = True)
Dim i As Long
If Not Channel.IsInvited3(Replace(Target, "*!", "")) Then
    Channel.Invites.Add Replace(Target, "*!", ""), Replace(Target, "*!", "")
    For i = 1 To Channel.All.Count
        SendWsock NickToObject(Channel.All(i)).Index, ":" & User & "!" & NickToObject(User).ID & " MODE " & Channel.Name & " +I " & Target
    Next i
End If
'1 = Command, 2 = Nick, 3 = Channel, 4 = Target
If SendLink Then SendLinks "InviteUser" & vbLf & User & vbLf & Channel.Name & vbLf & "" & vbLf & Target
End Sub

Public Sub UnInviteUser(Channel As clsChannel, Target As String, User As String, Optional SendLink As Boolean = True)
Dim i As Long
If Channel.IsInvited3(Replace(Target, "*!", "")) Then
    Channel.Invites.Remove Replace(Target, "*!", "")
    For i = 1 To Channel.All.Count
        SendWsock NickToObject(Channel.All(i)).Index, ":" & User & "!" & NickToObject(User).ID & " MODE " & Channel.Name & " -I " & Target
    Next i
End If
'1 = Command, 2 = Nick, 3 = Channel, 4 = Target
If SendLink Then SendLinks "UnInviteUser" & vbLf & User & vbLf & Channel.Name & vbLf & "" & vbLf & Target
End Sub

Public Sub ExceptionUser(Channel As clsChannel, Target As String, User As String, Optional SendLink As Boolean = True)
Dim i As Long
If Not Channel.IsException2(Replace(Target, "*!", "")) Then
    Channel.Exceptions.Add Replace(Target, "*!", ""), Replace(Target, "*!", "")
    For i = 1 To Channel.All.Count
        SendWsock NickToObject(Channel.All(i)).Index, ":" & User & "!" & NickToObject(User).ID & " MODE " & Channel.Name & " +e " & Target
    Next i
End If
'1 = Command, 2 = Nick, 3 = Channel, 4 = Target
If SendLink Then SendLinks "ExceptUser" & vbLf & User & vbLf & Channel.Name & vbLf & "" & vbLf & Target
End Sub

Public Sub UnExceptionUser(Channel As clsChannel, Target As String, User As String, Optional SendLink As Boolean = True)
Dim i As Long
If Channel.IsException2(Replace(Target, "*!", "")) Then
    Channel.Exceptions.Remove Replace(Target, "*!", "")
    For i = 1 To Channel.All.Count
        SendWsock NickToObject(Channel.All(i)).Index, ":" & User & "!" & NickToObject(User).ID & " MODE " & Channel.Name & " -e " & Target
    Next i
End If
'1 = Command, 2 = Nick, 3 = Channel, 4 = Target
If SendLink Then SendLinks "UnExceptUser" & vbLf & User & vbLf & Channel.Name & vbLf & "" & vbLf & Target
End Sub

Public Sub AddUserMode(Index As Long, Modes As String, Optional Silent As Boolean = False, Optional SendLink As Boolean = True)
Dim NewModes As String
Modes = LCase(Modes)
Dim i As Long
For i = 1 To Len(Modes)
    Select Case Mid(Modes, i, 1)
        Case "s"
            If Not Users(Index).IsMode("s") Then
                NewModes = NewModes & "s"
                Users(Index).AddModes "s"
            End If
        Case "w"
            If Not Users(Index).IsMode("w") Then
                NewModes = NewModes & "w"
                Users(Index).AddModes "w"
            End If
    End Select
Next i
If Silent Then Exit Sub
If Not NewModes = "" Then
    SendWsock Index, ":" & Users(Index).Nick & " MODE " & Users(Index).Nick & " +" & NewModes
End If
'1 = Command, 2 = Nick, 3 = +/-, 4 = Modes
If SendLink Then SendLinks "ModeUser" & vbLf & Users(Index).Nick & vbLf & "+" & vbLf & Modes
End Sub

Public Sub RemoveUsermode(Index As Long, Modes As String, Optional Silent As Boolean = False, Optional SendLink As Boolean = True)
Dim NewModes As String
Modes = LCase(Modes)
Dim i As Long
For i = 1 To Len(Modes)
    Select Case Mid(Modes, i, 1)
        Case "s"
            If Users(Index).IsMode("s") Then
                NewModes = NewModes & "s"
                Users(Index).Modes.Remove "s"
            End If
        Case "w"
            If Users(Index).IsMode("w") Then
                NewModes = NewModes & "w"
                Users(Index).Modes.Remove "w"
            End If
        Case "a"
            If Users(Index).IsMode("a") Then
                NewModes = NewModes & "a"
                Users(Index).Modes.Remove "a"
                Users(Index).Away = False
                Users(Index).AwayMsg = ""
                SendWsock Index, ":" & ServerName & " 305 " & Users(Index).Nick & " :You are no longer marked as being away"
            End If
        Case "o"
            If Users(Index).IRCOp Then
                Users(Index).MsgsSent = 0
                SendSvrMsg Users(Index).Nick & " gave up his Operator status", True, ServerName
                NewModes = NewModes & "o"
                Users(Index).Modes.Remove "o"
                Users(Index).IRCOp = False
                Operators = Operators - 1
                Users(Index).DNS = Users(Index).RealDNS
                Users(Index).RealDNS = ""
                SendNotice "", "You are not an operator anymore", ServerName, , CInt(Index)
            End If
            Users(Index).IRCOp = False
    End Select
Next i
If Silent Then Exit Sub
If Not NewModes = "" Then
    SendWsock Index, ":" & Users(Index).Nick & " MODE " & Users(Index).Nick & " -" & NewModes
End If
'1 = Command, 2 = Nick, 3 = +/-, 4 = Modes
If SendLink Then SendLinks "ModeUser" & vbLf & Users(Index).Nick & vbLf & "-" & vbLf & Modes
End Sub
