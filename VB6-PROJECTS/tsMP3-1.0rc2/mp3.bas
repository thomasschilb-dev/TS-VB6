Attribute VB_Name = "Module1"

Option Explicit
Dim I1, i2, i3
Dim i33, i32
Dim PrevResizeX As Long
Dim PrevResizeY As Long
'' This fixes some bugs in MP3 Snatch an
'     d provides an method of "generating"
'' artist/title/album information based
'     solely on the filename (for those files
'' without ID3 tags.)
'' John Lambert
'' jrl7@po.cwru.edu
'' http://home.cwru.edu/~jrl7/
'' Version 1.0
' Original Title: MP3 Snatch
' Author: Leigh Bowers
' WWW: http://www.esheep.freeserve.co.uk
'     /compulsion/index.html
' Email: compulsion@esheep.freeserve.co.
'     uk
Private mvarFilename As String


Private Type Info
    sTitle As String
    sArtist As String
    sAlbum As String
    sComment As String
    sYear As String
    sGenre As String
    End Type
    Private MP3Info As Info
    Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
    Public Const LB_FINDSTRINGEXACT = &H1A2
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Public Const LB_SETHORIZONTALEXTENT = &H194
                            

Public Sub AddScroll(list As ListBox)
    Dim i As Integer, intGreatestLen As Integer, lngGreatestWidth As Long
    'Find Longest Text in Listbox


    For i = 0 To list.ListCount - 1


        If Len(list.list(i)) > Len(list.list(intGreatestLen)) Then
            intGreatestLen = i
        End If
    Next i
    'Get Twips
    lngGreatestWidth = list.Parent.TextWidth(list.list(intGreatestLen) + Space(1))
    'Space(1) is used to prevent the last Ch
    '     aracter from being cut off
    'Convert to Pixels
    lngGreatestWidth = lngGreatestWidth \ Screen.TwipsPerPixelX
    'Use api to add scrollbar
    SendMessage list.hwnd, LB_SETHORIZONTALEXTENT, lngGreatestWidth, 0
    
End Sub


Public Property Get filename() As String
    filename = mvarFilename
End Property


Public Function FileExists(sfile As String) As Boolean
    
    Dim i As Integer
    On Error Resume Next
    
    ' get the next available file number
    i = FreeFile
    Open sfile For Input As #i
    


    If Err Then
        ' don't really need to set this to false
        '
        ' because this would be the default. Jus
        '     t
        ' put it here for clarity
        FileExists = False
    Else
        FileExists = True
    End If
    
    Close #i
End Function


Private Function IsValidFile(ByVal sFilename) As Boolean
    Dim bOk As Boolean
    ' make sure file exists
    bOk = CBool(Dir(sFilename, vbHidden) <> "")
    
    Dim aExtensions, ext
    aExtensions = Array(".mp3", ".mp2", ".mp1")
    Dim bOkayExtension As Boolean
    bOkayExtension = False


    If bOk Then


        For Each ext In aExtensions


            If InStr(1, sFilename, ext, vbTextCompare) > 0 Then
                bOkayExtension = True
            End If
        Next 'ext
    End If
    
    IsValidFile = bOk And bOkayExtension
End Function


Public Property Let filename(ByVal sPassFilename As String)
    Dim iFreefile As Integer
    Dim lFilePos As Long
    Dim sData As String * 128
    
    Dim sGenre() As String
    ' Genre
    Const sGenreMatrix As String = "Blues|Classic Rock|Country|Dance|Disco|Funk|Grunge|" + _
    "Hip-Hop|Jazz|Metal|New Age|Oldies|Other|Pop|R&B|Rap|Reggae|Rock|Techno|" + _
    "Industrial|Alternative|Ska|Death Metal|Pranks|Soundtrack|Euro-Techno|" + _
    "Ambient|Trip Hop|Vocal|Jazz+Funk|Fusion|Trance|Classical|Instrumental|Acid|" + _
    "House|Game|Sound Clip|Gospel|Noise|Alt. Rock|Bass|Soul|Punk|Space|Meditative|" + _
    "Instrumental Pop|Instrumental Rock|Ethnic|Gothic|Darkwave|Techno-Industrial|Electronic|" + _
    "Pop-Folk|Eurodance|Dream|Southern Rock|Comedy|Cult|Gangsta Rap|Top 40|Christian Rap|" + _
    "Pop/Punk|Jungle|Native American|Cabaret|New Wave|Phychedelic|Rave|Showtunes|Trailer|" + _
    "Lo-Fi|Tribal|Acid Punk|Acid Jazz|Polka|Retro|Musical|Rock & Roll|Hard Rock|Folk|" + _
    "Folk/Rock|National Folk|Swing|Fast-Fusion|Bebob|Latin|Revival|Celtic|Blue Grass|" + _
    "Avantegarde|Gothic Rock|Progressive Rock|Psychedelic Rock|Symphonic Rock|Slow Rock|" + _
    "Big Band|Chorus|Easy Listening|Acoustic|Humour|Speech|Chanson|Opera|Chamber Music|" + _
    "Sonata|Symphony|Booty Bass|Primus|Porn Groove|Satire|Slow Jam|Club|Tango|Samba|Folklore|" + _
    "Ballad|power Ballad|Rhythmic Soul|Freestyle|Duet|Punk Rock|Drum Solo|A Capella|Euro-House|" + _
    "Dance Hall|Goa|Drum & Bass|Club-House|Hardcore|Terror|indie|Brit Pop|Negerpunk|Polsk Punk|" + _
    "Beat|Christian Gangsta Rap|Heavy Metal|Black Metal|Crossover|Comteporary Christian|" + _
    "Christian Rock|Merengue|Salsa|Trash Metal|Anime|JPop|Synth Pop"
    ' Build the Genre array (VB6+ only)
    sGenre = Split(sGenreMatrix, "|")
    ' Store the filename (for "Get Filename"
    '     property)
    mvarFilename = sPassFilename
    ' Clear the info variables
    


    If Not IsValidFile(sPassFilename) Then ' bug fix
        Exit Property
    End If
    
    MP3Info.sTitle = ""
    MP3Info.sArtist = ""
    MP3Info.sAlbum = ""
    MP3Info.sYear = ""
    MP3Info.sComment = ""
    ' Ensure the MP3 file exists
    ' Retrieve the info data from the MP3
    iFreefile = FreeFile
    lFilePos = FileLen(mvarFilename) - 127


    If lFilePos > 0 Then ' bug fix
        Open mvarFilename For Binary As #iFreefile
        Get #iFreefile, lFilePos, sData
        Close #iFreefile
    End If
    
    ' Populate the info variables


    If Left(sData, 3) = "TAG" Then
        MP3Info.sTitle = Mid(sData, 4, 30)
        MP3Info.sArtist = Mid(sData, 34, 30)
        MP3Info.sAlbum = Mid(sData, 64, 30)
        MP3Info.sYear = Mid(sData, 94, 4)
        MP3Info.sComment = Mid(sData, 98, 30)
        Dim lGenre
        lGenre = Asc(Mid(sData, 128, 1))


        If lGenre <= UBound(sGenre) Then
            MP3Info.sGenre = sGenre(lGenre)
        Else
            MP3Info.sGenre = ""
        End If
    Else
        MP3Info = GetInfo(mvarFilename)
    End If
End Property
'' Try to get something meaningful out o
'     f the filename


Private Function GetInfo(ByVal sFilename) As Info
    Dim i As Info
    GetInfo = i
    Dim S
    S = sFilename


    If InStrRev(S, "\") > 0 Then 'it's a full path
        S = Mid(S, InStrRev(S, "\") + 1)
    End If
    
    'drop extension
    S = Left(S, InStrRev(S, ".", , vbTextCompare) - 1)
    S = Replace(Trim(S), " ", " ")
    S = Trim(S)
    


    If CountItems(S, " ") < 1 Then
        i.sTitle = Replace(S, "_", " ")
        GetInfo = i
        Exit Function
    End If
    
    S = Trim(Replace(S, "_", " "))


    If Left(S, 1) = "(" And CountItems(S, "-") < 3 Then
        i.sArtist = Mid(S, 2, InStr(S, ")") - 2)
        S = Trim(Mid(S, InStr(S, ")") + 1))


        If Left(S, 1) = "-" Then 'grab title
            i.sTitle = Trim(Mid(S, 2))
        Else 'grab title anyway


            If InStr(S, "-") > 0 Then
                i.sAlbum = Mid(S, InStr(S, "-") + 1)
                i.sTitle = Left(S, InStr(S, "-") - 1)
            Else
                i.sTitle = Trim(S)
            End If
        End If
    Else
        Dim aThings
        Dim l
        aThings = Split(S, "- ")


        For l = 0 To UBound(aThings)


            If Not IsNumeric(aThings(l)) Then


                If i.sArtist = "" Then
                    i.sArtist = aThings(l)
                Else


                    If IsNumeric(aThings(l - 1)) Then ' title


                        If i.sTitle = "" Then
                            i.sTitle = aThings(l)
                        End If
                    ElseIf i.sAlbum = "" Then
                        i.sAlbum = aThings(l)
                    End If
                End If
            End If
        Next ' i
        
    End If
    
    i.sArtist = Replace(Replace(i.sArtist, "(", ""), ")", "")
    


    If Left(S, 1) <> "(" And i.sTitle = "" And (InStr(sFilename, "\") <> InStrRev(sFilename, "\")) Then
        ' recurse
        GetInfo = GetInfo(FixDir(sFilename))
    Else
        GetInfo = i
    End If
End Function


Private Function CountItems(S, sToCount)
    Dim a
    a = Split(S, sToCount)


    If UBound(a) = -1 Then
        CountItems = 0
    Else
        CountItems = UBound(a) - LBound(a)
    End If
End Function
Public Function RandomNum(Min, Max) As Long
    RandomNum = Int((Max - Min + 9500) * Rnd + Min)
End Function



Private Function FixDir(sFullpath)
    Dim s1, s2
    s1 = Trim(Left(sFullpath, InStrRev(sFullpath, "\") - 1))
    s2 = Trim(Mid(sFullpath, InStrRev(sFullpath, "\") + 1))
    FixDir = s1 & " - " & s2
End Function
Public Sub RemoveDupes(list1 As ListBox)
Dim i
   On Error GoTo lol
For i = 0 To list1.ListCount - 1
DoEvents
list1.ListIndex = i
Dim x As String
Dim xa As String
Dim xaa As String
Dim xx As String
x = list1.list(i)
xx = list1.list(i + 1)
'trims all spaces in a the current item
xa = trimtext(x)
xaa = trimtext(xx)
'if dupe is found removes it
If LCase(xa) = LCase(xaa) Then
DoEvents
list1.RemoveItem i
i = i - 1
End If
Next i
MsgBox "Done!"
Exit Sub
lol:
MsgBox "Done!"
Exit Sub

End Sub
Function trimtext(txt As String) As String
'starts from the beginging to the end of the text
Dim i
For i = 1 To Len(txt)
Dim xx
DoEvents
Dim x
'checks the letters one by one for a space
x = Mid(txt, i)
x = Left(x, 1)
If x = " " Then
Else
xx = xx + x
End If
Next i
trimtext = xx
End Function




Public Property Get Title() As String
    Title = Trim(MP3Info.sTitle)
End Property


Public Property Get Artist() As String
    Artist = Trim(MP3Info.sArtist)
End Property


Public Property Get Genre() As String
    Genre = Trim(MP3Info.sGenre)
End Property


Public Property Get Album() As String
    Album = Trim(MP3Info.sAlbum)
End Property
Public Function ResizeAll(FormName As Form)
    Dim tmpControl As Control
    On Error Resume Next
    'Ignores errors in case the control does
    '     n't
    'have a width, height, etc.


    If PrevResizeX = 0 Then
        'If the previous form width was 0
        'Which means that this function wasn't r
        '     un before
        'then change prevresizex and y and exit


'     function
    PrevResizeX = FormName.ScaleWidth
    PrevResizeY = FormName.ScaleHeight
    Exit Function
End If


For Each tmpControl In FormName
    'A loop to make tmpControl equal to ever
    '     y
    'control on the form


    If TypeOf tmpControl Is Line Then
        'Checks the type of control, if its a
        'Line, change its X1, X2, Y1, Y2 values
        tmpControl.x1 = tmpControl.x1 / PrevResizeX * FormName.ScaleWidth
        tmpControl.x2 = tmpControl.x2 / PrevResizeX * FormName.ScaleWidth
        tmpControl.Y1 = tmpControl.Y1 / PrevResizeY * FormName.ScaleHeight
        tmpControl.Y2 = tmpControl.Y2 / PrevResizeY * FormName.ScaleHeight
        'These four lines see the previous ratio
        '
        'Of the control to the form, and change
        '     they're
        'current ratios to the same thing
    Else
        'Changes everything elses left, top
        'Width, and height
        tmpControl.Left = tmpControl.Left / PrevResizeX * FormName.ScaleWidth
        tmpControl.Top = tmpControl.Top / PrevResizeY * FormName.ScaleHeight
        tmpControl.Width = tmpControl.Width / PrevResizeX * FormName.ScaleWidth
        tmpControl.Height = tmpControl.Height / PrevResizeY * FormName.ScaleHeight
        'These four lines see the previous ratio
        '
        'Of the control to the form, and change
        '     they're
        'current ratios to the same thing
    End If
Next tmpControl
PrevResizeX = FormName.ScaleWidth
PrevResizeY = FormName.ScaleHeight
'Changes prevresize x and y to current w
'     idth
'and height
End Function




Function FindPartialInCombo(Ctl As Control, S As String)
    Dim i As Long, j As Long, k As Long, lun As Long
    FindPartialInCombo = -1
    i = 0: j = Ctl.ListCount - 1
    lun = Len(S)


    Do
        'non trovato, esce
        If i > j Then Exit Function
        k = (i + j) / 2


        Select Case StrComp(Left(Ctl.list(k), lun), S)
            Case 0: Exit Do
            Case -1: i = k + 1 ' If < look In the second half
            Case 1: j = k - 1 ' If > look In the first half
        End Select
Loop
'sequential search backwards to found th
'     e first matching element


Do While k > 0
    If StrComp(Left(Ctl.list(k - 1), lun), S) <> 0 Then Exit Do
    k = k - 1
Loop
FindPartialInCombo = k
Ctl = k
End Function
Public Sub listsearch()


If Form1.txt1.text = "" Then
Form1.list1.ListIndex = -1
Exit Sub
End If

For I1 = 0 To Form1.list1.ListCount - 1
i2 = Form1.list1.list(I1)

If InStr(1, i2, Form1.txt1.text, 1) = 1 Then
i32 = UCase(i2)

i2 = UCase$(i2)

i3 = Len(Form1.txt1.text)
'Form1.txt1.Text = i3

i33 = Mid(i32, 1, i3)

If UCase(i33) = UCase(Form1.txt1.text) Then

Form1.list1.text = Form1.txt1.text & Mid(i2, i3 + 1)

Exit Sub

Else

End If
End If

Next
End Sub



Function Search_ListBox(trig$, lst As ListBox) As Long
    'This function will search a listbox for
    '     a specified
    'string. It returns the index value of t
    '     he item if the
    'sting is found. If the string is not fo
    '     und, then -1 is
    'returned.
    Dim items As Long
    Dim n As Long
    items = lst.ListCount - 1


    For n = 0 To items Step 1


        If lst.list(n) = trig$ Then
            Search_ListBox = n
            Exit Function
        End If
    Next n
    Search_ListBox = -1
End Function
Public Sub SaveListBox(TheList As ListBox, Directory As String)
  'Example: Call SaveListBox(list1, "C:\Te
'     mp\MyList.dat")
    Dim savelist As Long
    On Error Resume Next
    Open Directory$ For Output As #1


    For savelist& = 0 To TheList.ListCount - 1
        Print #1, TheList.list(savelist&)
    Next savelist&
    Close #1
End Sub


Function LBDupe(lpBox As ListBox) As Integer
    Dim nCount As Integer, nPos1 As Integer, nPos2 As Integer, nDelete As Integer
    Dim sText As String


    If lpBox.ListCount < 3 Then
        LBDupe = 0
        Exit Function
    End If


    For nCount = 0 To lpBox.ListCount - 1


        Do


            DoEvents '2
                sText = lpBox.list(nCount) 'had To update this line, sorry
                nPos1 = SendMessageByString(lpBox.hwnd, LB_FINDSTRINGEXACT, nCount, sText)
                nPos2 = SendMessageByString(lpBox.hwnd, LB_FINDSTRINGEXACT, nPos1 + 1, sText)
                If nPos2 = -1 Or nPos2 = nPos1 Then Exit Do
                lpBox.RemoveItem nPos2
                nDelete = nDelete + 1
            Loop
        Next nCount
        LBDupe = nDelete
    End Function

Public Sub LoadListBox(TheList As ListBox, Directory As String)
   'Example: Call LoadListBox(list1, "C:\Te
'     mp\MyList.dat")
    Dim MyString As String
    On Error Resume Next
    Open Directory$ For Input As #1


    While Not EOF(1)
        Input #1, MyString$


        DoEvents
            TheList.AddItem MyString$
        Wend
        Close #1
        
    End Sub


Public Property Get Year() As String
    Year = Trim(MP3Info.sYear)
End Property


Public Property Get Comment() As String
    Comment = Trim(MP3Info.sComment)
End Property

            

