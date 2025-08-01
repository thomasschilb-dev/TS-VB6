VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMP3time"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Reads internal MP3 file data."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'
Const iSecsInDay As Long = 86400


Enum iConstants
    iSeconds = 0
    iMinutes = 1
    iHours = 2
    iMilliSec = 3
End Enum

'local variable(s) to hold property value(s)

Private mvarFilename As String 'local copy
Private mvarBitRate As Long 'local copy
Private mvarCRCProtected As Boolean 'local copy
Private mvarFrames As Long 'local copy
Private mvarFrameSize As Long 'local copy
Private mvarTime As Long 'local copy
Private mvarFrequency As Long 'local copy
Private mvarFileSize As Long 'local copy
Private mvarVersion As Byte 'local copy
Private mvarVersionText As String 'local copy
Private mvarLayer As Byte 'local copy
Private mvarLayerText As String 'local copy
Private mvarPadding As Boolean 'local copy
Private mvarPrivateBit As Boolean 'local copy
Private mvarChannelMode As Byte 'local copy
Private mvarChannelModeText As String 'local copy
Private mvarChannelModeExtention As Byte 'local copy
Private mvarCopyright As Boolean 'local copy
Private mvarOriginal As Boolean 'local copy
Private mvarEmphasis As Byte 'local copy
Private mvarEmphasisText As String 'local copy
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent ReadComplete[(arg1, arg2, ... , argn)]
Public Event ReadComplete(ByVal success As Boolean)
Attribute ReadComplete.VB_Description = "Called after a file read"

Private bitrate_lookup(7, 15) As Integer
Private freq_lookup(3, 7) As Long

Dim FrameRates(3)
Private Sub ZeroVars()
 mvarBitRate = 0
 mvarCRCProtected = 0
 mvarFrames = 0
 mvarFrameSize = 0
 mvarTime = 0
 mvarFrequency = 0
 mvarFileSize = 0
 mvarVersion = 0
 mvarLayer = 0
 mvarPadding = 0
 mvarPrivateBit = 0
 mvarChannelMode = 0
 mvarChannelModeExtention = 0
 mvarCopyright = 0
 mvarOriginal = 0
 mvarEmphasis = 0
 mvarVersionText = ""
 mvarChannelModeText = ""
 mvarLayerText = ""
 mvarEmphasisText = ""

End Sub
Public Function refresh() As Boolean
Attribute refresh.VB_Description = "refreshes the data from the MPEG file.  Filename must be valid."
ZeroVars
Dim State As Boolean
If Not FileExists(mvarFilename) Then
    RaiseEvent ReadComplete(False)
Else
    GetMP3data
    State = True
    'Take care of crisis stuff
    If mvarFrequency = 99999 Or mvarBitRate = 0 Then State = False
    If mvarFrequency = 0 Or mvarFileSize = 0 Then State = False
    If mvarTime = 0 Or mvarFrameSize = 0 Then State = False
    'Exit with a valid state
    RaiseEvent ReadComplete(State)
    refresh = State
End If

End Function

Public Property Get EmphasisText() As String
Attribute EmphasisText.VB_Description = "Emphasis mode descriptor text"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EmphasisText
    EmphasisText = mvarEmphasisText
End Property

Public Property Get Emphasis() As Byte
Attribute Emphasis.VB_Description = "Emphasis mode"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Emphasis
    Emphasis = mvarEmphasis
End Property

Public Property Get Original() As Boolean
Attribute Original.VB_Description = "True if file is a copy of an original"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Original
    Original = mvarOriginal
End Property

Public Property Get CRCProtected() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CRCProtected
    CRCProtected = mvarCRCProtected
End Property

Public Property Get Copyright() As Boolean
Attribute Copyright.VB_Description = "True if file is copyrightes"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Copyright
    Copyright = mvarCopyright
End Property

Public Property Get ChannelModeExtention() As Byte
Attribute ChannelModeExtention.VB_Description = "Contains the channel  mode extention if channel mode is 1 (Joint Stereo)"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ChannelModeExtention
    ChannelModeExtention = mvarChannelModeExtention
End Property

Public Property Get ChannelModeText() As String
Attribute ChannelModeText.VB_Description = "Contains the channel (stereo/mono) mode descriptor string"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ChannelModeText
    ChannelModeText = mvarChannelModeText
End Property

Public Property Get ChannelMode() As Byte
Attribute ChannelMode.VB_Description = "Contains the channel (stereo/mono) mode"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ChannelMode
    ChannelMode = mvarChannelMode
End Property
Public Function pPause(ByVal Number As Single, _
    Optional ByVal Unit As iConstants)
    Dim iStopTime, fakeTimer, sAfterMidnight, sBeforeMidnight


    If Unit = iSeconds Then
        Number = Number
    ElseIf Unit = iMinutes Then
        Number = Number * 60
    ElseIf Unit = iHours Then
        Number = Number * 3600
    ElseIf Unit = iMilliSec Then
        Number = Number / 1000
    End If
    fakeTimer = Timer
    iStopTime = fakeTimer + Number


    If iStopTime > iSecsInDay Then
        sAfterMidnight = iStopTime - iSecsInDay
        sBeforeMidnight = Number - sAfterMidnight
        fakeTimer = Timer


        While Timer < fakeTimer + sBeforeMidnight And Timer <> 0


            DoEvents
            Wend
            fakeTimer = Timer


            While Timer < fakeTimer + sAfterMidnight


                DoEvents
                Wend
            Else 'if pausing won't continue through midnight


                While Timer < iStopTime


                    DoEvents
                    Wend
                End If
            End Function


Public Property Get PrivateBit() As Boolean
Attribute PrivateBit.VB_Description = "Holds the PrivateBit data from the MPEG file"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrivateBit
    PrivateBit = mvarPrivateBit
End Property

Public Property Get Padding() As Boolean
Attribute Padding.VB_Description = "True if frames are padded"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Padding
    Padding = mvarPadding
End Property

Public Property Get LayerText() As String
Attribute LayerText.VB_Description = "Contains the MPEG layer mode descriptor text"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LayerText
    LayerText = mvarLayerText
End Property

Public Property Get Layer() As Byte
Attribute Layer.VB_Description = "Contains the MPEG layer mode"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Layer
    Layer = mvarLayer
End Property

Public Property Get VersionText() As String
Attribute VersionText.VB_Description = "Contains the MPEG version descriptor text"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VersionText
    VersionText = mvarVersionText
End Property

Public Property Get Version() As Byte
Attribute Version.VB_Description = "Contains the MPEG version"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Version
    Version = mvarVersion
End Property

Public Property Get Frequency() As Long
Attribute Frequency.VB_Description = "Contains the sampling frequency"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Frequency
    Frequency = mvarFrequency
End Property

Public Property Get BitRate() As Long
Attribute BitRate.VB_Description = "Contains the MPEG bitrate"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BitRate
    BitRate = mvarBitRate
End Property
Public Property Get Frames() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Frames
    Frames = mvarFrames
End Property

Public Property Get FrameSize() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Frames
    FrameSize = mvarFrameSize
End Property

Public Function ConvertSeconds(seconds) As Date
Dim tm As Date, t1 As Double
Const OneSecond = 1.15740740740741E-05

t1 = seconds * OneSecond
ConvertSeconds = t1

End Function

Public Property Get Time() As Date
Dim tm As Date, t1 As Double
Const OneSecond = 1.15740740740741E-05

t1 = mvarTime * OneSecond
Time = t1

End Property

Public Property Get seconds() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Time
    seconds = mvarTime
End Property

Public Property Get FileSize() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileSize
    FileSize = mvarFileSize
End Property


Public Property Let FileName(ByVal vData As String)
ZeroVars
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Filename = 5
    mvarFilename = vData
End Property


Public Property Get FileName() As String
Attribute FileName.VB_Description = "Sets/Returns the MP3 file to read"
Attribute FileName.VB_UserMemId = 0
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Filename
    FileName = mvarFilename
End Property



Private Sub Class_Initialize()
  Dim bitrate_core$(), bitrate_data$, BitRate%, Ver_Layer%, freq_data$, freq%
  Dim freq_core$()
  ' setup array for mpeg bitrate info
  bitrate_data = "" '
  bitrate_data = bitrate_data + "999,999,999,999,999,999,"
  bitrate_data = bitrate_data + "032,032,032,032,008,008,"
  bitrate_data = bitrate_data + "064,048,040,048,016,016,"
  bitrate_data = bitrate_data + "096,056,048,056,024,024,"
  bitrate_data = bitrate_data + "128,064,056,064,032,032,"
  bitrate_data = bitrate_data + "160,080,064,080,040,040,"
  bitrate_data = bitrate_data + "192,096,080,096,048,048,"
  bitrate_data = bitrate_data + "224,112,096,112,056,056,"
  bitrate_data = bitrate_data + "256,128,112,128,064,064,"
  bitrate_data = bitrate_data + "288,160,128,144,080,080,"
  bitrate_data = bitrate_data + "320,192,160,160,096,096,"
  bitrate_data = bitrate_data + "352,224,192,176,112,112,"
  bitrate_data = bitrate_data + "384,256,224,192,128,128,"
  bitrate_data = bitrate_data + "416,320,256,224,144,144,"
  bitrate_data = bitrate_data + "448,384,320,256,160,160,"
  bitrate_data = bitrate_data + "999,999,999,999,999,999,"
  bitrate_core = Split(bitrate_data, ",")
  
  For BitRate = 1 To 14
    For Ver_Layer = 0 To 2
      bitrate_lookup(7 - Ver_Layer, BitRate) = Val(bitrate_core((BitRate * 6) + Ver_Layer))
    Next
    For Ver_Layer = 0 To 2
      bitrate_lookup(3 - Ver_Layer, BitRate) = Val(bitrate_core((BitRate * 6) + 3 + Ver_Layer))
    Next
  Next

  ' setup array for mpeg freq info
  freq_data = "" '
  freq_data = freq_data + "44100,22050,11025,"
  freq_data = freq_data + "48000,24000,12000,"
  freq_data = freq_data + "32000,16000,08000,"
  freq_data = freq_data + "99999,99999,99999,"
  freq_core = Split(freq_data, ",")
  For freq = 0 To 3
    freq_lookup(3, freq) = Val(freq_core((freq * 3)))
    freq_lookup(2, freq) = Val(freq_core((freq * 3) + 1))
    freq_lookup(0, freq) = Val(freq_core((freq * 3) + 2))
    
  Next

FrameRates(0) = 38.5
FrameRates(1) = 32.5
FrameRates(2) = 27.8
FrameRates(3) = 0
End Sub


   Private Function shift_those_bits(dIN As String) As String
     ' need to left shift 4 bits losing most significant 4 bits
     Dim sd1, sd2, sd3, do1, do2, duff2 As Integer
     Dim duff As String
     duff = Left(dIN, 1)
     duff2 = Asc(duff)
     sd1 = Asc(Left(dIN, 1))
     sd2 = Asc(Mid(dIN, 2, 1))
     sd3 = Asc(Right(dIN, 1))
     
     do1 = ((sd1 And &HF) * 16) Or ((sd2 And &HF0) / 16)
     do2 = ((sd2 And &HF) * 16) Or ((sd3 And &HF0) / 16)
     shift_those_bits = Chr(do1) + Chr(do2)
   End Function

Private Function FileExists(FileName$) As Boolean
Dim Temp%
On Error Resume Next
Err = 0
Temp% = FreeFile
Open FileName For Input As Temp%
Close Temp%
FileExists = (Err = 0)
On Error GoTo 0
'Debug.Print Filename
End Function
 Private Sub GetMP3data()
 On Error Resume Next
     Dim dIN As String, Afile As Long
     Dim CR As String, i&, I1%
     Dim d1%, d2%, temp_String$, mp3bits_string
     Dim dd1%, dd2%
     Dim dSHIFT, mp3_ID1, mp3_prot, mp3_freq, Type_Layer, mp3_bitrate, FrameSize
     Dim Bytes(3) As Byte
     
     CR = Chr(10)
     Afile = FreeFile
     Open mvarFilename For Binary As #Afile
     ' read in 1st 4k of .mp3 file to find a frame header
     dIN = Input(4096, #Afile)
     mvarFileSize = LOF(Afile) ' needed to calculate track duration
     Close #Afile
     
     ' frame header starts with 12 set bits [sync]
     ' NB this ignores MPEG-2.5 which is 11 set bits, 1 zero bit.
     
     ' my search for the sync bits only works on nibble boundaries,
     ' I'm not sure if it is necessary to search on bit boundaries -
     ' if so then this search will be 4* slower and require a rewrite
     ' of this search section and shift_those_bits.
     i = 0
     Do Until i = 4095
ReEnter:       i = i + 1
       d1 = Asc(Mid(dIN, i, 1))
       d2 = Asc(Mid(dIN, i + 1, 1))
       If d1 = &HFF And (d2 And &HE0) = &HE0 Then
         'Debug.Print "Found at"; i
         ' get 20 hdr bits - they are last 20 bits of next 3 bytes
         temp_String = Mid(dIN, i + 1, 3)
         mp3bits_string = temp_String 'shift_those_bits(Mid(dIN, i + 1, 3))
         Exit Do
       End If
       ' if we haven't found the sync yet then shift left by 4 bits
       dSHIFT = shift_those_bits(Mid(dIN, i, 3))
       dd1 = Asc(Left(dSHIFT, 1))
       dd2 = Asc(Right(dSHIFT, 1))
       If dd1 = &HFF And (dd2 And &HE0) = &HE0 Then
         'Debug.Print "Found at"; i; "& a nibble"
         ' get 20 hdr bits - they are first 20 bits of next 3 bytes
         mp3bits_string = Mid(dIN, i + 2, 3)
         Exit Do
       End If
     Loop
     If i = 4096 Then Exit Sub 'not found!
     For I1 = 1 To 3
        Bytes(I1) = Asc(Mid(mp3bits_string, I1))
    Next
     ' 1st 20 bits of mp3bits_string are hdr info for this frame
     ' 1st bit is ID - 0=MPG-2, 1=MPG-1
     mvarVersion = (&H18 And Bytes(1)) / 8
     mp3_ID1 = (mvarVersion And 1)
          ' next 2 bits are Layer
     mvarLayer = (&H6 And Bytes(1)) / 2
     ' next bit is Protection
     mp3_prot = &H1 And Bytes(1)
     mvarCRCProtected = mp3_prot <> 0
     ' next 4 bits are bitrate
     mp3_bitrate = (&HF0 And Bytes(2)) / 16
     Type_Layer = (mp3_ID1 * 4) Or mvarLayer
     mvarBitRate = 1000 * CLng((bitrate_lookup(Type_Layer, mp3_bitrate)))
     'next 2 bits are frequency
     mp3_freq = (&HC And Bytes(2)) / 4
     mvarFrequency = freq_lookup(mvarVersion, mp3_freq)
     If mvarFrequency = 99999 Or mvarBitRate = 0 Then
        i = i + 4
        GoTo ReEnter
    End If
     ' next bit is Padding
     mvarPadding = ((&H2 And Bytes(2)) / 2) = 1
     ' next bit is private bit
     mvarPrivateBit = ((&H10 And Bytes(3)) / 2) = 1
     'next 2 bits are Channel mode
     mvarChannelMode = (&HC0 And Bytes(3)) / 64
     'next 2 bits are Channel mode extention
     mvarChannelModeExtention = (&H30 And Bytes(3)) / 16
     'next bit is copyright flag
     mvarCopyright = ((&H8 And Bytes(3)) / 8) = 1
     'next bit is original flag
     mvarOriginal = ((&H4 And Bytes(3)) / 4) = 1
     'next bit is original flag
     mvarEmphasis = &H3 And Bytes(3)
     
     Select Case mvarVersion
     Case 0
       mvarVersionText = "MPEG-2.5"
     Case 1
     Case 2
       mvarVersionText = "MPEG-2.0"
     Case 3
       mvarVersionText = "MPEG-1.0"
     End Select
     
      Select Case mvarLayer
        Case 1
          mvarLayerText = "Layer III"
            FrameSize = (144 * (mvarBitRate / mvarFrequency)) ' + Abs(mvarPadding)
        Case 2
          mvarLayerText = "Layer II"
            FrameSize = (144 * (mvarBitRate / mvarFrequency)) ' + Abs(mvarPadding)
        Case 3
          mvarLayerText = "Layer I"
             FrameSize = ((12 * (mvarBitRate / mvarFrequency) + Abs(mvarPadding))) * 4
     End Select
    
    Select Case mvarChannelMode
        Case 0
          mvarChannelModeText = "Stereo"
        Case 1
          mvarChannelModeText = "Joint Stereo (Stereo)"
          'compensate for interleave factor
          If mvarVersion < 3 Then FrameSize = Fix(FrameSize) / 2
          'Deal with Mpeg 2.5
          If mvarVersion = 0 Then FrameSize = Fix(FrameSize) / 2
        Case 2
          mvarChannelModeText = "Dual Channel (Stereo)"
        Case 3
          mvarChannelModeText = "Single Channel (Mono)"
          '2X data since only 1 channel
          If mvarVersion < 3 Then FrameSize = Fix(FrameSize) / 2
    End Select
      
    
    Select Case mvarEmphasis
        Case 0
          mvarEmphasisText = "None"
        Case 1
          mvarEmphasisText = "50/15 ms"
        Case 2
          mvarEmphasisText = "reserved"
        Case 3
          mvarEmphasisText = "CIT J.17"
    End Select
      ' calculate track time
      mvarFrameSize = Fix(FrameSize)
      mvarFrames = mvarFileSize / Fix(FrameSize)
      mvarTime = mvarFrames / FrameRates(mp3_freq) '38.5 frames per sec.
   End Sub

