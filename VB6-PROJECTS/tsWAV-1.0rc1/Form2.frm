VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Waveform picture"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub DrawWaves()
' Graph the waveform
   Dim x As Long               ' current X position
   Dim leftYOffset As Long     ' Y offset for left channel graph
   Dim rightYOffset As Long    ' Y offset for right channel graph
   Dim curLeftY As Long        ' current left channel Y value
   Dim curRightY As Long       ' current right channel Y value
   Dim lastX As Long           ' last X position
   Dim lastLeftY As Long       ' last left channel Y value
   Dim lastRightY As Long      ' last right channel Y value
   Dim maxAmplitude As Long    ' the maximum amplitude for a wavegraph on the form
   Dim leftVol As Double       ' buffer for retrieving the left volume level
   Dim rightVol As Double      ' buffer for retrieving the right volume level
   Dim scaleFactor As Double   ' samples per pixel on the wave graph
   Dim xStep As Double         ' pixels per sample on the wave graph
   Dim curSample As Long       ' current sample number
   
   ' clear the screen
   Me.Cls
   
   ' if no file is loaded, don't try to draw graph
   If (Module1.fFileLoaded = False) Then
       Exit Sub
   End If
   
   ' calculate drawing parameters
   scaleFactor = (Module1.drawTo - Module1.drawFrom) / Me.Width
   If (scaleFactor < 1) Then
       xStep = 1 / scaleFactor
   Else
       xStep = 1
   End If
   
   ' Draw the graph
   If (Module1.format.nChannels = 2) Then
      maxAmplitude = Me.Height / 4
      leftYOffset = maxAmplitude
      rightYOffset = maxAmplitude * 3
       
      For x = 0 To Me.Width Step xStep
         curSample = scaleFactor * x + Module1.drawFrom
         If (Module1.format.wBitsPerSample = 16) Then
             GetStereo16Sample curSample, leftVol, rightVol
         Else
             GetStereo8Sample curSample, leftVol, rightVol
         End If
         curRightY = CLng(rightVol * maxAmplitude)
         curLeftY = CLng(leftVol * maxAmplitude)
         Line (lastX, leftYOffset + lastLeftY)-(x, curLeftY + leftYOffset)
         Line (lastX, rightYOffset + lastRightY)-(x, curRightY + rightYOffset)
         lastLeftY = curLeftY
         lastRightY = curRightY
         lastX = x
      Next
   Else
      maxAmplitude = Me.Height / 2
      leftYOffset = maxAmplitude
      
      For x = 0 To Me.Width Step xStep
         curSample = scaleFactor * x + Module1.drawFrom
         If (Module1.format.wBitsPerSample = 16) Then
             GetMono16Sample curSample, leftVol
         Else
             GetMono8Sample curSample, leftVol
         End If
         curLeftY = CLng(leftVol * maxAmplitude)
         Line (lastX, leftYOffset + lastLeftY)-(x, curLeftY + leftYOffset)
         lastLeftY = curLeftY
         lastX = x
      Next
   End If

End Sub

Private Sub Form_Paint()
   DrawWaves
End Sub

Private Sub Form_Resize()
   DrawWaves
End Sub
