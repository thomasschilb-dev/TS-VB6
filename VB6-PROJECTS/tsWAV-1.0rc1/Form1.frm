VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Form"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3360
      Top             =   600
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   1560
      Width           =   3615
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   1200
      Width           =   3615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3960
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   120
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "file"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stop"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Play"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "end pos"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "start pos"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim filename As String          ' wave file to play
Dim errStr As String * 200      ' buffer for retrieving error messages
Dim waveForm As Form2           ' form to draw wavegraph on
Const MAX_SCROLL_VALUE = 1000   ' range for scroll controls

Private Sub Form_Load()
' Initialize form, create wavegraph form
   Module1.fFileLoaded = False
   CommonDialog1.filename = "*.wav"
   CommonDialog1.DefaultExt = "wav"
   Set waveForm = New Form2
   waveForm.Move Me.Left, Me.Top + Me.Height, Me.Width, Me.Height * 1.5
   waveForm.Show 0, Me
   HScroll1.Max = MAX_SCROLL_VALUE
   HScroll2.Max = MAX_SCROLL_VALUE
   HScroll2.Value = MAX_SCROLL_VALUE
End Sub

Private Sub Command1_Click()
'Open a wavefile and initialize the form
   CommonDialog1.ShowOpen
   filename = CommonDialog1.filename
   Text1.Text = filename
   LoadFile filename
   Module1.drawFrom = 0
   Module1.drawTo = Module1.numSamples
   HScroll1.Value = 0
   HScroll2.Value = MAX_SCROLL_VALUE
   waveForm.DrawWaves
End Sub

Private Sub Command2_Click()
' Start playing the wavefile
   If (Module1.fPlaying = False) Then
      ' -1 specifies the wave mapper
      Play -1
   End If
End Sub

Private Sub Command3_Click()
' Stop playing the wavefile
   StopPlay
End Sub

Private Sub HScroll1_Change()
' Set beginning position in wave file
   If HScroll1.Value >= HScroll2.Value Then
      HScroll1.Value = HScroll2.Value - 1
      End If
   SetPlayRange
End Sub

Private Sub HScroll2_Change()
' Set end position in wave file
   If HScroll2.Value <= HScroll1.Value Then
      HScroll2.Value = HScroll1.Value + 1
      End If
   SetPlayRange
End Sub

Private Sub SetPlayRange()
' Set the range to be played and redraw the wave graph
   Module1.drawFrom = CLng(Module1.numSamples * (HScroll1.Value / MAX_SCROLL_VALUE))
   Module1.drawTo = CLng(Module1.numSamples * (HScroll2.Value / MAX_SCROLL_VALUE))
   waveForm.DrawWaves
End Sub

Private Sub Timer1_Timer()
   If Module1.fPlaying = False Then
      Module1.CloseWaveOut
      Timer1.Enabled = False
   End If
End Sub
