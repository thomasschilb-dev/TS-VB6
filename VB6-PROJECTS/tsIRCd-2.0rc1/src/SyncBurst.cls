VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SyncBurst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim FileOpen As Boolean

Public Sub Send_File(FileToSend As String, CTLWinsock As Winsock)
Dim temp As String
Dim BlockSize As Long
Open FileToSend For Binary Access Read As #1
BlockSize = 2048
Do While Not EOF(1)
    temp = Space$(BlockSize)
    Get 1, , temp
    CTLWinsock.SendData temp
    DoEvents
Loop
CTLWinsock.SendData "xx"
Close #1
End Sub

Public Sub Receive_File(CTLWinsock As Winsock, strFile As String)
Dim temp As String
Dim data As String
CTLWinsock.GetData data
If FileOpen = False Then
    Open strFile For Binary Access Write As #2
    FileOpen = True
ElseIf FileOpen = True Then
    DoEvents
    End If
    temp = data
    If temp = "xx" Then
        Close #2
        FileOpen = False
    Else
        Put 2, , temp
    End If
End Sub
