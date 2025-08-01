VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Download files with Api"
   ClientHeight    =   750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   2895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMethod2 
      Caption         =   "Method #2 - without Dialog"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdMethod1 
      Caption         =   "Method #1 - with Dialog"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Maker: Jason Hensley
'Website: http://www.vbcodesource.com
'Uses: Shows two ways to download a url/file from the internet via api

Private Declare Function DoFileDownload Lib "shdocvw.dll" (ByVal lpszFile As String) As Long
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Private Sub Command1_Click()

End Sub

Private Sub cmdMethod1_Click()
    Dim thePath As String
    
    thePath = InputBox("What is the url to download the file?", " File Url", "http://")
    
    'The path has to be converted to Unicode
    thePath = StrConv(thePath, vbUnicode)
    
    DoFileDownload thePath
End Sub

Private Sub Command2_Click()

End Sub

Private Sub cmdMethod2_Click()
    Dim retVal As Long 'our return value
    Dim theUrl As String 'the url you want to download
    Dim savePath As String 'where you want to save the url
    Dim pathExist As Long 'will contain our path exist or not value
    
    theUrl = InputBox("What is the url you want to download?", " Url Path?", "http://")
    If theUrl = "" Then Exit Sub
    
    savePath = InputBox("What is the path and filename to save the url to?", " Path and Filename to save")
    If savePath = "" Then Exit Sub
    
    retVal = URLDownloadToFile(0, theUrl, savePath, 0, 0)

    If retVal = 0 Then
        MsgBox "File was downloaded successfully!", vbExclamation, " Download Successful"
    Else
        MsgBox "There was a error downloading the file. Make sure that the url is valid and try again!", vbCritical, " Error"
    End If
End Sub
