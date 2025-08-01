Attribute VB_Name = "Module2"
Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long


Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long


Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal szData As String, ByVal cbData As Long) As Long


Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long


    #If Win32 Then
        Public Const HKEY_CLASSES_ROOT = &H80000000
        Public Const HKEY_CURRENT_USER = &H80000001
        Public Const HKEY_LOCAL_MACHINE = &H80000002
        Public Const HKEY_USERS = &H80000003
        Public Const KEY_ALL_ACCESS = &H3F
        Public Const REG_OPTION_NON_VOLATILE = 0&
        Public Const REG_CREATED_NEW_KEY = &H1
        Public Const REG_OPENED_EXISTING_KEY = &H2
        Public Const ERROR_SUCCESS = 0&
        Public Const REG_SZ = (1)
    #End If


Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
    End Type
    Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

    Public Function bSetRegValue(ByVal hKey As Long, ByVal lpszSubKey As String, ByVal sSetValue As String, ByVal sValue As String) As Boolean
    On Error Resume Next
    Dim phkResult As Long
    Dim lResult As Long
    Dim SA As SECURITY_ATTRIBUTES
    Dim lCreate As Long
    RegCreateKeyEx hKey, lpszSubKey, 0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, SA, phkResult, lCreate
    lResult = RegSetValueEx(phkResult, sSetValue, 0, REG_SZ, sValue, CLng(Len(sValue) + 1))
    RegCloseKey phkResult
    bSetRegValue = (lResult = ERROR_SUCCESS)
End Function


Public Function bGetRegValue(ByVal hKey As Long, ByVal sKey As String, ByVal sSubKey As String) As String
    Dim lResult As Long
    Dim phkResult As Long
    Dim dWReserved As Long
    Dim szBuffer As String
    Dim lBuffSize As Long
    Dim szBuffer2 As String
    Dim lBuffSize2 As Long
    Dim lIndex As Long
    Dim lType As Long
    Dim sCompKey As String
    lIndex = 0
    lResult = RegOpenKeyEx(hKey, sKey, 0, 1, phkResult)


    Do While lResult = ERROR_SUCCESS And Not (bFound)
        szBuffer = Space(255)
        lBuffSize = Len(szBuffer)
        szBuffer2 = Space(255)
        lBuffSize2 = Len(szBuffer2)
        lResult = RegEnumValue(phkResult, lIndex, szBuffer, lBuffSize, dWReserved, lType, szBuffer2, lBuffSize2)


        If (lResult = ERROR_SUCCESS) Then
            sCompKey = Left(szBuffer, lBuffSize)


            If (sCompKey = sSubKey) Then
                bGetRegValue = Left(szBuffer2, lBuffSize2 - 1)
            End If
        End If
        lIndex = lIndex + 1
    Loop
    RegCloseKey phkResult
End Function
                        


'Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
'Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'Private Const HKEY_CLASSES_ROOT = &H80000000
Public Sub AssociateMyApp(ByVal sAppName As String, ByVal sEXE As String, ByVal sExt As String)
Dim lRegKey As Long
Call RegCreateKey(HKEY_CLASSES_ROOT, sExt, lRegKey)
Call RegSetValueEx(lRegKey, "", 0&, 1, ByVal sAppName, Len(sAppName))
Call RegCloseKey(lRegKey)
Call RegCreateKey(HKEY_CLASSES_ROOT, sAppName & "\Shell\Open\Command", lRegKey) ' adds info into the shell open command
Call RegSetValueEx(lRegKey, "", 0&, 1, ByVal sEXE, Len(sEXE))
Call RegCloseKey(lRegKey)
End Sub

Function GetValue(getcat, getfield, getfile) As String
    'example usage:
    'username = GetValue("UserInfo", "Userna
    '     me", "myprog.ini")
    If Dir(getfile) = "" Then Exit Function
    getcat = LCase(getcat)
    getfield = LCase(getfield)
    fnum = FreeFile
    Open getfile For Input As fnum


    Do While Not EOF(fnum)
        Line Input #fnum, l1
        l1 = Trim(l1)
        l1 = LCase(l1)


        If InStr(l1, "[") <> 0 Then


            If LCase(Mid(l1, (InStr(l1, "[") + 1), (Len(l1) - 2))) = getcat Then


                Do Until EOF(fnum) Or l2 = "["
                    Line Input #fnum, l2
                    l2 = Trim(l2)


                    If InStr(l2, "]") <> 0 Then
                        Close fnum
                        Exit Function
                    End If


                    If InStr(l2, "=") <> 0 Then


                        If LCase(Left(l2, (InStr(l2, "=") - 1))) = getfield Then
                            GetValue = Trim(Mid(l2, InStr(l2, "=") + 1, Len(l2)))
                            Close fnum
                            Exit Function
                        End If
                    End If
                Loop
            End If
        End If
    Loop
    Close fnum
End Function
Public Sub SavePlayList(list1 As ListBox, thefile$)

Open thefile$ For Output As #1
For i = 0 To list1.ListCount - 1
a$ = list1.List(i)
Print #1, a$
Next
Close 1



End Sub
Public Function GetLastBackSlash(text As String) As String
    Dim i, pos As Integer
    Dim lastslash As Integer


    For i = 1 To Len(text)
        pos = InStr(i, text, "\", vbTextCompare)
        If pos <> 0 Then lastslash = pos
    Next i
    GetLastBackSlash = Right(text, Len(text) - lastslash)
End Function
Public Sub LoadListView(list1 As ListBox, List2 As ListBox)
Dim a
Dim b
On Error Resume Next
For a = 0 To List2.ListCount - 1
b = GetLastBackSlash(List2.List(a))
list1.AddItem b
Next a
End Sub
Public Sub LoadPlayList(list1 As ListBox, filename$)

Open filename$ For Input As 1
While Not EOF(1)
Line Input #1, test
list1.AddItem RTrim(test)
Wend
Close 1

End Sub


Sub PutValue(putcat, putvar, putval, putfile)
    Dim fileCol(1 To 9000) As String
    Dim foundCat As Boolean
    Dim foundVar As Boolean
    Dim catPos As Integer
    Dim varPos As Integer
    fnum = FreeFile
    putcat = Trim(putcat)
    putcat = LCase(putcat)
    putfile = Trim(putfile)
    putfile = LCase(putfile)
    putvar = LCase(putvar)
    putvar = Trim(putvar)
    putval = LCase(putval)
    putval = Trim(putval)


    If Dir(putfile) = "" Then
        Open putfile For Append As #fnum
        Close #fnum
    End If
    Open putfile For Input As #fnum


    Do While Not EOF(fnum)


        DoEvents
            Counter = Counter + 1
            Line Input #fnum, l1
            fileCol(Counter) = l1
        Loop
        Close #fnum


        For i = 1 To Counter


            DoEvents


                If InStr(LCase(fileCol(i)), "[" & putcat & "]") <> 0 Then
                    foundCat = True
                    catPos = i


                    For X = i To Counter


                        DoEvents
                            If InStr(fileCol(X), "[") <> 0 And LCase(fileCol(X)) <> "[" & putcat & "]" Then Exit For


                            If InStr(LCase(fileCol(X)), putvar & "=") <> 0 Then
                                foundVar = True
                                varPos = X
                            End If
                        Next X
                    End If
                Next i


                If foundCat = True And foundVar = True Then
                    fileCol(varPos) = putvar & "=" & putval
                    Kill putfile
                    Open putfile For Append As #fnum


                    For i = 1 To Counter
                        Print #fnum, fileCol(i)


                        DoEvents
                        Next i
                        Close #fnum
                        Exit Sub
                    End If


                    If foundCat = True And foundVar = False Then
                        Kill putfile
                        Open putfile For Append As #fnum


                        For i = 1 To Counter
                            Print #fnum, fileCol(i)
                            If i = catPos Then Print #fnum, putvar & "=" & putval
                        Next i
                        Close #fnum
                        Exit Sub
                    End If


                    If foundCat = False And foundVar = False Then
                        Kill putfile
                        Open putfile For Append As #fnum


                        For i = 1 To Counter
                            Print #fnum, fileCol(i)
                        Next i
                        Print #fnum, "[" & putcat & "]"
                        Print #fnum, putvar & "=" & putval
                        Close #fnum
                    End If
                End Sub
