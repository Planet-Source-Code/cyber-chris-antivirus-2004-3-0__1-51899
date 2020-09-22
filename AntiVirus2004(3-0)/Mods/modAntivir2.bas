Attribute VB_Name = "modAntivir2"
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit

Public Function CalcCRC(ByVal strFilename As String) As String

  Dim cCRC32  As New cCRC32
  Dim lCRC32  As Long
  Dim cStream As New cBinaryFileStream

    cStream.file = strFilename
    lCRC32 = cCRC32.GetFileCrc32(cStream)
    CalcCRC = Hex$(lCRC32)

End Function

Public Sub CheckExe()

    On Error GoTo Ignore
    If GetSetting(AV.AVname, "Settings", "CRC", CalcCRC(App.path & "\" & App.EXEName & ".exe")) <> CalcCRC(App.path & "\" & App.EXEName & ".exe") Then
        MsgBox "Program heavily damaged!!!", vbCritical + vbOKOnly, "Error"
        End
    End If
    SaveSetting AV.AVname, "Settungs", "CRC", CalcCRC(App.path & "\" & App.EXEName & ".exe")
Ignore:

End Sub

Public Sub Checkfolder(Optional ByVal StrFolder As String)

  Dim Result As Variant
  Dim c      As Collection

    On Error Resume Next
    If StrFolder = vbNullString Then
        Set Result = SH.BrowseForFolder(frmMain.hWnd, "Select the folder you want to have scanned", 1)
        With Result.Items.Item
            FullPathSearch .path, c, , , , True
        End With 'RESULT.ITEMS.ITEM
     Else 'NOT STRFOLDER...
        FullPathSearch StrFolder, c, , , , True
    End If
    On Error GoTo 0

End Sub

Private Function CreateKey(lhKey As Long, _
                           SubKey As String, _
                           NewSubKey As String) As Boolean

  Dim lhKeyOpen    As Long
  Dim lhKeyNew     As Long
  Dim lDisposition As Long
  Dim lResult      As Long
  Dim Security     As SECURITY_ATTRIBUTES

    lhKeyOpen = OpenKey(lhKey, SubKey, KEY_CREATE_SUB_KEY)
    lResult = RegCreateKeyEx(lhKeyOpen, NewSubKey, 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, Security, lhKeyNew, lDisposition)
    If lResult = ERROR_SUCCESS Then
        CreateKey = True
        RegCloseKey (lhKeyNew)
     Else 'NOT LRESULT...
        CreateKey = False
    End If
    RegCloseKey (lhKeyOpen)

End Function

Private Function FindFiles(ByVal strPath As String, _
                           ByRef Files As Collection, _
                           Optional ByVal strPattern As String = "*.*", _
                           Optional ByVal Attributes As VbFileAttribute = vbNormal, _
                           Optional ByVal Recursive As Boolean = True) As Long

  
  Const vbErr_PathNotFound As Long = 76
  Const INVALID_VALUE      As Long = -1
  Dim FileAttr             As Long
  Dim Filename             As String
  Dim hFind                As Long
  Dim WFD                  As WIN32_FIND_DATA

    If Mid$(strPath, Len(strPath) - 1, 1) <> "\" Then
        strPath = strPath & "\"
    End If
    If Files Is Nothing Then
        Set Files = New Collection
    End If
    strPattern = LCase$(strPattern)
    hFind = FindFirstFileA(strPath & "*", WFD)
    If hFind = INVALID_VALUE Then
        Err.Raise vbErr_PathNotFound
    End If
    Do
        Filename = LeftB$(WFD.cFileName, InStrB(WFD.cFileName, vbNullChar))
        FileAttr = GetFileAttributesA(strPath & Filename)
        If FileAttr And vbDirectory Then
            If Recursive Then
                If FileAttr <> INVALID_VALUE Then
                    If Filename <> "." Then
                        If Filename <> ".." Then
                            FindFiles = FindFiles + FindFiles(strPath & Filename, Files, strPattern, Attributes)
                        End If
                    End If
                End If
            End If
         Else 'NOT FILEATTR...
            If (FileAttr And Attributes) = Attributes Then
                If LCase$(Filename) Like strPattern Then
                    FindFiles = FindFiles + 1
                    Files.Add strPath & Filename
                End If
            End If
        End If
    Loop While FindNextFileA(hFind, WFD)
    FindClose hFind

End Function

Public Sub FullPathSearch(ByRef strPath As String, _
                          ByRef Files As Collection, _
                          Optional ByVal Compare As VbCompareMethod = vbBinaryCompare, _
                          Optional ByVal strPattern As String = "*.*", _
                          Optional ByVal Attributes As VbFileAttribute = vbNormal, _
                          Optional ByVal Recursive As Boolean = False)
  
  Dim Candidates As Collection
  Dim file       As Variant

    Running = True
    If Files Is Nothing Then
        Set Files = New Collection
    End If
    FindFiles strPath, Candidates, strPattern, Attributes, Recursive
    For Each file In Candidates
        If CheckFile(file) Then
            Exit Sub
        End If
        If Running = False Then
            Exit Sub
        End If
    Next file
    MsgBox "Scan complete!"

End Sub

Private Function OpenKey(lhKey As Long, _
                         SubKey As String, _
                         ulOptions As Long) As Long

  Dim lhKeyOpen As Long
  Dim lResult   As Long

    lhKeyOpen = 0
    lResult = RegOpenKeyEx(lhKey, SubKey, 0, ulOptions, lhKeyOpen)
    If lResult <> ERROR_SUCCESS Then
        OpenKey = 0
     Else 'NOT LRESULT...
        OpenKey = lhKeyOpen
    End If

End Function

Public Function RegisterFile(sFileExt As String, _
                             sFileDescr As String, _
                             sAppID As String, _
                             sOpenCmd As String, _
                             sIconFile As String) As Boolean

  Dim hKey      As Long
  Dim bSuccess  As Boolean
  Dim bSuccess2 As Boolean

    bSuccess = False
    hKey = HKEY_LOCAL_MACHINE
    If CreateKey(hKey, REG_PRIMARY_KEY, sFileExt) Then
        If SetValue(hKey, REG_PRIMARY_KEY & sFileExt, sAppID) Then
            If CreateKey(hKey, REG_PRIMARY_KEY, sAppID) Then
                If SetValue(hKey, REG_PRIMARY_KEY & sAppID, sFileDescr) Then
                    If CreateKey(hKey, REG_PRIMARY_KEY & sAppID, REG_SHELL_KEY & REG_SHELL_OPEN_KEY & REG_SHELL_OPEN_COMMAND_KEY) Then
                        bSuccess = SetValue(hKey, REG_PRIMARY_KEY & sAppID & "\" & REG_SHELL_KEY & REG_SHELL_OPEN_KEY & REG_SHELL_OPEN_COMMAND_KEY, sOpenCmd)
                        If CreateKey(hKey, REG_PRIMARY_KEY & sAppID, REG_ICON_KEY) Then
                            bSuccess2 = SetValue(hKey, REG_PRIMARY_KEY & sAppID & "\" & REG_ICON_KEY, sIconFile)
                        End If
                    End If
                End If
            End If
        End If
    End If
    RegisterFile = (bSuccess = bSuccess2)

End Function

Private Function SetValue(lhKey As Long, _
                          SubKey As String, _
                          sValue As String) As Boolean

  Dim lhKeyOpen As Long
  Dim lResult   As Long
  Dim lTyp      As Long
  Dim lByte     As Long

    lByte = Len(sValue)
    lTyp = REG_SZ
    lhKeyOpen = OpenKey(lhKey, SubKey, KEY_SET_VALUE)
    lResult = RegSetValue(lhKey, SubKey, lTyp, sValue, lByte)
    If lResult <> ERROR_SUCCESS Then
        SetValue = False
     Else 'NOT LRESULT...
        SetValue = True
        RegCloseKey (lhKeyOpen)
    End If

End Function



