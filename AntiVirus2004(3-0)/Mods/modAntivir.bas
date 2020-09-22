Attribute VB_Name = "modAntivir"
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit

Public Sub BuildUI()

    If AV.Runmode <> TrayOnly Then
        With frmMain
            .lblText(1).Caption = GetSetting(AV.AVname, "Settings", "countFiles", 0) & ", Virus found: " & GetSetting(AV.AVname, "Settings", "countVirus", 0)
            .lblText(3).Caption = AV.Signature.SignatureDate
            If CDate(AV.Signature.SignatureDate) < date Then
                .lblText(3).ForeColor = vbRed
                .lblText(3).ToolTipText = "It is requiered to update your signatures!"
            End If
            .lblText(5).Caption = AV.Signature.SignatureCount
            .lblText(8).Caption = GetSetting(AV.AVname, "Settings", "Startup", "OFF")
            .lblText(16).Caption = GetSetting(AV.AVname, "Settings", "Quarintine", 0)
            Unload frmSecFiles
        End With 'FRMMAIN
    End If

End Sub

Public Function CheckFile(ByVal strFilename As String) As Boolean

  Dim strResult As String
  Dim temp()    As String
  Dim path      As String
  Dim c         As Collection

    CheckFile = False
    If UCase$(Mid$(strFilename, Len(strFilename) - 4, 4)) = ".ZIP" Then
        path = UnzipFile2RandomPath(strFilename)
        modAntivir2.FullPathSearch path, c, , , , True
        DoEvents
        DelTree path
     Else 'NOT UCASE$(MID$(STRFILENAME,...
        If GetFileOI(strFilename) Then
            strResult = Search(strFilename)
            If strResult <> "NOTHING" Then
                With Virus
                    .Filename = strFilename
                    .Reason = strResult
                    temp = Split(.Filename, "\")
                    .FileNameShort = temp(UBound(temp))
                End With 'Virus
                Log "Virus found: " & Virus.Reason & " in " & Virus.Filename
                SaveSetting AV.AVname, "Settings", "countVirus", GetSetting(AV.AVname, "Settings", "countVirus", 0) + 1
                frmAlert.Show
                CheckFile = True
            End If
        End If
        If IsFileaScript(strFilename) Then
            If SearchScript(strFilename) Then
                With Virus
                    .Filename = strFilename
                    .Reason = "Destructive Code found!"
                    temp = Split(.Filename, "\")
                    .FileNameShort = temp(UBound(temp))
                End With 'Virus
                SaveSetting AV.AVname, "Settings", "countVirus", GetSetting(AV.AVname, "Settings", "countVirus", 0) + 1
                frmAlert.Show
                CheckFile = True
            End If
        End If
    End If
    SaveSetting AV.AVname, "Settings", "countFiles", GetSetting(AV.AVname, "Settings", "countFiles", 0) + 1
    BuildUI
    DoEvents

End Function

Public Function FileExist(strFilename As String) As Boolean

    On Error Resume Next
    FileExist = True
    If FileLen(strFilename) = 0 Then
        FileExist = False
    End If
    On Error GoTo 0

End Function

Public Function FileText(ByVal strFilename As String) As String

  Dim handle As Long

    handle = FreeFile
    Open strFilename For Binary As #handle
    FileText = Space$(LOF(handle))
    Get #handle, , FileText
    Close #handle

End Function

Private Function IsWinNT() As Boolean

  Dim myOS As OSVERSIONINFO

    myOS.dwOSVersionInfoSize = Len(myOS)
    GetVersionEx myOS
    IsWinNT = (myOS.dwPlatformId = VER_PLATFORM_WIN32_NT)

End Function

Public Sub KeepOnTop(F As Form)

  Const SWP_NOMOVE   As Long = 2
  Const SWP_NOSIZE   As Long = 1
  Const HWND_TOPMOST As Long = -1

    SetWindowPos F.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

End Sub

Public Function LoadIcon(size As Long, _
                         ByVal strFilename As String) As IPictureDisp
  
  Dim Result    As Long
  Dim file      As String
  Dim Unkown    As IUnknown
  Dim Icon      As IconType
  Dim CLSID     As CLSIdType
  Dim ShellInfo As ShellFileInfoType
    file = strFilename
    Call SHGetFileInfo(file, 0, ShellInfo, Len(ShellInfo), size)
    With Icon
        .cbSize = Len(Icon)
        .picType = vbPicTypeIcon
        .hIcon = ShellInfo.hIcon
    End With 'Icon
    CLSID.Id(8) = &HC0
    CLSID.Id(15) = &H46
    Result = OleCreatePictureIndirect(Icon, CLSID, 1, Unkown)
    Set LoadIcon = Unkown

End Function

Public Sub Main()

    With AV
        .AVname = "CC Antivir 2004"
        .Signature.SignatureFilename = App.path & "\signatures.db"
        .Signature.SignatureOnlineFilename = "http://www.home.r-hs.de/philippinen/antivirus/sig/signature.db"
    End With 'AV
    BuildSigns
    CheckExe
    CheckDLL
    RegisterFile ".secure", "This file is secured by " & AV.AVname, "Anti Virus", App.path & "\" & App.EXEName & ".exe /R %1", App.path & "\secicon.ico"
    Select Case UCase$(Left$(Command, 2))
     Case "/S"
        CheckFile (Mid$(Command, 3, Len(Command) - 3))
     Case vbNullString
        BuildUI
        frmMain.Show
     Case "/G"
        BuildUI
        frmMain.Show
     Case "/U"
        frmUpdate.Show
     Case "/C"
        BuildUI
        frmMain.Show
        AV.Runmode = Normal
        Call frmMain.ShowFileSearch
     Case "/T"
        frmTray.Show
        AV.Runmode = TrayOnly
     Case "/F"
        AV.Runmode = ScanFile
        Checkfolder (Mid$(Command, 3, Len(Command) - 3))
     Case "/R"
        MsgBox "This file is secured! If you want to desecure it, goto: Main/Extras/Secured Files"
        End
     Case Else
        MsgBox "Invalid Parameter!"
    End Select

End Sub

Public Sub RemoveFile(ByVal strFilename As String)

  Dim Files As String
  Dim SFO   As SHFILEOPSTRUCT

    DoEvents
    Files = strFilename & Chr$(0)
    Files = Files & Chr$(0)
    With SFO
        .hWnd = frmAlert.hWnd
        .wFunc = FO_DELETE
        .pFrom = Files
        .pTo = "" & Chr$(0)
    End With 'SFO
    Call SHFileOperation(SFO)

End Sub

Public Function ShowOpenDlg(ByVal Owner As Form, _
                            Optional ByVal InitialDir As String, _
                            Optional ByVal strFilter As String, _
                            Optional ByVal DefaultExtension As String, _
                            Optional ByVal DlgTitle As String) As String

  Dim sBuf As String

    InitialDir = IIf(IsMissing(InitialDir), vbNullString, InitialDir)
    strFilter = IIf(IsMissing(strFilter), "Alle Dateien|*.*", Replace(strFilter, "|", vbNullChar)) & vbNullChar
    DefaultExtension = IIf(IsMissing(DefaultExtension), vbNullString, DefaultExtension)
    DlgTitle = IIf(IsMissing(DlgTitle), "Datei wählen", DlgTitle)
    sBuf = Space$(256)
    If IsWinNT Then
        Call GetFileNameFromBrowseW(Owner.hWnd, StrPtr(sBuf), Len(sBuf), StrPtr(InitialDir), StrPtr(DefaultExtension), StrPtr(strFilter), StrPtr(DlgTitle))
     Else 'ISWINNT = FALSE/0
        Call GetFileNameFromBrowseA(Owner.hWnd, sBuf, Len(sBuf), InitialDir, DefaultExtension, strFilter, DlgTitle)
    End If
    ShowOpenDlg = Trim$(sBuf)

End Function



