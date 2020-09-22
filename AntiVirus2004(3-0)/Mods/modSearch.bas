Attribute VB_Name = "modSearch"
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit
Private Sign(4096)               As String    'The Signatures will be loaded into this array
Private SignVirusType(4096)      As String * 1
Private SignVirusName(4096)      As String

Public Sub BuildSigns()

  'This builds the Signature - Array
  
  Dim sIn        As String
  Dim swords()   As String
  Dim X          As Long

  Dim Data()     As String
    sIn = FileText(AV.Signature.SignatureFilename)
    swords = Split(sIn, vbCrLf)
    ReDim Preserve swords(UBound(swords) - 1)
    sIn = ""
    For X = LBound(swords) To UBound(swords)
        Data = Split(swords(X) & ":" & ":", ":")
        Sign(X) = Data(0)
        SignVirusType(X) = Data(1)
        SignVirusName(X) = Data(2)
    Next X
    Sign(X + 1) = "#END#"
    AV.Signature.SignatureDate = Sign(0)
    AV.Signature.SignatureCount = UBound(swords) - 1

Exit Sub

    Err
    MsgBox "An error has occured while loading the signature File!" & vbCrLf & "This could be caused by an empty or damaged file!" & vbCrLf & vbCrLf & "The error message was: " & Err.Description, vbCritical + vbOKOnly, "Error"

End Sub

Public Function Search(ByVal strFilename As String) As String

  Dim Current  As Long
  Dim crc      As String

    crc = CalcCRC(strFilename)
    For Current = 1 To 4096
        If Sign(Current) = "#END#" Or LenB(Sign(Current)) = 0 Then
            GoTo Finish
        End If
        If crc = Sign(Current) Then
            DoEvents
            Search = SignVirusName(Current)
            Select Case SignVirusType(Current)
             Case "E"
                Virus.Type = Executable
             Case "S"
                Virus.Type = Script
            End Select
            Exit Function
         Else 'NOT FINDTERM(FNAME,...'NOT CRC...
            Search = "NOTHING"
        End If
        DoEvents
    Next Current
Finish:

End Function

Public Function SearchScript(ByVal strFilename As String) As Boolean

  Dim Textin         As String
  Dim temp           As Long
  Const Searchfor    As String = "DEL,KILL,FORMAT,REN,COPY,XCOPY"
  Dim Searchstring() As String

    Searchstring = Split(Searchfor, ",")
    Textin = UCase$(FileText(strFilename))
    SearchScript = False
    For temp = 0 To UBound(Searchstring)
        If InStr(1, Textin, Searchstring(temp), vbTextCompare) <> 0 Then
            SearchScript = True
            Exit Function
        End If
    Next temp

End Function
