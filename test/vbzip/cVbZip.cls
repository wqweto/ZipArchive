VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cVbZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
DefObj A-Z

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const ERR_MISSING_ARCHIVE   As String = "Missing archive"
Private Const ERR_NO_FILES          As String = "No files specified"
Private Const ERR_INVALID_LEVEL     As String = "Invalid compression level. Use 0 to 9"
Private Const STR_LIST_HDR          As String = "Date       Time     Attr        Size Compressed Name"
Private Const STR_LIST_SEP          As String = "------------------- ----- ---------- ---------- -------------------"

Private WithEvents m_oArchive   As cZipArchive
Attribute m_oArchive.VB_VarHelpID = -1
Private m_sAction               As String
Private m_sLastRow              As String
Private m_oOpt                  As Object

Private Enum UcsInfoIndexesEnum
    ucsIdxFileName
    ucsIdxAttributes
    ucsIdxCrc32
    ucsIdxSize
    ucsIdxCompressedSize
    ucsIdxComment
    ucsIdxLastModified
    ucsIdxMethod
    ucsIdxOffset
End Enum

'=========================================================================
' Methods
'=========================================================================

Public Function Init(vArgs As Variant) As Boolean
    Dim lIdx            As Long
    Dim vInfo           As Variant
    Dim sFile           As String
    Dim sMask           As String
    Dim sText           As String
    Dim lTotalSize      As Long
    Dim lTotalCompressed As Long
    Dim lCount          As Long
    Dim lLevel          As Long

    On Error GoTo EH
    Set m_oArchive = New cZipArchive
    Set m_oOpt = ParseOpt(vArgs)
    With m_oOpt
        If Not .Item("nologo") Then
            ConsolePrint "VbZip " & m_oArchive.SemVersion & " (c) 2017 by wqweto@gmail.com (" & m_oArchive.ThunkBuildDate & ")" & vbCrLf & vbCrLf
        End If
        If LenB(.Item("zip")) <> 0 Then
            ConsolePrint "Archive: " & .Item("zip") & vbCrLf
        End If
        Select Case LCase$(.Item("command"))
        Case "a"
            If LenB(.Item("zip")) = 0 Then
                ConsolePrint ERR_MISSING_ARCHIVE & vbCrLf
                GoTo QH
            End If
            If .Item("numfiles") = 0 Then
                ConsolePrint ERR_NO_FILES & vbCrLf
                GoTo QH
            End If
            If IsEmpty(.Item("-m")) Then
                lLevel = 6
            ElseIf .Item("-m") = "0" Then
                lLevel = 0
            Else
                lLevel = Val(.Item("-m"))
                If lLevel < 1 Or lLevel > 9 Then
                    ConsolePrint ERR_INVALID_LEVEL & vbCrLf
                    GoTo QH
                End If
            End If
            For lIdx = 1 To .Item("numfiles")
                sFile = .Item("file" & lIdx)
                If InStr(sFile, "*") > 0 Or InStr(sFile, "?") > 0 Then
                    '--- for *.* default to include empty folders
                    sMask = Mid$(sFile, InStrRev(sFile, "\") + 1)
                    If sMask = "*" Or sMask = "*.*" Then
                        .Item("-e") = True
                    End If
                    m_oArchive.AddFromFolder sFile, Recursive:=.Item("-r"), IncludeEmptyFolders:=.Item("-e")
                ElseIf (FileAttr(sFile) And vbDirectory) <> 0 Then
                    If .Item("-r") Then
                        sMask = Mid$(sFile, InStrRev(sFile, "\") + 1)
                        .Item("-e") = True
                    End If
                    m_oArchive.AddFromFolder PathCombine(sFile, "*.*"), Recursive:=.Item("-r"), IncludeEmptyFolders:=.Item("-e"), TargetFolder:=sMask
                Else
                    m_oArchive.AddFile sFile
                End If
            Next
            m_sAction = "Compressing "
            If Not m_oArchive.CompressArchive(.Item("zip"), Level:=lLevel) Then
                GoTo QH
            End If
        Case "l"
            If LenB(.Item("zip")) = 0 Then
                ConsolePrint ERR_MISSING_ARCHIVE & vbCrLf
                GoTo QH
            End If
            If Not m_oArchive.OpenArchive(.Item("zip")) Then
                GoTo QH
            End If
            sMask = LCase$(IIf(.Item("numfiles") > 0, .Item("file1"), "*"))
            For lIdx = 0 To m_oArchive.FileCount - 1
                vInfo = m_oArchive.FileInfo(lIdx)
                sFile = LCase$(At(vInfo, ucsIdxFileName))
                If sFile Like sMask And InStr(sMask, "\") > 0 _
                        Or Mid$(sFile, InStrRev(sFile, "\") + 1) Like sMask And InStr(sMask, "\") = 0 Then
                    If lCount = 0 Then
                        ConsolePrint STR_LIST_HDR & vbCrLf
                        ConsolePrint STR_LIST_SEP & vbCrLf
                    End If
                    sText = Format$(At(vInfo, ucsIdxLastModified), "yyyy\-MM\-dd hh\:mm\:ss")
                    sText = sText & " " & FormatAttr(At(vInfo, ucsIdxAttributes))
                    sText = sText & " " & Right$(Space(10) & Format$(At(vInfo, ucsIdxSize), "0"), 10)
                    sText = sText & " " & Right$(Space(10) & Format$(At(vInfo, ucsIdxCompressedSize), "0"), 10)
                    sText = sText & " " & At(vInfo, ucsIdxFileName)
                    ConsolePrint sText & vbCrLf
                    lTotalSize = lTotalSize + At(vInfo, ucsIdxSize)
                    lTotalCompressed = lTotalCompressed + At(vInfo, ucsIdxCompressedSize)
                    lCount = lCount + 1
                End If
            Next
            If lCount > 0 Then
                ConsolePrint STR_LIST_SEP & vbCrLf
                sText = Space(25)
                sText = sText & " " & Right$(Space(10) & Format$(lTotalSize, "0"), 10)
                sText = sText & " " & Right$(Space(10) & Format$(lTotalCompressed, "0"), 10)
                sText = sText & " " & Format$(lCount, "0") & " file(s)"
                ConsolePrint sText & vbCrLf
            Else
                ConsolePrint "No matches found" & vbCrLf
            End If
        Case "t"
            If LenB(.Item("zip")) = 0 Then
                ConsolePrint ERR_MISSING_ARCHIVE & vbCrLf
                GoTo QH
            End If
            If Not m_oArchive.OpenArchive(.Item("zip")) Then
                GoTo QH
            End If
            m_sAction = "Testing "
            m_oArchive.Extract vbNullString, OutputFile:=New cDummyStream
        Case "x"
            If LenB(.Item("zip")) = 0 Then
                ConsolePrint ERR_MISSING_ARCHIVE & vbCrLf
                GoTo QH
            End If
            If Not m_oArchive.OpenArchive(.Item("zip")) Then
                GoTo QH
            End If
            m_sAction = "Extracting "
            If .Item("numfiles") = 0 Then
                m_oArchive.Extract .Item("-o")
            Else
                For lIdx = 1 To .Item("numfiles")
                    sFile = .Item("file" & lIdx)
                    m_oArchive.Extract .Item("-o"), sFile
                Next
            End If
        Case Else
            If LenB(.Item("command")) = 0 Or .Item("command") = "?" Then
                ConsolePrint "Usage: vbzip.exe <command> [-options...] <archive_file> [files]..." & vbCrLf & _
                    "e.g. vbzip.exe a backup.zip doc1.txt reports*.xls" & vbCrLf & vbCrLf & _
                    "Commands:" & vbCrLf & _
                    "  a     add files to archive" & vbCrLf & _
                    "  l     list archive contents" & vbCrLf & _
                    "  t     test archive integrity" & vbCrLf & _
                    "  x     extract files from archive" & vbCrLf & vbCrLf & _
                    "Options:" & vbCrLf & _
                    "  -r    recurse subdirectories" & vbCrLf & _
                    "  -m    compression level (default 6)" & vbCrLf & _
                    "  -i    no percentage indicator" & vbCrLf & _
                    "  -e    include empty folders" & vbCrLf & _
                    "  -y    assume yes to all questions" & vbCrLf
            Else
                ConsolePrint "Invalid command: " & .Item("command") & vbCrLf
            End If
        End Select
        If Len(m_sLastRow) <> 0 Then
            ConsolePrint vbCr & Space(Len(m_sLastRow)) & vbCr
            ConsolePrint "Done" & vbCrLf
        End If
    End With
    '--- success
    Init = True
QH:
    Exit Function
EH:
    ConsolePrint "unhandled error: " & Err.Description & vbCrLf
End Function

Private Function FormatAttr(ByVal eAttr As VbFileAttribute) As String
    If eAttr = vbDirectory Then
        FormatAttr = "[DIR]"
    Else
        FormatAttr = Left$(IIf(eAttr And vbArchive, "A", vbNullString) & _
            IIf(eAttr And vbDirectory, "D", vbNullString) & _
            IIf(eAttr And vbReadOnly, "R", vbNullString) & _
            IIf(eAttr And vbHidden, "H", vbNullString) & _
            IIf(eAttr And vbSystem, "S", vbNullString) & ".....", 5)
    End If
End Function
    
Private Function ParseOpt(vArgs As Variant) As Object
    Dim oOpt            As Object
    Dim lIdx            As Long
    Dim bNoMoreOpt      As Boolean

    Set oOpt = CreateObject("Scripting.Dictionary")
    With oOpt
        .CompareMode = 1
        For lIdx = 0 To UBound(vArgs)
            Select Case Left$(vArgs(lIdx), 1 + bNoMoreOpt)
            Case "-", "/"
                Select Case Mid$(vArgs(lIdx), 2, 1)
                Case "o", "m"
                    If Len(vArgs(lIdx)) > 2 Then
                        .Item("-" & Mid$(vArgs(lIdx), 2, 1)) = Mid$(vArgs(lIdx), 3)
                    ElseIf LenB(vArgs(lIdx + 1)) <> 0 Then
                        .Item("-" & Mid$(vArgs(lIdx), 2, 1)) = vArgs(lIdx + 1)
                        lIdx = lIdx + 1
                    End If
                Case "-"
                    If Len(vArgs(lIdx)) = 2 Then
                        bNoMoreOpt = True
                    End If
                Case Else
                    .Item("-" & Mid$(vArgs(lIdx), 2)) = True
                End Select
            Case Else
                If LenB(.Item("command")) = 0 Then
                    .Item("command") = vArgs(lIdx)
                ElseIf LenB(.Item("zip")) = 0 Then
                    .Item("zip") = vArgs(lIdx)
                Else
                    .Item("numfiles") = .Item("numfiles") + 1
                    .Item("file" & .Item("numfiles")) = vArgs(lIdx)
                End If
            End Select
        Next
    End With
    Set ParseOpt = oOpt
End Function

Private Sub m_oArchive_Error(ByVal FileIdx As Long, Source As String, Description As String, Cancel As Boolean)
    If LenB(m_sLastRow) <> 0 Then
        ConsolePrint vbCrLf
    End If
    ConsolePrint "Error: " & Description & vbCrLf
    m_sLastRow = vbNullString
    If Not Cancel And Not m_oOpt.Item("-y") Then
        Do
            ConsolePrint "Do you want to continue (y/n/a)? "
            Select Case LCase$(Left$(ConsoleReadLine(), 1))
            Case "y"
                GoTo QH
            Case "n"
                Exit Do
            Case "a"
                m_oOpt.Item("-y") = True
                GoTo QH
            End Select
        Loop
        Cancel = True
    End If
QH:
End Sub

Private Sub m_oArchive_Progress(ByVal FileIdx As Long, ByVal Current As Long, ByVal Total As Long, Cancel As Boolean)
    Dim sPercent        As String
    Dim vInfo           As Variant

    If Current < Total Then
        If Not m_oOpt.Item("-i") Then
            sPercent = " - " & Format$(Current * 100# / Total, "0.0") & "%"
            m_sLastRow = ConsolePrint(vbCr & m_sAction & At(m_oArchive.FileInfo(FileIdx), 0) & sPercent)
        End If
    Else
        vInfo = m_oArchive.FileInfo(FileIdx)
        If At(vInfo, ucsIdxSize) <> 0 Then
            sPercent = " [" & Format$(At(vInfo, ucsIdxCompressedSize) * 100# / At(vInfo, ucsIdxSize), "0.0") & "%] "
        End If
        ConsolePrint IIf(LenB(m_sLastRow) <> 0, vbCr, vbNullString) & m_sAction & At(m_oArchive.FileInfo(FileIdx), 0) & sPercent & vbCrLf
    End If
End Sub
