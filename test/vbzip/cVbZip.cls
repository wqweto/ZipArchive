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
Private Const ERR_INVALID_STRENGTH  As String = "Invalid encryption method/strength. Use 0 to 3"
Private Const STR_LIST_HDR          As String = "Date       Time      Attr       Size Compressed Name"
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
    ucsIdxFlags
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

    On Error GoTo EH
    Set m_oArchive = New cZipArchive
    Set m_oOpt = GetOpt(vArgs, "mem:m:o:p:si")
    With m_oOpt
        .Item("command") = .Item("arg1")
        .Item("zip") = .Item("arg2")
        .Item("numfiles") = .Item("numarg") - 2
        For lIdx = 1 To .Item("numfiles")
            .Item("file" & lIdx) = .Item("arg" & lIdx + 2)
        Next
        If Not .Item("-nologo") Then
            ConsoleError "VbZip " & m_oArchive.SemVersion & " (c) 2017-2019 by wqweto@gmail.com (" & m_oArchive.ThunkBuildDate & ")" & vbCrLf & vbCrLf
        End If
        If LenB(.Item("zip")) <> 0 Then
            ConsoleError "Archive: " & .Item("zip") & vbCrLf
        End If
        Select Case LCase$(.Item("command"))
        Case "a"
            If LenB(.Item("zip")) = 0 Then
                ConsoleError ERR_MISSING_ARCHIVE & vbCrLf
                GoTo QH
            End If
            If Not pvLimitNumericOption("-m", 6, 0, 9) Then
                ConsoleError ERR_INVALID_LEVEL & vbCrLf
                GoTo QH
            End If
            If Not pvLimitNumericOption("-mem", 0, 0, 4) Then
                ConsoleError ERR_INVALID_STRENGTH & vbCrLf
                GoTo QH
            End If
            If .Exists("-si") Then
                m_oArchive.AddFile "stdin:", CStr(.Item("-si")), Password:=.Item("-p"), EncrStrength:=.Item("-mem")
            Else
                If .Item("numfiles") <= 0 Then
                    ConsoleError ERR_NO_FILES & vbCrLf
                    GoTo QH
                End If
                For lIdx = 1 To .Item("numfiles")
                    sFile = .Item("file" & lIdx)
                    If InStr(sFile, "*") > 0 Or InStr(sFile, "?") > 0 Then
                        '--- for *.* default to include empty folders
                        sMask = Mid$(sFile, InStrRev(sFile, "\") + 1)
                        If sMask = "*" Or sMask = "*.*" Then
                            .Item("-e") = True
                        End If
                        m_oArchive.AddFromFolder sFile, Recursive:=.Item("-r"), _
                            IncludeEmptyFolders:=.Item("-e"), Password:=.Item("-p"), EncrStrength:=.Item("-mem")
                    ElseIf (FileAttr(sFile) And vbDirectory) <> 0 Then
                        If .Item("-r") Then
                            sMask = Mid$(sFile, InStrRev(sFile, "\") + 1)
                            .Item("-e") = True
                        End If
                        m_oArchive.AddFromFolder PathCombine(sFile, "*.*"), Recursive:=.Item("-r"), TargetFolder:=sMask, _
                            IncludeEmptyFolders:=.Item("-e"), Password:=.Item("-p"), EncrStrength:=.Item("-mem")
                    Else
                        m_oArchive.AddFile sFile, Password:=.Item("-p"), EncrStrength:=.Item("-mem")
                    End If
                Next
            End If
            m_sAction = "Compressing "
            If .Item("-so") Then
                .Item("zip") = "stdout:"
            ElseIf InStrRev(.Item("zip"), ".") <= InStrRev(.Item("zip"), "\") Then
                .Item("zip") = .Item("zip") & ".zip"
            End If
            If Not m_oArchive.CompressArchive(.Item("zip"), Level:=.Item("-m")) Then
                GoTo QH
            End If
        Case "l"
            If LenB(.Item("zip")) = 0 Then
                ConsoleError ERR_MISSING_ARCHIVE & vbCrLf
                GoTo QH
            End If
            If Not pvOpenArchive(.Item("zip")) Then
                GoTo QH
            End If
            sMask = LCase$(IIf(.Item("numfiles") > 0, .Item("file1"), "*"))
            For lIdx = 0 To m_oArchive.FileCount - 1
                vInfo = m_oArchive.FileInfo(lIdx)
                sFile = LCase$(At(vInfo, ucsIdxFileName))
                If sFile Like sMask And InStr(sMask, "\") > 0 _
                        Or Mid$(sFile, InStrRev(sFile, "\") + 1) Like sMask And InStr(sMask, "\") = 0 Then
                    If lCount = 0 Then
                        ConsoleError STR_LIST_HDR & vbCrLf
                        ConsoleError STR_LIST_SEP & vbCrLf
                    End If
                    sText = Format$(At(vInfo, ucsIdxLastModified), "yyyy\-MM\-dd hh\:mm\:ss")
                    sText = sText & " " & pvFormatAttr(At(vInfo, ucsIdxAttributes), At(vInfo, ucsIdxFlags))
                    sText = sText & " " & Right$(Space$(10) & Format$(At(vInfo, ucsIdxSize), "0"), 10)
                    sText = sText & " " & Right$(Space$(10) & Format$(At(vInfo, ucsIdxCompressedSize), "0"), 10)
                    sText = sText & " " & At(vInfo, ucsIdxFileName)
                    ConsoleError sText & vbCrLf
                    lTotalSize = lTotalSize + At(vInfo, ucsIdxSize)
                    lTotalCompressed = lTotalCompressed + At(vInfo, ucsIdxCompressedSize)
                    lCount = lCount + 1
                End If
            Next
            If lCount > 0 Then
                ConsoleError STR_LIST_SEP & vbCrLf
                sText = Space$(25)
                sText = sText & " " & Right$(Space$(10) & Format$(lTotalSize, "0"), 10)
                sText = sText & " " & Right$(Space$(10) & Format$(lTotalCompressed, "0"), 10)
                sText = sText & " " & Format$(lCount, "0") & " file(s)"
                ConsoleError sText & vbCrLf
            ElseIf m_oArchive.FileCount > 0 Then
                ConsoleError "No matches found" & vbCrLf
            Else
                ConsoleError "Archive is empty" & vbCrLf
            End If
        Case "t"
            If LenB(.Item("zip")) = 0 Then
                ConsoleError ERR_MISSING_ARCHIVE & vbCrLf
                GoTo QH
            End If
            If Not pvOpenArchive(.Item("zip")) Then
                GoTo QH
            End If
            m_sAction = "Testing "
            m_oArchive.Extract New cDummyStream
        Case "x"
            If LenB(.Item("zip")) = 0 Then
                ConsoleError ERR_MISSING_ARCHIVE & vbCrLf
                GoTo QH
            End If
            If Not pvOpenArchive(.Item("zip")) Then
                GoTo QH
            End If
            m_sAction = "Extracting "
            If .Item("numfiles") = 0 Then
                m_oArchive.Extract CStr(.Item("-o")), Password:=.Item("-p")
            Else
                For lIdx = 1 To .Item("numfiles")
                    sFile = .Item("file" & lIdx)
                    m_oArchive.Extract CStr(.Item("-o")), Filter:=sFile, Password:=.Item("-p")
                Next
            End If
        Case Else
            If LenB(.Item("command")) = 0 Or .Item("command") = "?" Then
                ConsoleError "Usage: vbzip.exe <command> [-options...] <archive_file> [files]..." & vbCrLf & _
                    "e.g. vbzip.exe a backup.zip doc1.txt reports*.xls" & vbCrLf & vbCrLf & _
                    "Commands:" & vbCrLf & _
                    "  a           add files to archive" & vbCrLf & _
                    "  l           list archive contents" & vbCrLf & _
                    "  t           test archive integrity" & vbCrLf & _
                    "  x           extract files from archive" & vbCrLf & vbCrLf & _
                    "Options:" & vbCrLf & _
                    "  -r          recurse subfolders" & vbCrLf & _
                    "  -e          include empty folders" & vbCrLf & _
                    "  -m LEVEL    compression level [default: 6]" & vbCrLf & _
                    "  -o OUTPUT   output folder to extract to" & vbCrLf & _
                    "  -i          no percentage indicator" & vbCrLf & _
                    "  -y          assume yes to all questions" & vbCrLf & _
                    "  -p PASSWORD password used to encrypt/decrypt files" & vbCrLf & _
                    "  -mem METHOD encryption method" & vbCrLf & _
                    "  -so         output to stdout" & vbCrLf & _
                    "  -si NAME    input from stdin zip to NAME in archive" & vbCrLf
            Else
                ConsoleError "Invalid command: " & .Item("command") & vbCrLf
            End If
        End Select
        ConsoleError IIf(LenB(m_sLastRow) <> 0, vbCr & Space$(Len(m_sLastRow)) & vbCr, vbCrLf)
        ConsoleError "Done" & vbCrLf
    End With
    '--- success
    Init = True
QH:
    Exit Function
EH:
    ConsoleError "Unhandled error: " & Err.Description & vbCrLf
End Function

Private Function pvOpenArchive(sZipFile As String) As Boolean
    If m_oArchive.OpenArchive(sZipFile) Then
        pvOpenArchive = True
        Exit Function
    End If
    '--- if missing siffux try appending ".zip"
    If InStrRev(sZipFile, ".") <= InStrRev(sZipFile, "\") Then
        If m_oArchive.OpenArchive(sZipFile & ".zip") Then
            pvOpenArchive = True
        End If
    End If
End Function

Private Function pvLimitNumericOption(sOpt As String, dblDefault As Double, dblMin As Double, dblMax As Double) As Boolean
    Dim dblTemp         As Double
    
    With m_oOpt
        If IsEmpty(.Item(sOpt)) Then
            .Item(sOpt) = dblDefault
        ElseIf .Item(sOpt) = "0" Then
            .Item(sOpt) = 0#
        Else
            dblTemp = Val(.Item(sOpt))
            If dblTemp = 0 Or dblTemp < dblMin Or dblTemp > dblMax Then
                GoTo QH
            End If
        End If
    End With
    pvLimitNumericOption = True
QH:
End Function

Private Function pvFormatAttr(ByVal eAttr As VbFileAttribute, ByVal lFlags As Long) As String
    If eAttr = vbDirectory Then
        pvFormatAttr = "[DIR]"
    Else
        pvFormatAttr = Right$("....." & _
            IIf(eAttr And vbArchive, "A", vbNullString) & _
            IIf(eAttr And vbDirectory, "D", vbNullString) & _
            IIf(eAttr And vbReadOnly, "R", vbNullString) & _
            IIf(eAttr And vbHidden, "H", vbNullString) & _
            IIf(eAttr And vbSystem, "S", vbNullString) & _
            IIf(lFlags And 1, "E", vbNullString), 5)
    End If
End Function

'=========================================================================
' Control events
'=========================================================================

Private Sub m_oArchive_BeforeExtract(ByVal FileIdx As Long, File As Variant, SkipFile As Boolean, Cancel As Boolean)
    If m_oOpt.Item("-so") Then
        File = "stdout:"
    End If
End Sub

Private Sub m_oArchive_Error(ByVal FileIdx As Long, Source As String, Description As String, Cancel As Boolean)
    If LenB(m_sLastRow) <> 0 Then
        ConsoleError vbCrLf
    End If
    ConsoleError "Error: " & Description & IIf(m_oOpt.Item("-v"), vbCrLf & "Call-stack: " & Source, vbNullString) & vbCrLf
    m_sLastRow = vbNullString
    If Not Cancel And Not m_oOpt.Item("-y") Then
        Do
            ConsoleError "Do you want to continue (y/n/a)? "
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

    vInfo = m_oArchive.FileInfo(FileIdx)
    If Total = 0 Then
        If Not m_oOpt.Item("-i") Then
            sPercent = " [" & Format$(At(vInfo, ucsIdxCompressedSize) * 100 / IIf(Current > 0, Current, 1), "0.0") & "%] "
            ConsoleError vbCr & m_sAction & At(m_oArchive.FileInfo(FileIdx), 0) & sPercent
        End If
    ElseIf Current < Total Then
        If Not m_oOpt.Item("-i") Then
            sPercent = " - " & Format$(Current * 100# / Total, "0.0") & "%"
            m_sLastRow = ConsoleError(vbCr & m_sAction & At(m_oArchive.FileInfo(FileIdx), 0) & sPercent)
        End If
    Else
        If At(vInfo, ucsIdxSize) <> 0 Then
            sPercent = " [" & Format$(At(vInfo, ucsIdxCompressedSize) * 100# / At(vInfo, ucsIdxSize), "0.0") & "%] "
        End If
        ConsoleError IIf(LenB(m_sLastRow) <> 0, vbCr, vbNullString) & m_sAction & At(m_oArchive.FileInfo(FileIdx), 0) & sPercent & vbCrLf
    End If
End Sub
