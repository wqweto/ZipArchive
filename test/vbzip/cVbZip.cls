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
    Set m_oOpt = GetOpts(vArgs, "rem:o:iyp:(mem):(so)(nologo)")
    With m_oOpt
        .Item("command") = .Item("__arg__.array").Item(0)
        .Item("zip") = .Item("__arg__.array").Item(1)
        .Item("numfiles") = .Item("__arg__.array").Count - 2
        For lIdx = 1 To .Item("numfiles")
            .Item("file" & lIdx) = .Item("__arg__.array").Item(lIdx + 1)
        Next
        If Not .Item("nologo") Then
            ConsoleError "VbZip " & m_oArchive.SemVersion & " (c) 2017-2018 by wqweto@gmail.com (" & m_oArchive.ThunkBuildDate & ")" & vbCrLf & vbCrLf
        End If
        If .Exists("__unknown__") Then
            ConsoleError "Unknown option: -" & .Item("__unknown__") & vbCrLf & vbCrLf
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
            If .Item("numfiles") <= 0 Then
                ConsoleError ERR_NO_FILES & vbCrLf
                GoTo QH
            End If
            If Not pvLimitNumericOption("m", 6, 0, 9) Then
                ConsoleError ERR_INVALID_LEVEL & vbCrLf
                GoTo QH
            End If
            If Not pvLimitNumericOption("mem", 0, 0, 4) Then
                ConsoleError ERR_INVALID_STRENGTH & vbCrLf
                GoTo QH
            End If
            For lIdx = 1 To .Item("numfiles")
                sFile = .Item("file" & lIdx)
                If InStr(sFile, "*") > 0 Or InStr(sFile, "?") > 0 Then
                    '--- for *.* default to include empty folders
                    sMask = Mid$(sFile, InStrRev(sFile, "\") + 1)
                    If sMask = "*" Or sMask = "*.*" Then
                        .Item("e") = True
                    End If
                    m_oArchive.AddFromFolder sFile, Recursive:=.Item("r"), _
                        IncludeEmptyFolders:=.Item("e"), Password:=.Item("p"), EncrStrength:=.Item("mem")
                ElseIf (FileAttr(sFile) And vbDirectory) <> 0 Then
                    If .Item("r") Then
                        sMask = Mid$(sFile, InStrRev(sFile, "\") + 1)
                        .Item("e") = True
                    End If
                    m_oArchive.AddFromFolder PathCombine(sFile, "*.*"), Recursive:=.Item("r"), TargetFolder:=sMask, _
                        IncludeEmptyFolders:=.Item("e"), Password:=.Item("p"), EncrStrength:=.Item("mem")
                Else
                    m_oArchive.AddFile sFile, Password:=.Item("p"), EncrStrength:=.Item("mem")
                End If
            Next
            m_sAction = "Compressing "
            If Not m_oArchive.CompressArchive(.Item("zip"), Level:=.Item("m")) Then
                GoTo QH
            End If
        Case "l"
            If LenB(.Item("zip")) = 0 Then
                ConsoleError ERR_MISSING_ARCHIVE & vbCrLf
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
                        ConsoleError STR_LIST_HDR & vbCrLf
                        ConsoleError STR_LIST_SEP & vbCrLf
                    End If
                    sText = Format$(At(vInfo, ucsIdxLastModified), "yyyy\-MM\-dd hh\:mm\:ss")
                    sText = sText & " " & FormatAttr(At(vInfo, ucsIdxAttributes), At(vInfo, ucsIdxFlags))
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
            If Not m_oArchive.OpenArchive(.Item("zip")) Then
                GoTo QH
            End If
            m_sAction = "Testing "
            m_oArchive.Extract New cDummyStream
        Case "x"
            If LenB(.Item("zip")) = 0 Then
                ConsoleError ERR_MISSING_ARCHIVE & vbCrLf
                GoTo QH
            End If
            If Not m_oArchive.OpenArchive(.Item("zip")) Then
                GoTo QH
            End If
            m_sAction = "Extracting "
            If .Item("numfiles") = 0 Then
                m_oArchive.Extract CStr(.Item("o")), Password:=.Item("p")
            Else
                For lIdx = 1 To .Item("numfiles")
                    sFile = .Item("file" & lIdx)
                    m_oArchive.Extract CStr(.Item("o")), Filter:=sFile, Password:=.Item("p")
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
                    "  -so         write data to stdout" & vbCrLf
            Else
                ConsoleError "Invalid command: " & .Item("command") & vbCrLf
            End If
        End Select
        If Len(m_sLastRow) <> 0 Then
            ConsoleError vbCr & Space$(Len(m_sLastRow)) & vbCr
            ConsoleError "Done" & vbCrLf
        End If
    End With
    '--- success
    Init = True
QH:
    Exit Function
EH:
    ConsoleError "unhandled error: " & Err.Description & vbCrLf
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

Private Function FormatAttr(ByVal eAttr As VbFileAttribute, ByVal lFlags As Long) As String
    If eAttr = vbDirectory Then
        FormatAttr = "[DIR]"
    Else
        FormatAttr = Right$("....." & _
            IIf(eAttr And vbArchive, "A", vbNullString) & _
            IIf(eAttr And vbDirectory, "D", vbNullString) & _
            IIf(eAttr And vbReadOnly, "R", vbNullString) & _
            IIf(eAttr And vbHidden, "H", vbNullString) & _
            IIf(eAttr And vbSystem, "S", vbNullString) & _
            IIf(lFlags And 1, "E", vbNullString), 5)
    End If
End Function

Public Function GetOpts(vArgs As Variant, sOptDefs As String, Optional RetVal As Object) As Object
'---
'--- == sOptDefs format ==
'---
'--- `<single_char_opt>?(<long_opt>)?<has_arg_flag>?` where has_arg_flag=':'
'---
'--- Use `a(add):` for option with argument:
'--- ```
'---   -a --add FILE append FILE to image
'--- ```
'---
'--- Use `f(flag)(noflag)` for bool options:
'--- ```
'---   -f --flag     turn flag on
'---   --noflag      turn flag off
'--- ```
'---
'--- == JSON keys in result ==
'---
'---   Key                     Value
'---   ----------------------  --------------------------------------
'---   `<long_opt>`            Last option value if <long_opt> is specified
'---                           e.g. JsonItem(oOpt, "user") -> "bob"
'---   `<single_char_opt>`     Last option value if no <long_opt> is specified
'---                           e.g. JsonItem(oOpt, "u") -> "bob"
'---   `<long_opt> & ".array"` Array of option's all values if repeated
'---                           e.g. JsonItem(oOpt, "user.array") -> [ "bob", "ana" ]
'---   `"__arg__"`             Last optionless arg
'---   `"__arg__.array"`       All optionless args
'---   `"__arg__.opt"`         Snapshots of all options at the time each optionless args in
'--                            [__arg__.array] is parsed. This allows handling of a repeating
'---                           action: e.g. `-u userA http://url1 -u userB http://url2 ...` to
'---                           use different user for each url
'---   `"__unknown__"`         Last unmatched option
'---   `"__unknown__.array"`   All unmatched options
'---
    Dim oDefs           As Object
    Dim lIdx            As Long
    Dim sChar           As String
    Dim sShortOpt       As String
    Dim sLongOpt        As String
    Dim sHasArg         As String
    Dim lPos            As Long
    '--- output
    Dim sProp           As String
    Dim sValue          As String
    Dim oSnapshot       As Object
    Dim oItem           As Object
    
    '--- parse options definition
    For lIdx = 1 To Len(sOptDefs)
        sChar = Mid$(sOptDefs, lIdx, 1)
        '--- short
        If sChar <> "(" Then
            sShortOpt = sChar
            lIdx = lIdx + 1
            sChar = Mid$(sOptDefs, lIdx, 1)
        Else
            sShortOpt = vbNullString
        End If
        '--- long
        If sChar = "(" Then
            lPos = InStr(lIdx, sOptDefs, ")")
            sLongOpt = Mid$(sOptDefs, lIdx + 1, lPos - lIdx - 1)
            lIdx = lPos + 1
            sChar = Mid$(sOptDefs, lIdx, 1)
        Else
            sLongOpt = sShortOpt
        End If
        '--- has arg
        If sChar = ":" Then
            sHasArg = sChar
        Else
            sHasArg = vbNullString
            lIdx = lIdx - 1
        End If
        '--- populate oDefs object
        If LenB(sShortOpt) <> 0 Then
            JsonItem(oDefs, sShortOpt) = sLongOpt
        End If
        JsonItem(oDefs, "--" & sLongOpt) = sHasArg
    Next
    '--- parse current args from oDefs grammar
    For lIdx = 0 To UBound(vArgs)
        If Mid$(vArgs(lIdx), 1, 2) = "--" Then
            '--- long opt: parse "--user bob" only
            If Not IsEmpty(JsonItem(oDefs, vArgs(lIdx))) Then
                sProp = Mid$(vArgs(lIdx), 3)
ParseLongOpt:
                If JsonItem(oDefs, "--" & sProp) = ":" Then
                    lIdx = lIdx + 1
                    sValue = At(vArgs, lIdx)
                    GoTo AppendValue
                Else
                    JsonItem(RetVal, sProp) = True
                End If
            Else
                sProp = "__unknown__"
                sValue = Mid$(vArgs(lIdx), 3)
                GoTo AppendValue
            End If
        ElseIf Left$(vArgs(lIdx), 1) = "/" Or Left$(vArgs(lIdx), 1) = "-" Then
            '--- short opt: parse "/u:bob", "/u bob" and "/ubob"
            sProp = JsonItem(oDefs, Mid$(vArgs(lIdx), 2, 1))
            If LenB(sProp) <> 0 Then
                If JsonItem(oDefs, "--" & sProp) = ":" Then
                    Select Case Mid$(vArgs(lIdx), 3, 1)
                    Case ":"
                        sValue = Mid$(vArgs(lIdx), 4)
                    Case vbNullString
                        lIdx = lIdx + 1
                        sValue = At(vArgs, lIdx)
                    Case Else
                        sValue = Mid$(vArgs(lIdx), 3)
                    End Select
                    GoTo AppendValue
                Else
                    JsonItem(RetVal, sProp) = True
                End If
            Else
                '--- try to correct wrong "/long_opt" args
                sProp = Mid$(vArgs(lIdx), 2)
                If Not IsEmpty(JsonItem(oDefs, "--" & sProp)) Then
                    GoTo ParseLongOpt
                End If
                sProp = "__unknown__"
                sValue = Mid$(vArgs(lIdx), 2, 1)
                GoTo AppendValue
            End If
        Else
            '--- optionless arg
            sProp = "__arg__"
            sValue = vArgs(lIdx)
AppendValue:
            JsonItem(RetVal, sProp) = sValue
            If IsObject(JsonItem(RetVal, sProp & ".array")) Then
                Set oItem = JsonItem(RetVal, sProp & ".array")
            Else
                Set oItem = Nothing
                JsonItem(oItem, 0) = Empty
                JsonItem(RetVal, sProp & ".array") = oItem
            End If
            JsonItem(oItem, oItem.Count) = sValue
            If sProp = "__arg__" Then
                If IsObject(JsonItem(RetVal, sProp & ".opts")) Then
                    Set oItem = JsonItem(RetVal, sProp & ".opts")
                Else
                    Set oItem = Nothing
                    JsonItem(oItem, 0) = Empty
                    JsonItem(RetVal, sProp & ".opts") = oItem
                End If
                JsonClone RetVal, oSnapshot
                JsonItem(oSnapshot, sProp) = Empty
                JsonItem(oSnapshot, sProp & ".array") = Empty
                JsonItem(oSnapshot, sProp & ".opts") = Empty
                JsonItem(oItem, oItem.Count) = oSnapshot
            End If
        End If
    Next
    Set GetOpts = RetVal
End Function

Property Get JsonItem(oJson As Object, ByVal sKey As String) As Variant
    Dim vKey            As Variant
    
    If Not oJson Is Nothing Then
        vKey = IIf(CStr(C_Lng(sKey)) = sKey, C_Lng(sKey), sKey)
        If oJson.Exists(vKey) Then
            If IsObject(oJson.Item(vKey)) Then
                Set JsonItem = oJson.Item(vKey)
            Else
                JsonItem = oJson.Item(vKey)
            End If
        End If
    End If
End Property

Property Let JsonItem(oJson As Object, ByVal sKey As String, vValue As Variant)
    Dim vKey            As Variant
    
    vKey = IIf(CStr(C_Lng(sKey)) = sKey, C_Lng(sKey), sKey)
    If oJson Is Nothing Then
        Set oJson = CreateObject("Scripting.Dictionary")
'        oJson.CompareMode = IIf(VarType(vKey) = vbLong, vbBinaryCompare, vbTextCompare)
    End If
    If IsEmpty(vValue) Then
        If oJson.Exists(vKey) Then
            oJson.Remove vKey
        End If
    ElseIf IsObject(vValue) Then
        Set oJson.Item(vKey) = vValue
    Else
        oJson.Item(vKey) = vValue
    End If
End Property

Private Sub JsonClone(vSrc As Variant, vDest As Variant)
    Dim vKey            As Variant
    Dim vTemp           As Variant
    
    If IsObject(vSrc) Then
        Set vDest = CreateObject("Scripting.Dictionary")
        vDest.CompareMode = vSrc.CompareMode
        For Each vKey In vSrc.Keys
            JsonClone vSrc.Item(vKey), vTemp
            If IsObject(vTemp) Then
                Set vDest.Item(vKey) = vTemp
            Else
                vDest.Item(vKey) = vTemp
            End If
        Next
    Else
        vDest = vSrc
    End If
End Sub

Private Function C_Lng(Value As Variant) As Long
    On Error GoTo QH
    If IsNumeric(Value) Then
        C_Lng = CLng(Value)
    End If
QH:
End Function

Private Sub m_oArchive_BeforeExtract(ByVal FileIdx As Long, File As Variant, SkipFile As Boolean, Cancel As Boolean)
    If m_oOpt.Item("so") Then
        File = "stdout:"
    End If
End Sub

Private Sub m_oArchive_Error(ByVal FileIdx As Long, Source As String, Description As String, Cancel As Boolean)
    If LenB(m_sLastRow) <> 0 Then
        ConsoleError vbCrLf
    End If
    ConsoleError "Error: " & Description & vbCrLf
    m_sLastRow = vbNullString
    If Not Cancel And Not m_oOpt.Item("y") Then
        Do
            ConsoleError "Do you want to continue (y/n/a)? "
            Select Case LCase$(Left$(ConsoleReadLine(), 1))
            Case "y"
                GoTo QH
            Case "n"
                Exit Do
            Case "a"
                m_oOpt.Item("y") = True
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
        If Not m_oOpt.Item("i") Then
            sPercent = " - " & Format$(Current * 100# / Total, "0.0") & "%"
            m_sLastRow = ConsoleError(vbCr & m_sAction & At(m_oArchive.FileInfo(FileIdx), 0) & sPercent)
        End If
    Else
        vInfo = m_oArchive.FileInfo(FileIdx)
        If At(vInfo, ucsIdxSize) <> 0 Then
            sPercent = " [" & Format$(At(vInfo, ucsIdxCompressedSize) * 100# / At(vInfo, ucsIdxSize), "0.0") & "%] "
        End If
        ConsoleError IIf(LenB(m_sLastRow) <> 0, vbCr, vbNullString) & m_sAction & At(m_oArchive.FileInfo(FileIdx), 0) & sPercent & vbCrLf
    End If
End Sub
