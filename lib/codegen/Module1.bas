Attribute VB_Name = "Module1"
Option Explicit

'=========================================================================
' API
'=========================================================================

Private Const STD_OUTPUT_HANDLE         As Long = -11&

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function CharToOemBuff Lib "user32" Alias "CharToOemBuffA" (ByVal lpszSrc As String, lpszDst As Any, ByVal cchDstLength As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub vbzlib_extract_thunk Lib "debug_sshzlib.dll" (rtbl As Any, Ptr As Any)

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_hLib                  As Long
Private m_uRtbl                 As RelocTable
Private m_uInfo                 As ThunkInfo

Private Type RelocTable
    vbzlib_compress_init As Long
    vbzlib_compress_cleanup As Long
    vbzlib_compress_block As Long
    vbzlib_decompress_init As Long
    vbzlib_decompress_cleanup As Long
    vbzlib_decompress_block As Long
    vbzlib_crc32        As Long
    vbzlib_malloc       As Long
    vbzlib_realloc      As Long
    vbzlib_free         As Long
    zdat_lencodes       As Long
    zdat_distcodes      As Long
    zdat_mirrorbytes    As Long
    zdat_lenlenmap      As Long
End Type

Private Type ThunkInfo
    CodeStart           As Long
    CodeEnd             As Long
    DataStart           As Long
    DataEnd             As Long
    CodeSize            As Long
    DataSize            As Long
End Type

'=========================================================================
' Functions
'=========================================================================

Public Sub Main()
    CodegenThunk
End Sub

Public Sub CodegenThunk()
    Const ALIGN_SIZE    As Long = 16
    Const SPLIT_SIZE    As Long = 512
    Dim baBuffer()      As Byte
    Dim cOutput         As Collection
    Dim sText           As String
    Dim lIdx            As Long
    Dim lOffset         As Long
    
    m_hLib = LoadLibrary(App.Path & "\..\sshzlib\bin\debug_sshzlib.dll")
    Call vbzlib_extract_thunk(m_uRtbl, m_uInfo)
    Call FreeLibrary(m_hLib)
    m_uInfo.CodeSize = m_uInfo.CodeEnd - m_uInfo.CodeStart
'    m_uInfo.CodeSize = 6048 + 1024
    m_uInfo.DataSize = m_uInfo.DataEnd - m_uInfo.DataStart
    '--- merge and dump thunk code and data
    ReDim baBuffer(0 To pvAlign(m_uInfo.CodeSize, ALIGN_SIZE) + m_uInfo.DataSize - 1) As Byte
    Call CopyMemory(baBuffer(0), ByVal m_uInfo.CodeStart, m_uInfo.CodeSize)
    Call CopyMemory(baBuffer(pvAlign(m_uInfo.CodeSize, ALIGN_SIZE)), ByVal m_uInfo.DataStart, m_uInfo.DataSize)
    sText = ToBase64(baBuffer)
    ConsolePrint "CodeSize=%1, DataSize=%2, Len(sText)=%3" & vbCrLf, m_uInfo.CodeSize, m_uInfo.DataSize, Len(sText)
    Set cOutput = New Collection
    cOutput.Add "' Auto-generated on " & Now & ", CodeSize=" & m_uInfo.CodeSize & ", DataSize=" & m_uInfo.DataSize & ", ALIGN_SIZE=" & ALIGN_SIZE
    For lIdx = 0 To 100
        If LenB(Mid$(sText, 1 + lIdx * SPLIT_SIZE, SPLIT_SIZE)) = 0 Then
            Exit For
        End If
        If lIdx Mod 10 = 0 Then
            cOutput.Add "Private Const STR_THUNK" & (lIdx \ 10 + 1) & " As String = _"
        End If
        cOutput.Add "    """ & Mid$(sText, 1 + lIdx * SPLIT_SIZE, SPLIT_SIZE) & """" & IIf(lIdx Mod 10 <> 9 And Len(sText) > (lIdx + 1) * SPLIT_SIZE, " & _", vbNullString)
    Next
    '-- dump pfn & data offsets
    m_uRtbl.vbzlib_malloc = 0
    m_uRtbl.vbzlib_realloc = 0
    m_uRtbl.vbzlib_free = 0
    sText = vbNullString
    For lIdx = VarPtr(m_uRtbl.vbzlib_compress_init) To VarPtr(m_uRtbl.vbzlib_free) Step 4
        lOffset = Peek(lIdx)
        If lOffset <> 0 Then
            lOffset = lOffset - m_uInfo.CodeStart
        End If
        sText = sText & "|" & lOffset
    Next
    For lIdx = VarPtr(m_uRtbl.zdat_lencodes) To VarPtr(m_uRtbl.zdat_lenlenmap) Step 4
        lOffset = Peek(lIdx)
        lOffset = lOffset - m_uInfo.DataStart + pvAlign(m_uInfo.CodeSize, ALIGN_SIZE)
        sText = sText & "|" & lOffset
    Next
    cOutput.Add "Private Const STR_THUNK_OFFSETS As String = """ & Mid$(sText, 2) & """"
    cOutput.Add "Private Const STR_THUNK_BUILDDATE As String = """ & Now & """"
    cOutput.Add "' end of generated code"
    cOutput.Add vbNullString
    Clipboard.Clear
    Clipboard.SetText ConcatCollection(cOutput)
    ConsolePrint "Thunk codegen succesfully clip-copied!" & vbCrLf
End Sub

Private Function ConsolePrint(ByVal sText As String, ParamArray A() As Variant) As String
    Dim lIdx            As Long
    Dim sArg            As String
    Dim baBuffer()      As Byte
    Dim dwDummy         As Long
    Dim hOut            As Long
    
    '--- format
    For lIdx = UBound(A) To LBound(A) Step -1
        sArg = Replace(A(lIdx), "%", ChrW$(&H101))
        sText = Replace(sText, "%" & (lIdx - LBound(A) + 1), sArg)
    Next
    ConsolePrint = Replace(sText, ChrW$(&H101), "%")
    '--- output
    hOut = GetStdHandle(STD_OUTPUT_HANDLE)
    If hOut = 0 Then
        Debug.Print ConsolePrint;
    Else
        ReDim baBuffer(1 To Len(ConsolePrint)) As Byte
        If CharToOemBuff(ConsolePrint, baBuffer(1), UBound(baBuffer)) Then
            Call WriteFile(hOut, baBuffer(1), UBound(baBuffer), dwDummy, ByVal 0&)
        End If
    End If
End Function

Private Function ToBase64(baValue() As Byte) As String
    With VBA.CreateObject("MSXML2.DOMDocument").createElement("dummy")
        .DataType = "bin.base64"
        .NodeTypedValue = baValue
        ToBase64 = .Text
        ToBase64 = Replace(Replace(ToBase64, vbCrLf, vbNullString), vbLf, vbNullString)
    End With
End Function

Private Function ConcatCollection(oCol As Collection, Optional Separator As String = vbCrLf) As String
    Dim lSize           As Long
    Dim vElem           As Variant
    
    For Each vElem In oCol
        lSize = lSize + Len(vElem) + Len(Separator)
    Next
    If lSize > 0 Then
        ConcatCollection = String$(lSize - Len(Separator), 0)
        lSize = 1
        For Each vElem In oCol
            If lSize <= Len(ConcatCollection) Then
                Mid$(ConcatCollection, lSize, Len(vElem) + Len(Separator)) = vElem & Separator
            End If
            lSize = lSize + Len(vElem) + Len(Separator)
        Next
    End If
End Function

Private Function Peek(ByVal lPtr As Long) As Long
    Call CopyMemory(Peek, ByVal lPtr, 4)
End Function

Private Function pvAlign(ByVal lPtr As Long, ByVal lAlignSize As Long)
    pvAlign = (lPtr + lAlignSize - 1) And -lAlignSize
End Function
