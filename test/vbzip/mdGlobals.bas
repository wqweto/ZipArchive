Attribute VB_Name = "mdGlobals"
Option Explicit
DefObj A-Z

'=========================================================================
' API
'=========================================================================

Private Const STD_INPUT_HANDLE          As Long = -10&
Private Const STD_OUTPUT_HANDLE         As Long = -11&

Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function CharToOemBuff Lib "user32" Alias "CharToOemBuffA" (ByVal lpszSrc As String, lpszDst As Any, ByVal cchDstLength As Long) As Long
Private Declare Function OemToCharBuff Lib "user32" Alias "OemToCharBuffA" (lpszSrc As Any, ByVal lpszDst As String, ByVal cchDstLength As Long) As Long
Private Declare Function CommandLineToArgvW Lib "shell32" (ByVal lpCmdLine As Long, pNumArgs As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function ApiSysAllocString Lib "oleaut32" Alias "SysAllocString" (ByVal Ptr As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

'=========================================================================
' Functions
'=========================================================================

Public Sub Main()
    With New cVbZip
        If Not .Init(SplitArgs(Command$)) Then
            If Not InIde Then
                Call ExitProcess(1)
            End If
        End If
    End With
End Sub

Public Function ConsolePrint(ByVal sText As String, ParamArray A() As Variant) As String
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
        ReDim baBuffer(0 To Len(ConsolePrint) - 1) As Byte
        If CharToOemBuff(ConsolePrint, baBuffer(0), UBound(baBuffer) + 1) Then
            Call WriteFile(hOut, baBuffer(0), UBound(baBuffer) + 1, dwDummy, ByVal 0&)
        End If
    End If
End Function

Public Function ConsoleRead(Optional ByVal lSize As Long = 1) As String
    Dim hIn             As Long
    Dim baBuffer()      As Byte
    Dim sText           As String
    
    hIn = GetStdHandle(STD_INPUT_HANDLE)
    ReDim baBuffer(0 To lSize - 1) As Byte
    If ReadFile(hIn, baBuffer(0), UBound(baBuffer) + 1, lSize, 0) And lSize > 0 Then
        sText = String$(lSize, "!")
        Call OemToCharBuff(baBuffer(0), sText, lSize + 1)
    End If
    ConsoleRead = sText
End Function

Public Function ConsoleReadLine() As String
    Dim sChar           As String
    Dim sText           As String
    
    Do
        sChar = ConsoleRead()
        If sChar = vbLf Then
            Exit Do
        ElseIf sChar <> vbCr Then
            sText = sText & sChar
        End If
    Loop
    ConsoleReadLine = sText
End Function

Public Function SplitArgs(sText As String) As Variant
    Dim vRetVal         As Variant
    Dim lPtr            As Long
    Dim lArgc           As Long
    Dim lIdx            As Long
    Dim lArgPtr         As Long

    If LenB(sText) <> 0 Then
        lPtr = CommandLineToArgvW(StrPtr(sText), lArgc)
    End If
    If lArgc > 0 Then
        ReDim vRetVal(0 To lArgc - 1) As String
        For lIdx = 0 To UBound(vRetVal)
            Call CopyMemory(lArgPtr, ByVal lPtr + 4 * lIdx, 4)
            vRetVal(lIdx) = SysAllocString(lArgPtr)
        Next
    Else
        vRetVal = Split(vbNullString)
    End If
    Call LocalFree(lPtr)
    SplitArgs = vRetVal
End Function

Private Function SysAllocString(ByVal lPtr As Long) As String
    Dim lTemp           As Long

    lTemp = ApiSysAllocString(lPtr)
    Call CopyMemory(ByVal VarPtr(SysAllocString), lTemp, 4)
End Function

Public Function At(vArray As Variant, ByVal lIdx As Long) As Variant
    On Error GoTo QH
    At = vArray(lIdx)
QH:
End Function

Public Function FileAttr(sFile As String) As VbFileAttribute
    FileAttr = GetFileAttributes(sFile)
    If FileAttr = -1 Then
        FileAttr = &H8000
    End If
End Function

Public Function PathCombine(sPath As String, sFile As String) As String
    PathCombine = sPath & IIf(LenB(sPath) <> 0 And Right$(sPath, 1) <> "\" And LenB(sFile) <> 0, "\", vbNullString) & sFile
End Function

Public Property Get InIde() As Boolean
    Debug.Assert pvSetTrue(InIde)
End Property

Private Function pvSetTrue(bValue As Boolean) As Boolean
    bValue = True
    pvSetTrue = True
End Function

