VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3804
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9936
   LinkTopic       =   "Form1"
   ScaleHeight     =   3804
   ScaleWidth      =   9936
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   432
      Left            =   8400
      TabIndex        =   7
      Top             =   504
      Width           =   1356
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   432
      Left            =   6804
      TabIndex        =   6
      Top             =   504
      Width           =   1440
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   432
      Left            =   5208
      TabIndex        =   5
      Top             =   504
      Width           =   1440
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   432
      Left            =   3612
      TabIndex        =   4
      Top             =   504
      Width           =   1440
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   432
      Left            =   2016
      TabIndex        =   3
      Top             =   504
      Width           =   1440
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   432
      Left            =   420
      TabIndex        =   2
      Top             =   2016
      Width           =   1440
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   432
      Left            =   420
      TabIndex        =   0
      Top             =   504
      Width           =   1440
   End
   Begin VB.Label labProgress 
      Caption         =   "Starting. . ."
      Height          =   348
      Left            =   420
      TabIndex        =   1
      Top             =   1428
      Width           =   6144
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'--- for CreateFile
Private Const CREATE_ALWAYS         As Long = 2
Private Const GENERIC_WRITE         As Long = &H40000000

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function APIWriteFile Lib "kernel32" Alias "WriteFile" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private WithEvents m_oZip   As cZipArchive
Attribute m_oZip.VB_VarHelpID = -1
Private m_bCancel           As Boolean
Private m_dblLastUpdate     As Double
Private WithEvents m_oExtractInMemory As cZipArchive
Attribute m_oExtractInMemory.VB_VarHelpID = -1

Private Sub Command1_Click()
    Dim dblTimer        As Double
    Dim bResult         As Boolean
    Dim sLastError      As String

    dblTimer = Timer
    Set m_oZip = New cZipArchive
    m_bCancel = False
    With m_oZip
'        .AddFile "D:\TEMP\Unzip\Unzip.zip"
        .AddFile "D:\TEMP\thunk.bin"
        .AddFile "D:\TEMP\aaa.pdf", "report.pdf", Level:=0
        .AddFile "D:\TEMP\enwik8.txt"
        bResult = .CompressArchive("D:\TEMP\aaa.zip", UseUtf8:=vbTrue)
        sLastError = .LastError
    End With
    Set m_oZip = Nothing
    labProgress.Caption = IIf(bResult, "Done. ", sLastError & ". ") & Format(Timer - dblTimer, "0.000") & " elapsed"
End Sub

Private Sub Command2_Click()
    m_bCancel = True
End Sub

Private Sub Command3_Click()
    Dim dblTimer        As Double
    Dim bResult         As Boolean
    Dim sLastError      As String
    Dim oOutput         As cMemoryStream
    Dim lLevel          As Long

    dblTimer = Timer
    Set m_oZip = New cZipArchive
    Set oOutput = New cMemoryStream
    m_bCancel = False
    lLevel = 8
    With m_oZip
'        .AddFile pvInitMemStream("D:\TEMP\Unzip\Unzip.zip"), "mem/Unzip.zip"
        .AddFile pvInitMemStream("D:\TEMP\thunk.bin"), "mem/thunk.bin"
        .AddFile pvInitMemStream("D:\TEMP\aaa.pdf"), "mem/report.pdf"
        .AddFile pvInitMemStream("D:\TEMP\enwik8.txt"), "mem/enwik8.txt"
        .AddFile ReadBinaryFile("D:\TEMP\aaa.pdf"), "mem/report2.pdf"
        bResult = .CompressArchive(oOutput, Level:=lLevel)
        sLastError = .LastError
    End With
    Set m_oZip = Nothing
    WriteBinaryFile "D:\temp\aaa" & lLevel & ".zip", oOutput.Contents
    labProgress.Caption = IIf(bResult, "Done. ", sLastError & ". ") & Format(Timer - dblTimer, "0.000") & " elapsed"
End Sub

Private Function pvInitMemStream(sFile As String) As cMemoryStream
    Set pvInitMemStream = New cMemoryStream
    pvInitMemStream.Contents = ReadBinaryFile(sFile)
End Function

Private Sub Command4_Click()
    Dim dblTimer        As Double
    Dim bResult         As Boolean
    Dim sLastError      As String

    dblTimer = Timer
    Set m_oZip = New cZipArchive
    m_bCancel = False
    With m_oZip
        .OpenArchive "D:\temp\aaa2.zip"
        Set m_oExtractInMemory = m_oZip
        bResult = .Extract("D:\Temp\Unzip", "Kit\Documentation\*.md")
        Set m_oExtractInMemory = Nothing
        sLastError = .LastError
    End With
    Set m_oZip = Nothing
    labProgress.Caption = IIf(bResult, "Done. ", sLastError & ". ") & Format(Timer - dblTimer, "0.000") & " elapsed"
End Sub

Private Sub Command5_Click()
    Dim dblTimer        As Double
    Dim bResult         As Boolean
    Dim sLastError      As String
    Dim baOutput()      As Byte
    Dim oBuffer         As cBufferStream

    ChDrive "D:"
    ChDir "D:\TEMP\Unzip\SQL-Server-First-Responder-Kit-2017-02"
    dblTimer = Timer
    Set m_oZip = New cZipArchive
    m_bCancel = False
    With m_oZip
        .AddFromFolder ".\*.sql;*.md", Recursive:=True, TargetFolder:="Kit", IncludeEmptyFolders:=True
'        .AddFromFolder "D:\TEMP\Unzip\Empty\*.*", Recursive:=True, TargetFolder:="Kit", IncludeEmptyFolders:=True
        ReDim baOutput(0 To 10000000) As Byte
        Set oBuffer = New cBufferStream
        oBuffer.Init VarPtr(baOutput(0)), UBound(baOutput) + 1
        bResult = .CompressArchive(oBuffer)
        bResult = .CompressArchive("D:\TEMP\aaa3.zip")
        Debug.Assert FileLen("D:\TEMP\aaa3.zip") = oBuffer.Size
        sLastError = .LastError
    End With
    Set m_oZip = Nothing
    labProgress.Caption = IIf(bResult, "Done. ", sLastError & ". ") & Format(Timer - dblTimer, "0.000") & " elapsed"
End Sub

Private Sub Command6_Click()
    Dim dblTimer        As Double
    Dim bResult         As Boolean
    Dim sLastError      As String
    Dim baOutput()      As Byte

    dblTimer = Timer
    Set m_oZip = New cZipArchive
    m_bCancel = False
    With m_oZip
        .OpenArchive ReadBinaryFile("D:\temp\aaa2.zip")
        bResult = .Extract(baOutput, 1)
        WriteBinaryFile "D:\temp\report.pdf", baOutput
        sLastError = .LastError
        labProgress.Caption = IIf(bResult, "Done. ", sLastError & ". ") & Format(Timer - dblTimer, "0.000") & " elapsed"
        dblTimer = Timer
        Debug.Print "Size=" & UBound(baOutput) + 1 & ", CRC32=0x" & Hex$(.CalcCrc32Array(baOutput)) & ", Elapsed=" & Format$(Timer - dblTimer, "0.000")
    End With
    Set m_oZip = Nothing
End Sub

Private Sub Command7_Click()
    Dim baBuffer()      As Byte
    Dim sCompressed     As String
    Dim baOutput()      As Byte
    
    baBuffer = ReadBinaryFile("D:\TEMP\area41.pdf")
    With New cZipArchive
        If Not .DeflateBase64(baBuffer, sCompressed) Then
            MsgBox .LastError, vbExclamation
        End If
        Debug.Print UBound(baBuffer) & "->" & Len(sCompressed), Timer
        If Not .InflateBase64(sCompressed, baOutput) Then
            MsgBox .LastError, vbExclamation
        End If
    End With
    WriteBinaryFile "d:\temp\aaa.pdf", baOutput
End Sub

Private Sub m_oExtractInMemory_BeforeExtract(ByVal FileIdx As Long, File As Variant, SkipFile As Boolean, Cancel As Boolean)
    Set File = New cMemoryStream
End Sub

Private Sub m_oZip_Error(ByVal FileIdx As Long, Source As String, Description As String, Cancel As Boolean)
    Debug.Print FileIdx, Source, Description
    Cancel = m_bCancel
End Sub

Private Sub m_oZip_Progress(ByVal FileIdx As Long, ByVal Current As Long, ByVal Total As Long, Cancel As Boolean)
    Dim sPercent        As String
    
    If Total <> 0 Then
        sPercent = Format$(Current * 100# / Total, "0.0") & "%"
    End If
    labProgress.Tag = "Processing " & m_oZip.FileInfo(FileIdx)(0) & " - " & sPercent
    If m_dblLastUpdate + 0.05 < Timer Then
        labProgress.Caption = labProgress.Tag
        DoEvents
        m_dblLastUpdate = Timer
    End If
    Cancel = m_bCancel
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not m_oZip Is Nothing Then
        Cancel = 1
        m_bCancel = True
    End If
End Sub

Private Function ReadBinaryFile(sFile As String) As Byte()
    Dim baBuffer()      As Byte
    Dim nFile           As Integer
    
    On Error GoTo EH
'    baBuffer = ApiEmptyByteArray()
    nFile = FreeFile
    Open sFile For Binary Access Read As nFile
    If LOF(nFile) > 0 Then
        ReDim baBuffer(0 To LOF(nFile) - 1) As Byte
        Get nFile, , baBuffer
    End If
    Close nFile
    ReadBinaryFile = baBuffer
    Exit Function
EH:
    Close nFile
End Function

Public Sub WriteBinaryFile(sFile As String, baData() As Byte)
    Dim hFile           As Long
    Dim dwDummy         As Long
    
    hFile = CreateFile(sFile, GENERIC_WRITE, 0, 0, CREATE_ALWAYS, 0, 0)
    Call APIWriteFile(hFile, baData(0), UBound(baData) + 1, dwDummy, 0)
    Call CloseHandle(hFile)
End Sub
