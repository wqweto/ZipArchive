VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMemoryStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=========================================================================
' API
'=========================================================================

'--- for SetFilePointer
Private Const FILE_BEGIN                As Long = 0
Private Const FILE_CURRENT              As Long = 1
Private Const FILE_END                  As Long = 2

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_baBuffer()            As Byte
Private m_lPosition             As Long
Private m_lSize                 As Long

'=========================================================================
' Properties
'=========================================================================

Public Property Get Contents() As Byte()
    If m_lSize > 0 Then
        Contents = m_baBuffer
        ReDim Preserve Contents(0 To m_lSize - 1) As Byte
    Else
        Contents = vbNullString
    End If
End Property

Public Property Let Contents(baValue() As Byte)
    m_baBuffer = baValue
    m_lPosition = 0
    m_lSize = UBound(m_baBuffer) + 1
End Property

'=========================================================================
' VFS interface
'=========================================================================

Public Function VfsSetFilePointer(lPosition As Long, ByVal lMoveMethod As Long) As Long
    Select Case lMoveMethod
    Case FILE_BEGIN
        m_lPosition = lPosition
    Case FILE_CURRENT
        m_lPosition = m_lPosition + lPosition
    Case FILE_END
        m_lPosition = m_lSize + lPosition
    End Select
    VfsSetFilePointer = m_lPosition
End Function

Public Function VfsReadFile(ByVal lPtr As Long, ByVal lSize As Long) As Long
    If m_lPosition + lSize > m_lSize Then
        lSize = m_lSize - m_lPosition
    End If
    If lSize > 0 Then
        Call CopyMemory(ByVal lPtr, m_baBuffer(m_lPosition), lSize)
        m_lPosition = m_lPosition + lSize
    End If
    VfsReadFile = lSize
End Function

Public Function VfsWriteFile(ByVal lPtr As Long, ByVal lSize As Long) As Long
    Dim lNewSize            As Long
    
    If m_lPosition + lSize > UBound(m_baBuffer) + 1 Then
        lNewSize = IIf(UBound(m_baBuffer) < 512, 512, UBound(m_baBuffer) + 1)
        Do While m_lPosition + lSize > lNewSize
            lNewSize = lNewSize * 2
        Loop
        ReDim Preserve m_baBuffer(0 To lNewSize - 1) As Byte
    End If
    If lSize > 0 Then
        Call CopyMemory(m_baBuffer(m_lPosition), ByVal lPtr, lSize)
        m_lPosition = m_lPosition + lSize
        If m_lPosition > m_lSize Then
            m_lSize = m_lPosition
        End If
    End If
    VfsWriteFile = lSize
End Function

Public Sub VfsSetEndOfFile(sMetaData As String)
    m_lSize = m_lPosition
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    m_baBuffer = vbNullString
End Sub
