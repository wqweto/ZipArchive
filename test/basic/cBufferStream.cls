VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBufferStream"
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

Private m_lBufferPtr            As Long
Private m_lBufferSize           As Long
Private m_lPosition             As Long
Private m_lSize                 As Long

'=========================================================================
' Properties
'=========================================================================

Public Property Get Size() As Long
    Size = m_lSize
End Property

'=========================================================================
' Methods
'=========================================================================

Public Sub Init(ByVal lBufferPtr As Long, ByVal lBufferSize As Long)
    m_lBufferPtr = lBufferPtr
    m_lBufferSize = lBufferSize
    m_lPosition = 0
    m_lSize = 0
End Sub

Private Function UnsignedAdd(ByVal lUnsignedPtr As Long, ByVal lSignedOffset As Long) As Long
    '--- note: safely add *signed* offset to *unsigned* ptr for *unsigned* retval w/o overflow in LARGEADDRESSAWARE processes
    UnsignedAdd = ((lUnsignedPtr Xor &H80000000) + lSignedOffset) Xor &H80000000
End Function

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
        Call CopyMemory(ByVal lPtr, ByVal UnsignedAdd(m_lBufferPtr, m_lPosition), lSize)
        m_lPosition = m_lPosition + lSize
    End If
    VfsReadFile = lSize
End Function

Public Function VfsWriteFile(ByVal lPtr As Long, ByVal lSize As Long) As Long
    Debug.Assert m_lPosition + lSize < m_lBufferSize
    If lSize > 0 Then
        Call CopyMemory(ByVal UnsignedAdd(m_lBufferPtr, m_lPosition), ByVal lPtr, lSize)
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

