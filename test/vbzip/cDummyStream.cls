VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDummyStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
DefObj A-Z

'=========================================================================
' API
'=========================================================================

'--- for SetFilePointer
Private Const FILE_BEGIN                As Long = 0
Private Const FILE_CURRENT              As Long = 1
Private Const FILE_END                  As Long = 2

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_cPosition             As Currency
Private m_cSize                 As Currency

'=========================================================================
' VFS interface
'=========================================================================

Public Function VfsSetFilePointer(ByVal cPosition As Long, ByVal lMoveMethod As Long) As Long
    Select Case lMoveMethod
    Case FILE_BEGIN
        m_cPosition = cPosition
    Case FILE_CURRENT
        m_cPosition = m_cPosition + cPosition
    Case FILE_END
        m_cPosition = m_cSize + cPosition
    End Select
    VfsSetFilePointer = m_cPosition
End Function

Public Function VfsReadFile(ByVal lPtr As Long, ByVal lSize As Long) As Long
    #If lPtr Then '--- touch args
    #End If
    If m_cPosition + lSize > m_cSize Then
        lSize = m_cSize - m_cPosition
    End If
    m_cPosition = m_cPosition + lSize
    VfsReadFile = lSize
End Function

Public Function VfsWriteFile(ByVal lPtr As Long, ByVal lSize As Long) As Long
    #If lPtr Then '--- touch args
    #End If
    m_cPosition = m_cPosition + lSize
    If m_cPosition > m_cSize Then
        m_cSize = m_cPosition
    End If
    VfsWriteFile = lSize
End Function

Public Sub VfsSetEndOfFile(sMetaData As String)
    #If sMetaData Then '--- touch args
    #End If
    m_cSize = m_cPosition
End Sub

