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

Private m_lPosition             As Long
Private m_lSize                 As Long

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
    #If lPtr Then '--- touch args
    #End If
    If m_lPosition + lSize > m_lSize Then
        lSize = m_lSize - m_lPosition
    End If
    m_lPosition = m_lPosition + lSize
    VfsReadFile = lSize
End Function

Public Function VfsWriteFile(ByVal lPtr As Long, ByVal lSize As Long) As Long
    #If lPtr Then '--- touch args
    #End If
    m_lPosition = m_lPosition + lSize
    If m_lPosition > m_lSize Then
        m_lSize = m_lPosition
    End If
    VfsWriteFile = lSize
End Function

Public Sub VfsSetEndOfFile(sMetaData As String)
    #If sMetaData Then '--- touch args
    #End If
    m_lSize = m_lPosition
End Sub

