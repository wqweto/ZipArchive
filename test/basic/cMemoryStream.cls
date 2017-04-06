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

'--- for SetFilePointer
Private Const FILE_BEGIN                As Long = 0
Private Const FILE_CURRENT              As Long = 1
Private Const FILE_END                  As Long = 2

Private Declare Function ApiEmptyByteArray Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal VarType As VbVarType = vbByte, Optional ByVal Low As Long = 0, Optional ByVal Count As Long = 0) As Byte()
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private m_baBuffer()            As Byte
Private m_lPosition             As Long
Private m_lSize                 As Long

Property Get Contents() As Byte()
    If m_lSize > 0 Then
'        ReDim Preserve m_baBuffer(0 To m_lSize - 1) As Byte
        Contents = m_baBuffer
        ReDim Preserve Contents(0 To m_lSize - 1) As Byte
    Else
        Contents = ApiEmptyByteArray()
    End If
End Property

Property Let Contents(baValue() As Byte)
    m_baBuffer = baValue
    m_lPosition = 0
    m_lSize = UBound(m_baBuffer) + 1
End Property

Public Function SetFilePointer(lPosition As Long, ByVal lMoveMethod As Long) As Boolean
    Select Case lMoveMethod
    Case FILE_BEGIN
        m_lPosition = lPosition
    Case FILE_CURRENT
        m_lPosition = m_lPosition + lPosition
    Case FILE_END
        m_lPosition = m_lSize + lPosition
    End Select
    lPosition = m_lPosition
    '--- success
    SetFilePointer = True
End Function

Public Function ReadFile(ByVal lPtr As Long, ByVal lSize As Long) As Long
    If m_lPosition + lSize > m_lSize Then
        lSize = m_lSize - m_lPosition
    End If
    If lSize > 0 Then
        Call CopyMemory(ByVal lPtr, m_baBuffer(m_lPosition), lSize)
        m_lPosition = m_lPosition + lSize
    End If
    ReadFile = lSize
End Function

Public Function WriteFile(ByVal lPtr As Long, ByVal lSize As Long) As Long
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
    WriteFile = lSize
End Function

Public Function SetEndOfFile(sMetaData As String) As Boolean
    m_lSize = m_lPosition
    '--- success
    SetEndOfFile = True
End Function

Private Sub Class_Initialize()
    m_baBuffer = ApiEmptyByteArray()
End Sub
