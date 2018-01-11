VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cZipCrypto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
' cZipArchive (c) 2017-2018 by wqweto@gmail.com
'
' A single-class pure VB6 library for zip archives management
'
'=========================================================================
Option Explicit
DefObj A-Z
'Private Const MODULE_NAME As String = "cZipCrypto"

#Const ImplUseShared = ZIP_USESHARED

'=========================================================================
' API
'=========================================================================

Private Const MS_PRIMITIVE_PROVIDER         As String = "Microsoft Primitive Provider"
Private Const BCRYPT_SHA1_ALGORITHM         As String = "SHA1"
Private Const BCRYPT_AES_ALGORITHM          As String = "AES"
Private Const BCRYPT_HASH_LENGTH            As String = "HashDigestLength"
Private Const BCRYPT_OBJECT_LENGTH          As String = "ObjectLength"
Private Const BCRYPT_CHAINING_MODE          As String = "ChainingMode"
Private Const BCRYPT_CHAIN_MODE_ECB         As String = "ChainingModeECB"
Private Const BCRYPT_ALG_HANDLE_HMAC_FLAG   As Long = 8
Private Const BCRYPT_HASH_REUSABLE_FLAG     As Long = &H20
'--- for VariantChangeType
Private Const VT_I8                         As Long = &H14
'--- for FormatMessage
Private Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CoCreateGuid Lib "ole32" (pguid As Any) As Long
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function BCryptOpenAlgorithmProvider Lib "bcrypt.dll" (phAlgorithm As Long, ByVal pszAlgId As Long, ByVal pszImplementation As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptCloseAlgorithmProvider Lib "bcrypt.dll" (ByVal hAlgorithm As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptGetProperty Lib "bcrypt.dll" (ByVal hObject As Long, ByVal pszProperty As Long, pbOutput As Any, ByVal cbOutput As Long, cbResult As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptSetProperty Lib "bcrypt.dll" (ByVal hObject As Long, ByVal pszProperty As Long, ByVal pbInput As Long, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptGenerateSymmetricKey Lib "bcrypt.dll" (ByVal hAlgorithm As Long, phKey As Long, pbKeyObject As Any, ByVal cbKeyObject As Long, pbSecret As Any, ByVal cbSecret As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptDestroyKey Lib "bcrypt.dll" (ByVal hKey As Long) As Long
Private Declare Function BCryptEncrypt Lib "bcrypt.dll" (ByVal hKey As Long, pbInput As Any, ByVal cbInput As Long, ByVal pPaddingInfo As Long, ByVal pbIV As Long, ByVal cbIV As Long, pbOutput As Any, ByVal cbOutput As Long, pcbResult As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptDeriveKeyPBKDF2 Lib "bcrypt.dll" (ByVal pPrf As Long, pbPassword As Any, ByVal cbPassword As Long, pbSalt As Any, ByVal cbSalt As Long, ByVal cIterations As Long, ByVal dwDummy As Long, pbDerivedKey As Any, ByVal cbDerivedKey As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptCreateHash Lib "bcrypt.dll" (ByVal hAlgorithm As Long, phHash As Long, ByVal pbHashObject As Long, ByVal cbHashObject As Long, pbSecret As Any, ByVal cbSecret As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptDestroyHash Lib "bcrypt.dll" (ByVal hHash As Long) As Long
Private Declare Function BCryptHashData Lib "bcrypt.dll" (ByVal hHash As Long, pbInput As Any, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptFinishHash Lib "bcrypt.dll" (ByVal hHash As Long, pbOutput As Any, ByVal cbOutput As Long, ByVal dwFlags As Long) As Long
#If ImplUseShared = 0 Then
    Private Declare Function VariantChangeType Lib "oleaut32" (Dest As Variant, src As Variant, ByVal wFlags As Integer, ByVal vt As Long) As Long
    Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Args As Any) As Long
#End If

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const CRYPTO_AES_BLOCK_SIZE As Long = 16

Private m_sLastError            As String
Private m_uCrypto               As UcsZipCryptoType

Private Type UcsZipCryptoType
    hPbkdf2Alg          As Long
    hHmacAlg            As Long
    hHmacHash           As Long
    HmacHashLen         As Long
    hAesAlg             As Long
    hAesKey             As Long
    AesKeyObjData()     As Byte
    AesKeyObjLen        As Long
    Nonce               As Variant ' LongLong (VT_I8)
    EncrData()          As Byte
    EncrPos             As Long
    TradKey0            As Long
    TradKey1            As Variant ' LongLong (VT_I8)
    TradKey2            As Long
    TradCrc32Table()    As Long
    MAX_UINT32          As Variant ' LongLong (VT_I8)
End Type

'=========================================================================
' Properties
'=========================================================================

Property Get LastError() As String
    LastError = m_sLastError
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function CryptoAesInit(baPass() As Byte, baSalt() As Byte, ByVal lKeyLen As Long, nPassVer As Integer) As Boolean
    Dim baDerivedKey()  As Byte
    Dim lResult         As Long '--- discarded
    
    '--- init member vars
    m_uCrypto.Nonce = CLngLng(0)
    ReDim m_uCrypto.EncrData(0 To CRYPTO_AES_BLOCK_SIZE - 1) As Byte
    m_uCrypto.EncrPos = CRYPTO_AES_BLOCK_SIZE
    '--- generate RFC 2898 based derived key
    If m_uCrypto.hPbkdf2Alg = 0 Then
        If BCryptOpenAlgorithmProvider(m_uCrypto.hPbkdf2Alg, StrPtr(BCRYPT_SHA1_ALGORITHM), StrPtr(MS_PRIMITIVE_PROVIDER), BCRYPT_ALG_HANDLE_HMAC_FLAG) <> 0 Then
            GoTo QH
        End If
    End If
    ReDim baDerivedKey(0 To 2 * lKeyLen + 1) As Byte
    If BCryptDeriveKeyPBKDF2(m_uCrypto.hPbkdf2Alg, baPass(0), UBound(baPass) + 1, baSalt(0), UBound(baSalt) + 1, 1000, 0, baDerivedKey(0), UBound(baDerivedKey) + 1, 0) <> 0 Then
        GoTo QH
    End If
    '--- extract Password Verification Value from last 2 bytes of derived key
    Call CopyMemory(nPassVer, baDerivedKey(2 * lKeyLen), 2)
    '--- init AES w/ ECB from first half of derived key
    If m_uCrypto.hAesAlg = 0 Then
        If BCryptOpenAlgorithmProvider(m_uCrypto.hAesAlg, StrPtr(BCRYPT_AES_ALGORITHM), StrPtr(MS_PRIMITIVE_PROVIDER), 0) <> 0 Then
            GoTo QH
        End If
        If BCryptGetProperty(m_uCrypto.hAesAlg, StrPtr(BCRYPT_OBJECT_LENGTH), m_uCrypto.AesKeyObjLen, 4, lResult, 0) <> 0 Then
            GoTo QH
        End If
        If BCryptSetProperty(m_uCrypto.hAesAlg, StrPtr(BCRYPT_CHAINING_MODE), StrPtr(BCRYPT_CHAIN_MODE_ECB), LenB(BCRYPT_CHAIN_MODE_ECB), 0) <> 0 Then
            GoTo QH
        End If
    End If
    If m_uCrypto.hAesKey <> 0 Then
        Call BCryptDestroyKey(m_uCrypto.hAesKey)
        m_uCrypto.hAesKey = 0
    End If
    ReDim m_uCrypto.AesKeyObjData(0 To m_uCrypto.AesKeyObjLen - 1) As Byte
    If BCryptGenerateSymmetricKey(m_uCrypto.hAesAlg, m_uCrypto.hAesKey, m_uCrypto.AesKeyObjData(0), m_uCrypto.AesKeyObjLen, baDerivedKey(0), lKeyLen, 0) <> 0 Then
        GoTo QH
    End If
    '-- init HMAC from second half of derived key
    If m_uCrypto.hHmacAlg = 0 Then
        If BCryptOpenAlgorithmProvider(m_uCrypto.hHmacAlg, StrPtr(BCRYPT_SHA1_ALGORITHM), StrPtr(MS_PRIMITIVE_PROVIDER), BCRYPT_ALG_HANDLE_HMAC_FLAG) <> 0 Then
            GoTo QH
        End If
        If BCryptGetProperty(m_uCrypto.hHmacAlg, StrPtr(BCRYPT_HASH_LENGTH), m_uCrypto.HmacHashLen, 4, lResult, 0) <> 0 Then
            GoTo QH
        End If
    End If
    If m_uCrypto.hHmacHash <> 0 Then
        Call BCryptDestroyHash(m_uCrypto.hHmacHash)
        m_uCrypto.hHmacHash = 0
    End If
    If BCryptCreateHash(m_uCrypto.hHmacAlg, m_uCrypto.hHmacHash, 0, 0, baDerivedKey(lKeyLen), lKeyLen, BCRYPT_HASH_REUSABLE_FLAG) <> 0 Then
        GoTo QH
    End If
    '--- success
    CryptoAesInit = True
    Exit Function
QH:
    m_sLastError = GetSystemMessage(Err.LastDllError)
End Function

Public Sub CryptoTerminate()
    If m_uCrypto.hPbkdf2Alg <> 0 Then
        Call BCryptCloseAlgorithmProvider(m_uCrypto.hPbkdf2Alg, 0)
        m_uCrypto.hPbkdf2Alg = 0
    End If
    If m_uCrypto.hHmacHash <> 0 Then
        Call BCryptDestroyHash(m_uCrypto.hHmacHash)
        m_uCrypto.hHmacHash = 0
    End If
    If m_uCrypto.hHmacAlg <> 0 Then
        Call BCryptCloseAlgorithmProvider(m_uCrypto.hHmacAlg, 0)
        m_uCrypto.hHmacAlg = 0
    End If
    If m_uCrypto.hAesKey <> 0 Then
        Call BCryptDestroyKey(m_uCrypto.hAesKey)
        m_uCrypto.hAesKey = 0
    End If
    If m_uCrypto.hAesAlg <> 0 Then
        Call BCryptCloseAlgorithmProvider(m_uCrypto.hAesAlg, 0)
        m_uCrypto.hAesAlg = 0
    End If
End Sub

Public Function CryptoAesEncrypt(baData() As Byte, Optional ByVal Offset As Long, Optional ByVal Size As Long = -1) As Boolean
    CryptoAesEncrypt = pvCryptoAesProcess(baData, Offset, Size, HashAfter:=True)
End Function

Public Function CryptoAesDecrypt(baData() As Byte, Optional ByVal Offset As Long, Optional ByVal Size As Long = -1) As Boolean
    CryptoAesDecrypt = pvCryptoAesProcess(baData, Offset, Size, HashBefore:=True)
End Function

Private Function pvCryptoAesProcess( _
            baData() As Byte, _
            ByVal lOffset As Long, _
            ByVal lSize As Long, _
            Optional ByVal HashBefore As Boolean, _
            Optional ByVal HashAfter As Boolean) As Boolean
    Dim lPos            As Long
    Dim lIdx            As Long
    Dim lResult         As Long '--- discarded
    
    If lSize < 0 Then
        lSize = UBound(baData) + 1 - lOffset
    End If
    ReDim baNonceData(0 To CRYPTO_AES_BLOCK_SIZE - 1) As Byte
    If HashBefore Then
        If BCryptHashData(m_uCrypto.hHmacHash, baData(lOffset), lSize, 0) <> 0 Then
            GoTo QH
        End If
    End If
    With m_uCrypto
        lPos = .EncrPos
        For lIdx = lOffset To lOffset + lSize - 1
            If lPos = CRYPTO_AES_BLOCK_SIZE Then
                .Nonce = .Nonce + 1
                Call CopyMemory(baNonceData(0), ByVal VarPtr(.Nonce) + 8, 8)
                If BCryptEncrypt(.hAesKey, baNonceData(0), CRYPTO_AES_BLOCK_SIZE, 0, 0, 0, .EncrData(0), CRYPTO_AES_BLOCK_SIZE, lResult, 0) <> 0 Then
                    GoTo QH
                End If
                lPos = 0
            End If
            baData(lIdx) = baData(lIdx) Xor .EncrData(lPos)
            lPos = lPos + 1
        Next
        .EncrPos = lPos
    End With
    If HashAfter Then
        If BCryptHashData(m_uCrypto.hHmacHash, baData(lOffset), lSize, 0) <> 0 Then
            GoTo QH
        End If
    End If
    '--- success
    pvCryptoAesProcess = True
    Exit Function
QH:
    m_sLastError = GetSystemMessage(Err.LastDllError)
End Function

Public Function CryptoAesGetFinalHash(ByVal lSize As Long) As Byte()
    Dim baResult()      As Byte
    
    ReDim baResult(0 To m_uCrypto.HmacHashLen - 1) As Byte
    Call BCryptFinishHash(m_uCrypto.hHmacHash, baResult(0), m_uCrypto.HmacHashLen, 0)
    ReDim Preserve baResult(0 To lSize - 1) As Byte
    CryptoAesGetFinalHash = baResult
End Function

Public Function CryptoAesRandomSalt(ByVal lSize As Long) As Byte()
    Dim baResult()      As Byte
    
    ReDim baResult(0 To 15) As Byte
    Call CoCreateGuid(baResult(0))
    ReDim Preserve baResult(0 To lSize - 1) As Byte
    CryptoAesRandomSalt = baResult
End Function

Public Function CryptoTradInit(baPass() As Byte) As Boolean
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim lValue          As Long
    
    With m_uCrypto
        .TradKey0 = 305419896
        .TradKey1 = CLngLng(591751049)
        .TradKey2 = 878082192
        .MAX_UINT32 = CLngLng(2 ^ 32 - 1)
        If Peek(ArrPtr(.TradCrc32Table)) = 0 Then
            ReDim .TradCrc32Table(0 To 255) As Long
            For lIdx = 0 To 255
                lValue = lIdx
                For lJdx = 0 To 7
                    If (lValue And 1&) <> 0 Then
                        lValue = (((lValue And &HFFFFFFFE) \ 2&) And &H7FFFFFFF) Xor &HEDB88320
                    Else
                        lValue = ((lValue And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
                    End If
                Next
                .TradCrc32Table(lIdx) = lValue
            Next
        End If
    End With
    For lIdx = 0 To UBound(baPass)
        pvCryptoTradUpdateKeys baPass(lIdx)
    Next
    '--- success
    CryptoTradInit = True
End Function

Public Function CryptoTradDecrypt(baData() As Byte, Optional ByVal Offset As Long, Optional ByVal Size As Long = -1) As Boolean
    Dim lIdx            As Long
    Dim lTemp           As Long
    
    If Size < 0 Then
        Size = UBound(baData) + 1 - Offset
    End If
    For lIdx = Offset To Offset + Size - 1
        lTemp = m_uCrypto.TradKey2 Or 2&
        lTemp = CLng((CLngLng(lTemp) * (lTemp Xor 1&)) And &HFF00&) \ &H100&
        baData(lIdx) = baData(lIdx) Xor lTemp
        pvCryptoTradUpdateKeys baData(lIdx)
    Next
    '--- success
    CryptoTradDecrypt = True
End Function

Private Sub pvCryptoTradUpdateKeys(ByVal bytChar As Byte)
    With m_uCrypto
        .TradKey0 = pvCryptoTradCalcCrc32(.TradKey0, bytChar)
        .TradKey1 = (.TradKey1 + (.TradKey0 And &HFF&)) And .MAX_UINT32
        .TradKey1 = (.TradKey1 * 134775813 + 1&) And .MAX_UINT32
        bytChar = (.TradKey1 And &HFF000000) \ &H1000000
        .TradKey2 = pvCryptoTradCalcCrc32(.TradKey2, bytChar)
    End With
End Sub

Public Function pvCryptoTradCalcCrc32(ByVal lCrc32 As Long, ByVal bytChar As Byte) As Long
    pvCryptoTradCalcCrc32 = (((lCrc32 And &HFFFFFF00) \ &H100&) And &HFFFFFF) Xor (m_uCrypto.TradCrc32Table((lCrc32 And &HFF&) Xor bytChar))
End Function

'= shared ================================================================

#If ImplUseShared = 0 Then

Private Function GetSystemMessage(ByVal lLastDllError As Long) As String
    Dim lSize            As Long
   
    GetSystemMessage = Space$(2000)
    lSize = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0&, lLastDllError, 0&, GetSystemMessage, Len(GetSystemMessage), 0&)
    If lSize > 2 Then
        If Mid$(GetSystemMessage, lSize - 1, 2) = vbCrLf Then
            lSize = lSize - 2
        End If
    End If
    GetSystemMessage = Left$(GetSystemMessage, lSize)
End Function

Private Function CLngLng(vValue As Variant) As Variant
    Call VariantChangeType(CLngLng, vValue, 0, VT_I8)
End Function

Private Function Peek(ByVal lPtr As Long) As Long
    Call CopyMemory(Peek, ByVal lPtr, 4)
End Function

#End If ' Not ImplUseShared

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Terminate()
    CryptoTerminate
End Sub
