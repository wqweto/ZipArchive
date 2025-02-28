VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cZipArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
' cZipArchive (c) 2017-2025 by wqweto@gmail.com
'
' A single-class pure VB6 library for zip archives management
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cZipArchive"

#Const ImplCompress = ZIP_NOCOMPRESS = 0
#Const ImplDecompress = ZIP_NOEXTRACT = 0
#Const ImplCrypto = ZIP_CRYPTO <> 0
#Const ImplUseShared = ZIP_USESHARED <> 0
#Const ImplInflate = ZIP_INFLATE <> 0
#Const ImplLogging = DEBUG_MODE

'=========================================================================
' Public enums
'=========================================================================

Public Enum ZipFileInfoIndexEnum
    zipIdxFileName = 0
    zipIdxAttributes
    zipIdxCrc32
    zipIdxSize
    zipIdxCompressedSize
    zipIdxComment
    zipIdxLastModified
    zipIdxMethod
    zipIdxOffset
    zipIdxFlags
End Enum

'=========================================================================
' Public events
'=========================================================================

Event Progress(ByVal FileIdx As Long, ByVal Current As Currency, ByVal Total As Currency, Cancel As Boolean)
Event Error(ByVal FileIdx As Long, Source As String, Description As String, Cancel As Boolean)
#If ImplCompress Then
    Event BeforeCompress(ByVal FileIdx As Long, Level As Long, SkipFile As Boolean, Cancel As Boolean)
#End If ' ImplCompress
#If ImplDecompress Then
    Event BeforeExtract(ByVal FileIdx As Long, File As Variant, SkipFile As Boolean, Cancel As Boolean)
    Event ExtractComplete(ByVal FileIdx As Long, File As Variant)
#End If ' ImplDecompress

'=========================================================================
' API
'=========================================================================

Private Const VT_BYREF                      As Long = &H4000
'--- for FindFirstFile
Private Const MAX_PATH                      As Long = 260
Private Const INVALID_HANDLE_VALUE          As Long = -1
'--- for CreateFile
Private Const CREATE_ALWAYS                 As Long = 2
Private Const OPEN_EXISTING                 As Long = &H3
Private Const GENERIC_WRITE                 As Long = &H40000000
Private Const GENERIC_READ                  As Long = &H80000000
Private Const FILE_SHARE_READ               As Long = &H1
Private Const FILE_SHARE_WRITE              As Long = &H2
'--- for VirtualAlloc
Private Const PAGE_EXECUTE_READWRITE        As Long = &H40
Private Const MEM_COMMIT                    As Long = &H1000
'--- for SetFilePointer
Private Const FILE_BEGIN                    As Long = 0
Private Const FILE_CURRENT                  As Long = 1
Private Const FILE_END                      As Long = 2
'--- for MultiByteToWideChar
Private Const CP_UTF8                       As Long = 65001
Private Const CP_OEMCP                      As Long = 1
'--- for FormatMessage
Private Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200
'--- for CryptBinaryToString/CryptStringToBinary
Private Const CRYPT_STRING_BASE64           As Long = 1
'--- for CNG
Private Const MS_PRIMITIVE_PROVIDER         As String = "Microsoft Primitive Provider"
Private Const BCRYPT_SHA1_ALGORITHM         As String = "SHA1"
Private Const BCRYPT_AES_ALGORITHM          As String = "AES"
Private Const BCRYPT_HASH_LENGTH            As String = "HashDigestLength"
Private Const BCRYPT_OBJECT_LENGTH          As String = "ObjectLength"
Private Const BCRYPT_CHAINING_MODE          As String = "ChainingMode"
Private Const BCRYPT_CHAIN_MODE_ECB         As String = "ChainingModeECB"
Private Const BCRYPT_ALG_HANDLE_HMAC_FLAG   As Long = 8
'--- for GetStdHandle
Private Const STD_INPUT_HANDLE              As Long = -10
Private Const STD_OUTPUT_HANDLE             As Long = -11
Private Const STD_ERROR_HANDLE              As Long = -12

Private Enum LongPtr
    [_]
End Enum
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileW" (ByVal lpFileName As Long, ByVal lpFindFileData As Long) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal NoSecurity As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function SetFilePointerEx Lib "kernel32" (ByVal hFile As Long, ByVal liDistanceToMove As Currency, lpNewFilePointer As Any, ByVal dwMoveMethod As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As Any, lpLastAccessTime As Any, lpLastWriteTime As Any) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesW" (ByVal lpFileName As Long, ByVal dwFileAttributes As Long) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As Any, lpLocalFileTime As FILETIME) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare Function SystemTimeToVariantTime Lib "oleaut32" (lpSystemTime As SYSTEMTIME, pvTime As Date) As Long
Private Declare Function VariantTimeToSystemTime Lib "oleaut32" (ByVal vTime As Date, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleW" (ByVal lpModuleName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcW" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, Optional ByVal Msg As Long, Optional ByVal wParam As Long, Optional ByVal lParam As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cbMultiByte As Long, lpWideCharStr As Any, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableW" (ByVal lpName As Long, ByVal lpBuffer As Long, ByVal nSize As Long) As Long
Private Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableW" (ByVal lpName As Long, ByVal lpValue As Long) As Long
Private Declare Function PathMatchSpecW Lib "shlwapi" (ByVal pszFileParam As Long, ByVal pszSpec As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#If ImplCompress Then
    Private Declare Function FileTimeToDosDateTime Lib "kernel32" (lpFileTime As FILETIME, ByVal lpFatDate As Long, ByVal lpFatTime As Long) As Long
    Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileW" (ByVal hFindFile As Long, ByVal lpFindFileData As Long) As Long
#End If ' ImplCompress
#If ImplDecompress Then
    Private Declare Function DosDateTimeToFileTime Lib "kernel32" (ByVal wFatDate As Long, ByVal wFatTime As Long, lpFileTime As FILETIME) As Long
#End If ' ImplDecompress
#If ImplCrypto Then
    Private Declare Function RtlGenRandom Lib "advapi32" Alias "SystemFunction036" (RandomBuffer As Any, ByVal RandomBufferLength As Long) As Long
    Private Declare Function BCryptOpenAlgorithmProvider Lib "bcrypt" (phAlgorithm As Long, ByVal pszAlgId As Long, ByVal pszImplementation As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptCloseAlgorithmProvider Lib "bcrypt" (ByVal hAlgorithm As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptGetProperty Lib "bcrypt" (ByVal hObject As Long, ByVal pszProperty As Long, pbOutput As Any, ByVal cbOutput As Long, cbResult As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptSetProperty Lib "bcrypt" (ByVal hObject As Long, ByVal pszProperty As Long, ByVal pbInput As Long, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptGenerateSymmetricKey Lib "bcrypt" (ByVal hAlgorithm As Long, phKey As Long, pbKeyObject As Any, ByVal cbKeyObject As Long, pbSecret As Any, ByVal cbSecret As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptDestroyKey Lib "bcrypt" (ByVal hKey As Long) As Long
    Private Declare Function BCryptEncrypt Lib "bcrypt" (ByVal hKey As Long, pbInput As Any, ByVal cbInput As Long, ByVal pPaddingInfo As Long, ByVal pbIV As Long, ByVal cbIV As Long, pbOutput As Any, ByVal cbOutput As Long, pcbResult As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptDeriveKeyPBKDF2 Lib "bcrypt" (ByVal pPrf As Long, pbPassword As Any, ByVal cbPassword As Long, pbSalt As Any, ByVal cbSalt As Long, ByVal cIterations As Long, ByVal dwDummy As Long, pbDerivedKey As Any, ByVal cbDerivedKey As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptCreateHash Lib "bcrypt" (ByVal hAlgorithm As Long, phHash As Long, ByVal pbHashObject As Long, ByVal cbHashObject As Long, pbSecret As Any, ByVal cbSecret As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptDestroyHash Lib "bcrypt" (ByVal hHash As Long) As Long
    Private Declare Function BCryptHashData Lib "bcrypt" (ByVal hHash As Long, pbInput As Any, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptFinishHash Lib "bcrypt" (ByVal hHash As Long, pbOutput As Any, ByVal cbOutput As Long, ByVal dwFlags As Long) As Long
#End If ' ImplCrypto
#If ImplUseShared = 0 Then
    Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageW" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As Long, ByVal nSize As Long, Args As Any) As Long
    Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long
    Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryW" (ByVal lpPathName As Long, ByVal lpSecurityAttributes As Long) As Long
    Private Declare Function CryptBinaryToString Lib "crypt32" Alias "CryptBinaryToStringW" (ByVal pbBinary As Long, ByVal cbBinary As Long, ByVal dwFlags As Long, ByVal pszString As Long, ByRef pcchString As Long) As Long
    Private Declare Function CryptStringToBinary Lib "crypt32" Alias "CryptStringToBinaryW" (ByVal pszString As Long, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, ByRef pcbBinary As Long, ByRef pdwSkip As Long, ByRef pdwFlags As Long) As Long
#End If ' Not ImplUseShared

Private Type FILETIME
    dwLowDateTime       As Long
    dwHighDateTime      As Long
End Type

Private Type SYSTEMTIME
    wYear               As Integer
    wMonth              As Integer
    wDayOfWeek          As Integer
    wDay                As Integer
    wHour               As Integer
    wMinute             As Integer
    wSecond             As Integer
    wMilliseconds       As Integer
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes    As Long
    ftCreationTime      As FILETIME
    ftLastAccessTime    As FILETIME
    ftLastWriteTime     As FILETIME
    nFileSizeHigh       As Long
    nFileSizeLow        As Long
    dwReserved0         As Long
    dwReserved1         As Long
    cFileName           As String * MAX_PATH
    cAlternate          As String * 14
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const LNG_SIG_LOCAL         As Long = &H4034B50
Private Const LNG_SIG_CENTRAL       As Long = &H2014B50
Private Const LNG_SIG_END           As Long = &H6054B50
Private Const LNG_SIG_DATADESC      As Long = &H7064B50
Private Const LNG_SIG_EXTLOCAL      As Long = &H8074B50
Private Const LNG_SIG_EOCDR         As Long = &H6064B50
Private Const LNG_SIG_LOCATOR       As Long = &H7064B50
Private Const LNG_LEN_LOCAL         As Long = 30
Private Const LNG_LEN_CENTRAL       As Long = 46
Private Const LNG_LEN_END           As Long = 22
Private Const LNG_LEN_EXTLOCAL      As Long = 16
Private Const LNG_LEN_EOCDR         As Long = 56
Private Const LNG_LEN_LOCATOR       As Long = 20
Private Const LNG_VER_DEFAULT       As Long = 10
Private Const LNG_VER_DEFLATE       As Long = 20
Private Const LNG_VER_ZIP64         As Long = 45
Private Const LNG_METHOD_DEFLATE    As Long = 8
Private Const LNG_METHOD_STORE      As Long = 0
Private Const LNG_IOBUF_SIZE        As Long = 65536
Private Const LNG_METHOD_AES        As Long = 99
Private Const LNG_EXTRADATA_AES_HEADER As Integer = &H9901
Private Const LNG_EXTRADATA_AES_VENDOR As Integer = &H4541 '--- "AE"
Private Const LNG_EXTRADATA_AES_SIZE As Long = 7
Private Const LNG_EXTRADATA_AE_2_MAXSIZE As Long = 20
Private Const LNG_EXTRADATA_ZIP64_HEADER As Long = 1
Private Const LNG_EXTRADATA_ZIP64_LOCAL_SIZE As Long = 16
Private Const LNG_EXTRADATA_INFOZIP_UNICODE_PATH As Long = &H7075
Private Const LNG_ENC_HEADER_SIZE   As Long = 12
Private Const LNG_AES_AUTHCODE_SIZE As Long = 10
Private Const LNG_AES_BLOCK_SIZE    As Long = 16
Private Const LNG_PIPE_ENDED_ERROR  As Long = 109
Private Const LNG_SIZE_LIMIT        As Currency = 2 ^ 32 - 2
Private Const LNG_ENTRIES_LIMIT     As Long = 2 ^ 16 - 2
'--- errors
Private Const ERR_USER_CANCEL       As String = "User cancelled"
Private Const ERR_INIT_COMPRESSOR   As String = "Cannot init deflate compressor"
Private Const ERR_COMPRESSING       As String = "Error compressing"
Private Const ERR_EMPTY_ARCHIVE     As String = "Empty archive"
Private Const ERR_INVALID_ARCHIVE   As String = "Invalid archive"
Private Const ERR_ARCHIVE_NOT_OPEN  As String = "Archive not opened"
Private Const ERR_INVALID_LOCALHDR  As String = "Invalid local header"
Private Const ERR_UNSUPPORTED_METHOD As String = "Unsupported compression method %1"
Private Const ERR_INIT_DECOMPRESSOR As String = "Cannot init inflate decompressor"
Private Const ERR_READING_ARCHIVE   As String = "Error reading archive"
Private Const ERR_DECOMPRESSING     As String = "Error decompressing"
Private Const ERR_CRC_CHECK         As String = "CRC check failed"
Private Const ERR_ENTRY_INVALID_SIG As String = "Entry %1 has invalid signature"
Private Const ERR_ENTRY_NO_FILENAME As String = "Entry %1 has no filename"
Private Const ERR_UNSUPPORTED_ENCRYPTION As String = "Unsupported encryption"
Private Const ERR_PASSWORD_REQUIRED As String = "Password required"
Private Const ERR_INVALID_PASSWORD  As String = "Invalid password"
Private Const ERR_INVALID_AUTHCODE  As String = "Invalid authentication code"
Private Const ERR_INVALID_BYTEARRAY As String = "Invalid byte array"
Private Const ERR_MISSING_ZIP64     As String = "Missing Zip64 extension above 4GB limit"
Private Const ERR_NO_CRYPTO_COMPILED As String = "Crypto not compiled (ZIP_CRYPTO = 1)"
'--- strings
Private Const STR_BUFFER            As String = "[buffer]"
Private Const STR_STREAM            As String = "[stream]"
Private Const STR_ENTRY             As String = "[entry %1]"
Private Const STR_ARCHIVE           As String = "[archive]"
Private Const STR_THUNK_OFFSETS     As String = "592|656|704|784|800|912|0|288|368|416|0|0|0|6896|7264|6640|7244"
Private Const STR_THUNK_BUILDDATE   As String = "12.1.2018 17:15:52"

Private m_uRtbl                 As ZipRelocTableType
Private m_sLastError            As String
Private m_bCancel               As Boolean
Private m_lFileCount            As Long
Private m_uFiles()              As ZipFileInfo
Private m_lCurrentFile          As Long
Private m_lCodePage             As Long
#If ImplDecompress Then
    Private m_sComment          As String
    Private m_vArchiveFile      As Variant
#End If ' ImplDecompress
#If ImplCrypto Then
    Private m_uCrypto           As ZipCryptoType
#End If ' ImplCrypto
Private m_baBuffer()            As Byte
Private m_lBuffIdx              As Long

Private Type ZipVfsType
    Handle              As Long
    Stream              As Object
    BufferArray         As Variant
    BufferBase          As LongPtr
    BufferSize          As Long
    BufferPtr           As LongPtr
    Data                As WIN32_FIND_DATA
    FileName            As String
    SourceArchive       As cZipArchive
    SourceFileIdx       As Long
    SourceFileInfo      As Variant
End Type

Private Type ZipRelocTableType
    CompressInit        As Long
    CompressCleanup     As Long
    CompressBlock       As Long
    DecompressInit      As Long
    DecompressCleanup   As Long
    DecompressBlock     As Long
    CalcCrc32           As Long
    MemNonce            As Long
    MemXor              As Long
    ZipCrypt            As Long
    MallocImpl          As Long
    ReallocImpl         As Long
    FreeImpl            As Long
    LenCodes            As Long
    DistCodes           As Long
    MirrorBytes         As Long
    LenLenMap           As Long
    Crc32Table          As Long
End Type

Private Type ZipBuffersType
    InBlock             As Long
    InLen               As Long
    OutBlock            As Long
    OutLen              As Long
    Final               As Long
    Greedy              As Long
    MaxMatch            As Long
    NiceLen             As Long
End Type

Private Enum ZipRelocationIndexesEnum
    zipRelCompressInit = 0
    zipRelCompressCleanup
    zipRelCompressBlock
    zipRelDecompressInit
    zipRelDecompressCleanup
    zipRelDecompressBlock
    zipRelCalcCrc32
    zipRelMemNonce
    zipRelMemXor
    zipRelZipCrypt
    zipRelMallocImpl
    zipRelReallocImpl
    zipRelFreeImpl
    zipRelLenCodes
    zipRelDistCodes
    zipRelMirrorBytes
    zipRelLenLenMap
End Enum

Private Enum ZipHeaderFlagsEnum
    zipFlagEncrypted = 2 ^ 0                         ' If set, indicates that the file is encrypted
    zipFlagCompressOptionMask = 2 ^ 1 Or 2 ^ 2       ' 0 - normal, 1 - maximum, 2 - fast, 3 - super fast
    zipFlagHasDataDescriptor = 2 ^ 3                 ' Values are put in the data descriptor immediately following the compressed data
    zipFlagStrongEncrypted = 2 ^ 6                   ' Strong encryption
    zipFlagUseUtf8 = 2 ^ 11                          ' Language encoding flag (EFS)
End Enum

Private Type Zip64ExtraDataType
    HeaderId            As Integer
    DataSize            As Integer
    USize               As Currency
    CSize               As Currency
    Offset              As Currency
    DiskNum             As Long
End Type

Private Type Zip64EndHeaderType
    Signature           As Long
    Size                As Currency
    VerMade             As Integer
    VerExt              As Integer
    DiskNum             As Long
    DiskStart           As Long
    Entries             As Currency
    TotEntr             As Currency
    CenSize             As Currency
    CenOff              As Currency
End Type

Private Type Zip64EndLocatorType
    Signature           As Long
    DiskEocdr           As Long
    Offset              As Currency
    DiskNum             As Long
End Type

Private Type ZipLocalHeaderType
    Signature           As Long                     ' Signature
    VerExt              As Integer                  ' version needed to extract
    Flags               As Integer                  ' encrypt and compression flags
    Method              As Integer                  ' compression method
    FTime               As Integer                  ' time last modifies, dos format
    FDate               As Integer                  ' date last modifies, dos format
    '--- padding
    Crc32               As Long                     ' CRC32 for uncompressed file
    CSize               As Long                     ' compressed size
    USize               As Long                     ' uncompressed size
    LenFname            As Integer                  ' Length filename
    LenExt              As Integer                  ' Length for extra field
    Ext                 As Zip64ExtraDataType
End Type

Private Type ZipCentralHeaderType
    Signature           As Long                     ' Signature
    VerMade             As Integer                  ' version made by
    VerExt              As Integer                  ' version needed to extract
    Flags               As Integer                  ' encrypt and compression flags
    Method              As Integer                  ' compression method
    FTime               As Integer                  ' time last modifies, dos format
    FDate               As Integer                  ' date last modifies, dos format
    '--- padding
    Crc32               As Long                     ' CRC32 for uncompressed file
    CSize               As Long                     ' compressed size
    USize               As Long                     ' uncompressed size
    LenFname            As Integer                  ' Length filename
    LenExt              As Integer                  ' Length for extra field
    LenCom              As Integer                  ' Length for comment field
    DiskStart           As Integer                  ' start disk number
    AttribI             As Integer                  ' internal file attributes
    '--- padding
    AttribX             As Long                     ' external file attributes
    Offset              As Long                     ' relative offset of local header
    Ext                 As Zip64ExtraDataType
End Type

Private Type ZipEndHeaderType
    Signature           As Long                     ' Signature
    DiskNum             As Integer                  ' This disk number
    DiskStart           As Integer                  ' Start disk number
    Entries             As Integer                  ' Entries on this disk
    TotEntr             As Integer                  ' Number of total entries
    CenSize             As Long                     ' Size of entire cetral directory
    CenOff              As Long                     ' Offset of central on starting disk
    LenCom              As Integer                  ' Length of comment field
    Ext                 As Zip64EndHeaderType
End Type

Private Type ZipAesExtraDataType
    HeaderId            As Integer                  ' Extra field header ID (0x9901)
    DataSize            As Integer                  ' Data size (currently 7, but subject to possible increase in the future)
    Version             As Integer                  ' Integer version number specific to the zip vendor
    VendorId            As Integer                  ' 2-character vendor ID
    Strength            As Byte                     ' Integer mode value indicating AES encryption strength
    '--- padding
    Method              As Integer                  ' The actual compression method used to compress the file
End Type

Private Type ZipFileInfo
    FileName            As String
    Attributes          As Long
    Crc32               As Long
    Size                As Currency
    CompressedSize      As Currency
    Comment             As String
    LastModified        As Date
    Method              As Long
    Offset              As Currency
    Flags               As Long
    Extra()             As Byte
    SourceFile          As Variant
    Level               As Long
#If ImplCrypto Then
    Aes                 As ZipAesExtraDataType
    DecDat              As Byte
    Password            As String
#End If ' ImplCrypto
End Type

Private Type ZipCryptoType
    hPbkdf2Alg          As Long
    hHmacAlg            As Long
    hHmacHash           As Long
    HmacLen             As Long
    hAesAlg             As Long
    hAesKey             As Long
    AesKeyObjData()     As Byte
    AesKeyObjLen        As Long
    Nonce(0 To 1)       As Long
    EncrData()          As Byte
    EncrPos             As Long
    TradKey(0 To 3)     As Long
End Type

'=========================================================================
' Properties
'=========================================================================

Public Property Get SemVersion() As String
    SemVersion = "0.3.2"
End Property

Public Property Get ThunkBuildDate() As String
    ThunkBuildDate = STR_THUNK_BUILDDATE
End Property

Public Property Get LastError() As String
    LastError = m_sLastError
End Property

Public Property Let LastError(sValue As String)
    m_sLastError = sValue
End Property

Public Property Get FileCount() As Long
    FileCount = m_lFileCount
End Property

Public Property Let FileCount(ByVal lValue As Long)
    m_lFileCount = lValue
End Property

Public Property Get FileInfo(FileIdx As Variant, Optional ByVal InfoIdx As ZipFileInfoIndexEnum = -1) As Variant
    Dim sFileName       As String
    Dim lIdx            As Long
    
    If VarType(FileIdx) = vbString Then
        sFileName = LCase$(FileIdx)
        For lIdx = 0 To m_lFileCount - 1
            If LCase$(m_uFiles(lIdx).FileName) = sFileName Then
                Exit For
            End If
        Next
    Else
        lIdx = FileIdx
    End If
    If lIdx >= 0 And lIdx < m_lFileCount Then
        With m_uFiles(lIdx)
            Select Case InfoIdx
            Case Is < 0
                FileInfo = Array(.FileName, .Attributes, .Crc32, .Size, .CompressedSize, .Comment, .LastModified, .Method, .Offset, .Flags)
            Case zipIdxFileName
                FileInfo = .FileName
            Case zipIdxAttributes
                FileInfo = .Attributes
            Case zipIdxCrc32
                FileInfo = .Crc32
            Case zipIdxSize
                FileInfo = .Size
            Case zipIdxCompressedSize
                FileInfo = .CompressedSize
            Case zipIdxComment
                FileInfo = .Comment
            Case zipIdxLastModified
                FileInfo = .LastModified
            Case zipIdxMethod
                FileInfo = .Method
            Case zipIdxOffset
                FileInfo = .Offset
            Case zipIdxFlags
                FileInfo = .Flags
            End Select
        End With
    End If
End Property

Public Property Get CodePage() As Long
    CodePage = m_lCodePage
End Property

Public Property Let CodePage(ByVal lValue As Long)
    m_lCodePage = lValue
End Property

#If ImplDecompress Then

Public Property Get Comment() As String
    Comment = m_sComment
End Property

Friend Property Get frArchiveFile() As Variant
    AssignVariant frArchiveFile, m_vArchiveFile
End Property

#End If ' ImplDecompress

'=========================================================================
' Methods
'=========================================================================

#If ImplCompress Then

Public Function AddFile( _
            File As Variant, _
            Optional Name As String, _
            Optional Comment As String, _
            Optional Password As String, _
            Optional EncrStrength As Long, _
            Optional Level As Long = -1) As Boolean
    Const FUNC_NAME     As String = "AddFile"
    Dim uFile           As ZipVfsType
    
    On Error GoTo EH
    pvSetError
    If m_lFileCount = 0 Then
        ReDim m_uFiles(0 To 2) As ZipFileInfo
    ElseIf m_lFileCount > UBound(m_uFiles) Then
        ReDim Preserve m_uFiles(0 To 2 * UBound(m_uFiles)) As ZipFileInfo
    End If
    uFile = pvVfsOpen(File)
    With m_uFiles(m_lFileCount)
        If LenB(Name) <> 0 Then
            .FileName = Name
        Else
            .FileName = Mid$(uFile.FileName, InStrRev(uFile.FileName, "\") + 1)
        End If
        If Right$(.FileName, 1) = "\" Then
            .Size = 0
            .Attributes = uFile.Data.dwFileAttributes Or vbDirectory
        Else
            .Size = pvToInt64(uFile.Data.nFileSizeLow, uFile.Data.nFileSizeHigh)
            .Attributes = uFile.Data.dwFileAttributes
        End If
        .Comment = Comment
        .LastModified = pvFromFileTime(uFile.Data.ftLastWriteTime)
        .Extra = vbNullString
        If IsObject(File) Then
            '--- remove VT_BYREF if any
            Set .SourceFile = C_Obj(File)
        Else
            .SourceFile = File
        End If
        .Level = Level
        #If ImplCrypto Then
            .Password = Password
            .Aes.Strength = EncrStrength
        #Else
            If LenB(Password) <> 0 Then
                On Error GoTo 0
                Err.Raise vbObjectError, , ERR_NO_CRYPTO_COMPILED
            End If
        #End If ' ImplCrypto
        If Not uFile.SourceArchive Is Nothing Then
            .FileName = uFile.SourceFileInfo(zipIdxFileName)
            .Comment = uFile.SourceFileInfo(zipIdxComment)
        End If
    End With
    m_lFileCount = m_lFileCount + 1
    '--- success
    AddFile = True
QH:
    On Error Resume Next
    pvVfsClose uFile
    Exit Function
EH:
    pvSetError MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description
    Resume QH
End Function

Public Function CompressArchive( _
            ArchiveFile As Variant, _
            Optional Comment As String, _
            Optional ByVal Level As Long = 6, _
            Optional ByVal UseUtf8 As VbTriState = vbUseDefault, _
            Optional ByVal UseZip64 As VbTriState = vbUseDefault) As Boolean
    Const FUNC_NAME     As String = "CompressArchive"
    Dim baBuffer()      As Byte
    Dim baNext()        As Byte
    Dim uArchiveFile    As ZipVfsType
    Dim uLocal          As ZipLocalHeaderType
    Dim lIdx            As Long
    Dim uFile           As ZipVfsType
    Dim lSize           As Long
    Dim cHeaderOffset   As Currency
    Dim cBeginOffset    As Currency
    Dim hCtx            As Long
    Dim uBuf            As ZipBuffersType
    Dim lResult         As Long
    Dim uCentral        As ZipCentralHeaderType
    Dim uEndHdr         As ZipEndHeaderType
    Dim lEntries        As Long
    Dim lComprLevel     As Long
    Dim bSkip           As Boolean
#If ImplCrypto Then
    Dim lSaltSize       As Long
    Dim baSalt()        As Byte
    Dim nPassVer        As Integer
    Dim lJdx            As Long
    Dim uAes            As ZipAesExtraDataType
#End If ' ImplCrypto
    
    On Error GoTo EH
    pvSetError
    m_lCurrentFile = -1
    ReDim baBuffer(0 To LNG_IOBUF_SIZE - 1) As Byte
    ReDim baNext(0 To LNG_IOBUF_SIZE - 1) As Byte
    uArchiveFile = pvVfsCreate(ArchiveFile)
    uCentral.VerExt = LNG_VER_DEFAULT
    For lIdx = 0 To m_lFileCount - 1
        m_lCurrentFile = lIdx
        With m_uFiles(lIdx)
            lComprLevel = IIf(.Level >= 0, .Level, Level)
            RaiseEvent BeforeCompress(lIdx, lComprLevel, bSkip, m_bCancel)
            If m_bCancel Then
                m_sLastError = ERR_USER_CANCEL
                GoTo QH
            End If
            If bSkip Then
                GoTo SkipFile
            End If
            .Offset = pvVfsSeek(uArchiveFile, 0, FILE_CURRENT)
            On Error GoTo EH_Continue
            uFile = pvVfsOpen(.SourceFile)
            If uFile.SourceArchive Is Nothing Then
                lComprLevel = Clamp(lComprLevel, 0, 9)
                uBuf.Greedy = (lComprLevel <= 4)
                uBuf.MaxMatch = At(Array(0, 2, 6, 12, 24, 8, 16, 32, 64, 1000), lComprLevel)
                uBuf.NiceLen = At(Array(0, 8, 10, 14, 24, 30, 65, 130, 200, 32768), lComprLevel)
                pvToDosDateTime .LastModified, uLocal.FDate, uLocal.FTime
                uLocal.Method = LNG_METHOD_DEFLATE
                uLocal.VerExt = LNG_VER_DEFLATE
                uLocal.Crc32 = -1
                uLocal.Ext.USize = 0
                uLocal.Ext.CSize = 0
                If UseUtf8 = vbUseDefault Then
                    uLocal.Flags = IIf(pvFromOemString(pvToOemString(.FileName, m_lCodePage), m_lCodePage) <> .FileName, zipFlagUseUtf8, 0)
                Else
                    uLocal.Flags = IIf(UseUtf8 = vbTrue, zipFlagUseUtf8, 0)
                End If
                #If ImplCrypto Then
                    If LenB(.Password) <> 0 Then
                        uLocal.Flags = uLocal.Flags Or zipFlagEncrypted
                        If .Aes.Strength = 0 Then
                            uLocal.Flags = uLocal.Flags Or zipFlagHasDataDescriptor
                            If Not pvCryptoTradInit(ToUtf8Array(.Password)) Then
                                Err.Raise vbObjectError, , m_sLastError
                            End If
                            '--- encrypt 12-byte random header w/ last byte used for password check
                            pvCryptoGetRandom LNG_ENC_HEADER_SIZE, baSalt
                            Call CopyMemory(baSalt(LNG_ENC_HEADER_SIZE - 1), ByVal UnsignedAdd(VarPtr(uLocal.FDate), -1), 1)
                            If Not pvCryptoTradCrypt(baSalt) Then
                                Err.Raise vbObjectError, , m_sLastError
                            End If
                        Else
                            lSaltSize = (.Aes.Strength + 1) * 4
                            pvCryptoGetRandom lSaltSize, baSalt
                            If Not pvCryptoAesInit(ToUtf8Array(.Password), baSalt, lSaltSize * 2, nPassVer) Then
                                Err.Raise vbObjectError, , m_sLastError
                            End If
                            '--- reserve space in extra data for AE-1/2
                            lJdx = UBound(.Extra) + 1
                            ReDim Preserve .Extra(0 To lJdx + LNG_EXTRADATA_AES_SIZE + 3) As Byte
                        End If
                    End If
                #End If ' ImplCrypto
                cHeaderOffset = pvVfsSeek(uArchiveFile, 0, FILE_CURRENT)
                pvOutputLocalHeader uArchiveFile, uLocal, .FileName, .Extra, .Size, UseZip64
                #If ImplCrypto Then
                    If LenB(.Password) <> 0 Then
                        '--- prepend salt/enc_header before file data
                        pvVfsWrite uArchiveFile, VarPtr(baSalt(0)), UBound(baSalt) + 1
                        If .Aes.Strength > 0 Then
                            '--- prepend Password Verification Value for AES
                            pvVfsWrite uArchiveFile, VarPtr(nPassVer), 2
                        End If
                    End If
                #End If ' ImplCrypto
                cBeginOffset = pvVfsSeek(uArchiveFile, 0, FILE_CURRENT)
                If (.Attributes And vbDirectory + vbVolume) = 0 Then
                    If lComprLevel > 0 Then
                        hCtx = CallWindowProc(m_uRtbl.CompressInit, VarPtr(m_uRtbl))
                        If hCtx = 0 Then
                            Err.Raise vbObjectError, , ERR_INIT_COMPRESSOR
                        End If
                        lSize = pvVfsRead(uFile, VarPtr(baNext(0)), LNG_IOBUF_SIZE)
                        Do
                            Call CopyMemory(baBuffer(0), baNext(0), lSize)
                            uBuf.InBlock = VarPtr(baBuffer(0))
                            uBuf.InLen = lSize
                            lSize = pvVfsRead(uFile, VarPtr(baNext(0)), LNG_IOBUF_SIZE)
                            uBuf.Final = (lSize = 0)
                            lResult = CallWindowProc(m_uRtbl.CompressBlock, hCtx, VarPtr(uBuf), VarPtr(uLocal.Crc32), lComprLevel) '--- level ignored
                            If lResult = 0 Or uBuf.OutBlock = 0 Then
                                Err.Raise vbObjectError, , ERR_COMPRESSING
                            End If
                            #If ImplCrypto Then
                                If LenB(.Password) <> 0 Then
                                    If UBound(baBuffer) < uBuf.OutLen Then
                                        ReDim baBuffer(0 To (uBuf.OutLen And -2048) + 2047) As Byte
                                    End If
                                    Call CopyMemory(baBuffer(0), ByVal uBuf.OutBlock, uBuf.OutLen)
                                    If .Aes.Strength = 0 Then
                                        If Not pvCryptoTradCrypt(baBuffer, Size:=uBuf.OutLen) Then
                                            Err.Raise vbObjectError, , m_sLastError
                                        End If
                                    Else
                                        If Not pvCryptoAesCrypt(baBuffer, Size:=uBuf.OutLen, HashAfter:=True) Then
                                            Err.Raise vbObjectError, , m_sLastError
                                        End If
                                    End If
                                    pvVfsWrite uArchiveFile, VarPtr(baBuffer(0)), uBuf.OutLen
                                    Call CoTaskMemFree(uBuf.OutBlock)
                                    uBuf.OutBlock = 0
                                End If
                            #End If ' ImplCrypto
                            If uBuf.OutBlock <> 0 Then
                                pvVfsWrite uArchiveFile, uBuf.OutBlock, uBuf.OutLen
                                Call CoTaskMemFree(uBuf.OutBlock)
                                uBuf.OutBlock = 0
                            End If
                            uLocal.Ext.USize = uLocal.Ext.USize + uBuf.InLen
                            uLocal.Ext.CSize = uLocal.Ext.CSize + uBuf.OutLen
                            .CompressedSize = uLocal.Ext.CSize
                            RaiseEvent Progress(lIdx, uLocal.Ext.USize, .Size, m_bCancel)
                            If m_bCancel Then
                                m_sLastError = ERR_USER_CANCEL
                                GoTo QH
                            End If
                        Loop While lSize <> 0
                        If hCtx <> 0 Then
                            Call CallWindowProc(m_uRtbl.CompressCleanup, hCtx)
                            hCtx = 0
                        End If
                        If uLocal.Ext.CSize = 0 And uLocal.Ext.USize = 0 Then
                            uLocal.Method = LNG_METHOD_STORE
                            uLocal.VerExt = LNG_VER_DEFAULT
                        End If
                    End If
                    '--- if data is incompressible and not encrypted -> store file
                    If lComprLevel = 0 Or (uLocal.Ext.CSize > uLocal.Ext.USize And (uLocal.Flags And zipFlagEncrypted) = 0) Then
                        pvVfsSeek uFile, 0, FILE_BEGIN
                        pvVfsSeek uArchiveFile, cBeginOffset, FILE_BEGIN
                        uLocal.Method = LNG_METHOD_STORE
                        uLocal.VerExt = LNG_VER_DEFAULT
                        uLocal.Crc32 = -1
                        uLocal.Ext.USize = 0
                        uLocal.Ext.CSize = 0
                        #If ImplCrypto Then
                            If LenB(.Password) <> 0 Then
                                If .Aes.Strength = 0 Then
                                    pvVfsSeek uArchiveFile, cHeaderOffset, FILE_BEGIN
                                    uLocal.Flags = uLocal.Flags And Not zipFlagEncrypted And Not zipFlagHasDataDescriptor
                                    pvOutputLocalHeader uArchiveFile, uLocal, .FileName, .Extra, .Size, UseZip64
                                    .Password = vbNullString
                                Else
                                    If Not pvCryptoAesInit(ToUtf8Array(.Password), baSalt, lSaltSize * 2, nPassVer) Then
                                        Err.Raise vbObjectError, , m_sLastError
                                    End If
                                End If
                            End If
                        #End If ' ImplCrypto
                        Do
                            RaiseEvent Progress(lIdx, uLocal.Ext.USize, .Size, m_bCancel)
                            If m_bCancel Then
                                m_sLastError = ERR_USER_CANCEL
                                GoTo QH
                            End If
                            lSize = pvVfsRead(uFile, VarPtr(baBuffer(0)), LNG_IOBUF_SIZE)
                            If lSize = 0 Then
                                Exit Do
                            End If
                            Call CallWindowProc(m_uRtbl.CalcCrc32, VarPtr(m_uRtbl), VarPtr(baBuffer(0)), lSize, VarPtr(uLocal.Crc32))
                            #If ImplCrypto Then
                                If LenB(.Password) <> 0 Then
                                    If .Aes.Strength = 0 Then
                                        If Not pvCryptoTradCrypt(baBuffer, Size:=lSize) Then
                                            Err.Raise vbObjectError, , m_sLastError
                                        End If
                                    Else
                                        If Not pvCryptoAesCrypt(baBuffer, Size:=lSize, HashAfter:=True) Then
                                            Err.Raise vbObjectError, , m_sLastError
                                        End If
                                    End If
                                End If
                            #End If ' ImplCrypto
                            pvVfsWrite uArchiveFile, VarPtr(baBuffer(0)), lSize
                            uLocal.Ext.USize = uLocal.Ext.USize + lSize
                            uLocal.Ext.CSize = uLocal.Ext.CSize + lSize
                        Loop
                    End If
                End If
                pvVfsClose uFile
                uLocal.Crc32 = uLocal.Crc32 Xor -1
                #If ImplCrypto Then
                    If LenB(.Password) <> 0 Then
                        If .Aes.Strength = 0 Then
                            uLocal.Ext.CSize = uLocal.Ext.CSize + LNG_ENC_HEADER_SIZE
                        Else
                            '--- append HMAC-SHA1 authentication tag after file data
                            baSalt = pvCryptoAesGetFinalHash(LNG_AES_AUTHCODE_SIZE)
                            pvVfsWrite uArchiveFile, VarPtr(baSalt(0)), LNG_AES_AUTHCODE_SIZE
                            '--- update AE-1/2 in extra data
                            uAes.HeaderId = LNG_EXTRADATA_AES_HEADER
                            uAes.DataSize = LNG_EXTRADATA_AES_SIZE
                            uAes.Version = IIf(.Size <= LNG_EXTRADATA_AE_2_MAXSIZE, 2, 1)
                            uAes.VendorId = LNG_EXTRADATA_AES_VENDOR
                            uAes.Strength = lSaltSize / 4 - 1
                            uAes.Method = uLocal.Method
                            lJdx = UBound(.Extra) - LNG_EXTRADATA_AES_SIZE - 3
                            Debug.Assert VarPtr(uAes.Strength) + LenB(uAes.Strength) - VarPtr(uAes.HeaderId) = 9
                            Call CopyMemory(.Extra(lJdx + 0), uAes.HeaderId, 9)
                            Call CopyMemory(.Extra(lJdx + 9), uAes.Method, LenB(uAes.Method))
                            '--- update local header
                            uLocal.Method = LNG_METHOD_AES
                            uLocal.Ext.CSize = uLocal.Ext.CSize + lSaltSize + 2 + LNG_AES_AUTHCODE_SIZE
                            If uAes.Version = 2 Then
                                uLocal.Crc32 = 0
                            End If
                        End If
                    End If
                #End If ' ImplCrypto
                cBeginOffset = pvVfsSeek(uArchiveFile, 0, FILE_CURRENT)
                pvVfsSeek uArchiveFile, .Offset, FILE_BEGIN
                pvOutputLocalHeader uArchiveFile, uLocal, .FileName, .Extra, .Size, UseZip64
                pvVfsSeek uArchiveFile, cBeginOffset, FILE_BEGIN
            Else
                .CompressedSize = uFile.SourceFileInfo(zipIdxCompressedSize)
                If Not uFile.SourceArchive.frCopyCompressed(uArchiveFile, uFile.SourceFileIdx, uLocal, Me, lIdx) Then
                    m_sLastError = uFile.SourceArchive.LastError
                    Err.Raise vbObjectError, , m_sLastError
                End If
                uLocal.Ext.USize = uFile.SourceFileInfo(zipIdxSize)
                uLocal.Ext.CSize = uFile.SourceFileInfo(zipIdxCompressedSize)
            End If
            .Crc32 = uLocal.Crc32
            Debug.Assert .Size = uLocal.Ext.USize
            .Size = uLocal.Ext.USize
            .CompressedSize = uLocal.Ext.CSize
            .Method = uLocal.Method
            .Flags = uLocal.Flags
            If uLocal.VerExt > uCentral.VerExt Then
                uCentral.VerExt = uLocal.VerExt
            End If
            If (uLocal.Flags And zipFlagHasDataDescriptor) <> 0 Then
                pvOutputDataDecriptor uArchiveFile, uLocal, UseZip64
                uLocal.Crc32 = 0
            End If
            If False Then
SkipFile:
                On Error GoTo EH
                pvVfsSeek uArchiveFile, .Offset, FILE_BEGIN
                .Offset = -1
                If hCtx <> 0 Then
                    Call CallWindowProc(m_uRtbl.CompressCleanup, hCtx)
                    hCtx = 0
                End If
            End If
            On Error GoTo EH
        End With
    Next
    m_lCurrentFile = -1
    cBeginOffset = pvVfsSeek(uArchiveFile, 0, FILE_CURRENT)
    uCentral.VerMade = LNG_VER_ZIP64
    For lIdx = 0 To m_lFileCount - 1
        m_lCurrentFile = lIdx
        With m_uFiles(lIdx)
            If .Offset >= 0 Then
                uCentral.Method = .Method
                pvToDosDateTime .LastModified, uCentral.FDate, uCentral.FTime
                uCentral.Crc32 = .Crc32
                uCentral.AttribX = .Attributes
                uCentral.Flags = .Flags
                uCentral.Ext.CSize = .CompressedSize
                uCentral.Ext.USize = .Size
                uCentral.Ext.Offset = .Offset
                pvOutputCentralHeader uArchiveFile, uCentral, .FileName, .Extra, .Comment, UseZip64
                lEntries = lEntries + 1
            End If
        End With
    Next
    m_lCurrentFile = -1
    If lEntries = 0 Then
        Err.Raise vbObjectError, , ERR_EMPTY_ARCHIVE
    End If
    uEndHdr.Ext.Entries = lEntries
    uEndHdr.Ext.TotEntr = lEntries
    uEndHdr.Ext.CenSize = pvVfsSeek(uArchiveFile, 0, FILE_CURRENT) - cBeginOffset
    uEndHdr.Ext.CenOff = cBeginOffset
    pvOutputEndHeader uArchiveFile, uEndHdr, Comment, UseZip64
    pvVfsSetEof uArchiveFile, "[output_archive]", Now
    '--- success
    CompressArchive = True
QH:
    On Error Resume Next
    If uBuf.OutBlock <> 0 Then
        Call CoTaskMemFree(uBuf.OutBlock)
        uBuf.OutBlock = 0
    End If
    If hCtx <> 0 Then
        Call CallWindowProc(m_uRtbl.CompressCleanup, hCtx)
        hCtx = 0
    End If
    pvVfsClose uFile
    pvVfsClose uArchiveFile
    Exit Function
EH:
    pvSetError MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description
    Resume QH
EH_Continue:
    If pvSetError(MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description, CanContinue:=True) Then
        Resume QH
    Else
        Resume SkipFile
    End If
End Function

Public Function AddFromFolder( _
            sFolderAndMask As String, _
            Optional Recursive As Boolean, _
            Optional TargetFolder As String, _
            Optional IncludeEmptyFolders As Boolean, _
            Optional Password As String, _
            Optional EncrStrength As Long, _
            Optional Level As Long = -1) As Boolean
    Const FUNC_NAME     As String = "AddFromFolder"
    Dim lIdx            As Long
    Dim sFolder         As String
    Dim sMask           As String
    Dim lRootOffset     As Long
    Dim cFolders        As Collection
    Dim vElem           As Variant
    Dim sName           As String
    
    On Error GoTo EH
    pvSetError
    Set cFolders = New Collection
    lIdx = InStrRev(sFolderAndMask, "\")
    If lIdx > 0 Then
        sFolder = Left$(sFolderAndMask, lIdx - 1)
    End If
    sMask = Mid$(sFolderAndMask, lIdx + 1)
    If LenB(sFolder) = 0 Then
        sFolder = CurDir$()
    End If
    lRootOffset = Len(sFolder)
    cFolders.Add sFolder
    Do While cFolders.Count > 0
        sFolder = cFolders.Item(1)
        cFolders.Remove 1
        lIdx = 0
        For Each vElem In pvEnumFiles(sFolder, sMask, -vbDirectory)
            If AddFile(vElem, PathCombine(TargetFolder, Mid$(vElem, lRootOffset + 2)), Password:=Password, EncrStrength:=EncrStrength, Level:=Level) Then
                lIdx = lIdx + 1
                '--- success (entries added)
                AddFromFolder = True
            End If
        Next
        If lIdx = 0 And IncludeEmptyFolders Then
            sName = PathCombine(TargetFolder, Mid$(sFolder, lRootOffset + 2))
            If LenB(sName) <> 0 Then
                If AddFile(sFolder & "\", sName & "\", Password:=Password, EncrStrength:=EncrStrength, Level:=Level) Then
                    '--- success (entries added)
                    AddFromFolder = True
                End If
            End If
        End If
        If Recursive Then
            For Each vElem In pvEnumFiles(sFolder, "*.*", vbDirectory)
                cFolders.Add vElem
            Next
        End If
    Loop
QH:
    Exit Function
EH:
    pvSetError MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description
    Resume QH
End Function

Private Function pvEnumFiles(sFolder As String, sMask As String, ByVal eAttrib As VbFileAttribute) As Collection
    Dim sFile           As String
    Dim hFind           As Long
    Dim uData           As WIN32_FIND_DATA
    
    On Error GoTo EH
    Set pvEnumFiles = New Collection
    sFile = PathCombine(sFolder, "*.*")
    hFind = FindFirstFile(StrPtr(sFile), VarPtr(uData))
    If hFind = INVALID_HANDLE_VALUE Then
        If Err.LastDllError <> 2 Then
            On Error GoTo 0
            Err.Raise vbObjectError, "FindFirstFile", GetSystemMessage(Err.LastDllError) & " (" & sFile & ")"
        End If
    Else
        Do
            If eAttrib > 0 And (uData.dwFileAttributes And eAttrib) <> 0 _
                    Or eAttrib < 0 And (uData.dwFileAttributes And -eAttrib) = 0 Then
                sFile = Left$(uData.cFileName, InStr(uData.cFileName, Chr$(0)) - 1)
                If sFile <> "." And sFile <> ".." Then
                    If PathMatchSpecW(StrPtr(sFile), StrPtr(sMask)) <> 0 Then
                        pvEnumFiles.Add PathCombine(sFolder, sFile)
                    End If
                End If
            End If
        Loop While FindNextFile(hFind, VarPtr(uData)) <> 0
        Call FindClose(hFind)
        hFind = 0
    End If
    Exit Function
EH:
    If hFind <> 0 And hFind <> INVALID_HANDLE_VALUE Then
        Call FindClose(hFind)
        hFind = 0
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub pvOutputLocalHeader(uFile As ZipVfsType, uHdr As ZipLocalHeaderType, FileName As String, baExtra() As Byte, ByVal cFileSize As Currency, ByVal UseZip64 As VbTriState)
    Dim baFName()       As Byte
    
    If UseZip64 = vbTrue Or uHdr.Ext.CSize > LNG_SIZE_LIMIT Or uHdr.Ext.USize > LNG_SIZE_LIMIT Or cFileSize > LNG_SIZE_LIMIT Then
        If UseZip64 = vbFalse Then
            Err.Raise vbObjectError, , ERR_MISSING_ZIP64
        End If
        UseZip64 = vbTrue
        uHdr.VerExt = LNG_VER_ZIP64
        uHdr.CSize = -1
        uHdr.USize = -1
    Else
        uHdr.CSize = pvToInt32(uHdr.Ext.CSize)
        uHdr.USize = pvToInt32(uHdr.Ext.USize)
    End If
    baFName = pvToOemString(Replace(FileName, "\", "/"), IIf((uHdr.Flags And zipFlagUseUtf8) <> 0, CP_UTF8, m_lCodePage))
    uHdr.Signature = LNG_SIG_LOCAL
    uHdr.LenFname = UBound(baFName) + 1
    uHdr.LenExt = IIf(UseZip64 = vbTrue, 4 + LNG_EXTRADATA_ZIP64_LOCAL_SIZE, 0) + UBound(baExtra) + 1
    Debug.Assert VarPtr(uHdr.FDate) + LenB(uHdr.FDate) - VarPtr(uHdr.Signature) + VarPtr(uHdr.LenExt) + 2 - VarPtr(uHdr.Crc32) = LNG_LEN_LOCAL
    Debug.Assert VarPtr(uHdr.FDate) + LenB(uHdr.FDate) - VarPtr(uHdr.Signature) = 14
    pvVfsWrite uFile, VarPtr(uHdr.Signature), 14
    Debug.Assert VarPtr(uHdr.LenExt) + LenB(uHdr.LenExt) - VarPtr(uHdr.Crc32) = 16
    pvVfsWrite uFile, VarPtr(uHdr.Crc32), 16
    Debug.Assert uHdr.LenFname > 0
    If uHdr.LenFname > 0 Then
        pvVfsWrite uFile, VarPtr(baFName(0)), uHdr.LenFname
    End If
    If UseZip64 = vbTrue Then
        uHdr.Ext.HeaderId = LNG_EXTRADATA_ZIP64_HEADER
        uHdr.Ext.DataSize = LNG_EXTRADATA_ZIP64_LOCAL_SIZE
        uHdr.Ext.USize = uHdr.Ext.USize / 10000@
        uHdr.Ext.CSize = uHdr.Ext.CSize / 10000@
        Debug.Assert VarPtr(uHdr.Ext.CSize) + LenB(uHdr.Ext.CSize) - VarPtr(uHdr.Ext.HeaderId) = 20
        pvVfsWrite uFile, VarPtr(uHdr.Ext.HeaderId), 20
        uHdr.Ext.USize = uHdr.Ext.USize * 10000@
        uHdr.Ext.CSize = uHdr.Ext.CSize * 10000@
    End If
    If UBound(baExtra) >= 0 Then
        pvVfsWrite uFile, VarPtr(baExtra(0)), UBound(baExtra) + 1
    End If
End Sub

Private Sub pvOutputDataDecriptor(uFile As ZipVfsType, uHdr As ZipLocalHeaderType, ByVal UseZip64 As VbTriState)
    pvVfsWrite uFile, VarPtr(LNG_SIG_DATADESC), 4
    pvVfsWrite uFile, VarPtr(uHdr.Crc32), 4
    If UseZip64 = vbTrue Or uHdr.Ext.USize > LNG_SIZE_LIMIT Or uHdr.Ext.CSize > LNG_SIZE_LIMIT Then
        If UseZip64 = vbFalse Then
            Err.Raise vbObjectError, , ERR_MISSING_ZIP64
        End If
        uHdr.VerExt = LNG_VER_ZIP64
        uHdr.Ext.USize = uHdr.Ext.USize / 10000@
        uHdr.Ext.CSize = uHdr.Ext.CSize / 10000@
        Debug.Assert VarPtr(uHdr.Ext.USize) + LenB(uHdr.Ext.USize) - VarPtr(uHdr.Ext.CSize) = 16
        pvVfsWrite uFile, VarPtr(uHdr.Ext.CSize), 16
        uHdr.Ext.USize = uHdr.Ext.USize * 10000@
        uHdr.Ext.CSize = uHdr.Ext.CSize * 10000@
    Else
        uHdr.CSize = pvToInt32(uHdr.Ext.CSize)
        uHdr.USize = pvToInt32(uHdr.Ext.USize)
        Debug.Assert VarPtr(uHdr.USize) + LenB(uHdr.USize) - VarPtr(uHdr.CSize) = 8
        pvVfsWrite uFile, VarPtr(uHdr.CSize), 8
        uHdr.CSize = 0
        uHdr.USize = 0
    End If
End Sub

Private Sub pvOutputCentralHeader(uFile As ZipVfsType, uHdr As ZipCentralHeaderType, FileName As String, baExtra() As Byte, Comment As String, ByVal UseZip64 As VbTriState)
    Dim lCodePage       As Long
    Dim baFName()       As Byte
    Dim baComment()     As Byte
    
    If UseZip64 = vbTrue Or uHdr.Ext.USize > LNG_SIZE_LIMIT Or uHdr.Ext.CSize > LNG_SIZE_LIMIT Or uHdr.Ext.Offset > LNG_SIZE_LIMIT Then
        If UseZip64 = vbFalse Then
            Err.Raise vbObjectError, , ERR_MISSING_ZIP64
        End If
        UseZip64 = vbTrue
        uHdr.VerExt = LNG_VER_ZIP64
        uHdr.Ext.DataSize = 0
    End If
    If uHdr.Ext.CSize > LNG_SIZE_LIMIT Then
        uHdr.CSize = -1
        uHdr.Ext.DataSize = uHdr.Ext.DataSize + 8
    Else
        uHdr.CSize = pvToInt32(uHdr.Ext.CSize)
    End If
    If uHdr.Ext.USize > LNG_SIZE_LIMIT Then
        uHdr.USize = -1
        uHdr.Ext.DataSize = uHdr.Ext.DataSize + 8
    Else
        uHdr.USize = pvToInt32(uHdr.Ext.USize)
    End If
    If uHdr.Ext.Offset > LNG_SIZE_LIMIT Then
        uHdr.Offset = -1
        uHdr.Ext.DataSize = uHdr.Ext.DataSize + 8
    Else
        uHdr.Offset = pvToInt32(uHdr.Ext.Offset)
    End If
    lCodePage = IIf((uHdr.Flags And zipFlagUseUtf8) <> 0, CP_UTF8, m_lCodePage)
    baFName = pvToOemString(Replace(FileName, "\", "/"), lCodePage)
    baComment = pvToOemString(Comment, lCodePage)
    uHdr.Signature = LNG_SIG_CENTRAL
    uHdr.LenFname = UBound(baFName) + 1
    uHdr.LenExt = IIf(UseZip64 = vbTrue, 4 + uHdr.Ext.DataSize, 0) + UBound(baExtra) + 1
    uHdr.LenCom = UBound(baComment) + 1
    Debug.Assert VarPtr(uHdr.FDate) + LenB(uHdr.FDate) - VarPtr(uHdr.Signature) + VarPtr(uHdr.AttribI) + 2 - VarPtr(uHdr.Crc32) + VarPtr(uHdr.Offset) + 4 - VarPtr(uHdr.AttribX) = LNG_LEN_CENTRAL
    Debug.Assert VarPtr(uHdr.FDate) + LenB(uHdr.FDate) - VarPtr(uHdr.Signature) = 16
    pvVfsWrite uFile, VarPtr(uHdr.Signature), 16
    Debug.Assert VarPtr(uHdr.AttribI) + LenB(uHdr.AttribI) - VarPtr(uHdr.Crc32) = 22
    pvVfsWrite uFile, VarPtr(uHdr.Crc32), 22
    Debug.Assert VarPtr(uHdr.Offset) + LenB(uHdr.Offset) - VarPtr(uHdr.AttribX) = 8
    pvVfsWrite uFile, VarPtr(uHdr.AttribX), 8
    Debug.Assert uHdr.LenFname > 0
    If uHdr.LenFname > 0 Then
        pvVfsWrite uFile, VarPtr(baFName(0)), uHdr.LenFname
    End If
    If UseZip64 = vbTrue Then
        uHdr.Ext.HeaderId = LNG_EXTRADATA_ZIP64_HEADER
        Debug.Assert VarPtr(uHdr.Ext.DataSize) + LenB(uHdr.Ext.DataSize) - VarPtr(uHdr.Ext.HeaderId) = 4
        pvVfsWrite uFile, VarPtr(uHdr.Ext.HeaderId), 4
        If uHdr.USize = -1 Then
            uHdr.Ext.USize = uHdr.Ext.USize / 10000@
            pvVfsWrite uFile, VarPtr(uHdr.Ext.USize), LenB(uHdr.Ext.USize)
            uHdr.Ext.USize = uHdr.Ext.USize * 10000@
        End If
        If uHdr.CSize = -1 Then
            uHdr.Ext.CSize = uHdr.Ext.CSize / 10000@
            pvVfsWrite uFile, VarPtr(uHdr.Ext.CSize), LenB(uHdr.Ext.CSize)
            uHdr.Ext.CSize = uHdr.Ext.CSize * 10000@
        End If
        If uHdr.Offset = -1 Then
            uHdr.Ext.Offset = uHdr.Ext.Offset / 10000@
            pvVfsWrite uFile, VarPtr(uHdr.Ext.Offset), LenB(uHdr.Ext.Offset)
            uHdr.Ext.Offset = uHdr.Ext.Offset * 10000@
        End If
    End If
    If UBound(baExtra) >= 0 Then
        pvVfsWrite uFile, VarPtr(baExtra(0)), UBound(baExtra) + 1
    End If
    If uHdr.LenCom > 0 Then
        pvVfsWrite uFile, VarPtr(baComment(0)), uHdr.LenCom
    End If
End Sub

Private Sub pvOutputEndHeader(uFile As ZipVfsType, uHdr As ZipEndHeaderType, Comment As String, ByVal UseZip64 As VbTriState)
    Dim baComment()     As Byte
    Dim uLocator        As Zip64EndLocatorType
    
    
    If UseZip64 = vbTrue Or uHdr.Ext.Entries > LNG_ENTRIES_LIMIT Or uHdr.Ext.TotEntr > LNG_ENTRIES_LIMIT Or uHdr.Ext.CenSize > LNG_SIZE_LIMIT Or uHdr.Ext.CenOff > LNG_SIZE_LIMIT Then
        If UseZip64 = vbFalse Then
            Err.Raise vbObjectError, , ERR_MISSING_ZIP64
        End If
        uHdr.CenOff = -1
        '--- Zip64 End of Central Directory record
        uHdr.Ext.Signature = LNG_SIG_EOCDR
        uHdr.Ext.Size = (LNG_LEN_EOCDR - 12) / 10000@
        uHdr.Ext.VerMade = LNG_VER_ZIP64
        uHdr.Ext.VerExt = LNG_VER_ZIP64
        uHdr.Ext.Entries = uHdr.Ext.Entries / 10000@
        uHdr.Ext.TotEntr = uHdr.Ext.TotEntr / 10000@
        uHdr.Ext.CenSize = uHdr.Ext.CenSize / 10000@
        uHdr.Ext.CenOff = uHdr.Ext.CenOff / 10000@
        uLocator.Offset = pvVfsSeek(uFile, 0, FILE_CURRENT) / 10000@
        Debug.Assert VarPtr(uHdr.Ext.CenOff) + 8 - VarPtr(uHdr.Ext.Signature) = LNG_LEN_EOCDR
        pvVfsWrite uFile, VarPtr(uHdr.Ext.Signature), LNG_LEN_EOCDR
        uHdr.Ext.Entries = uHdr.Ext.Entries * 10000@
        uHdr.Ext.TotEntr = uHdr.Ext.TotEntr * 10000@
        uHdr.Ext.CenSize = uHdr.Ext.CenSize * 10000@
        uHdr.Ext.CenOff = uHdr.Ext.CenOff * 10000@
        '--- Zip64 End of Central Directory Locator record
        uLocator.Signature = LNG_SIG_LOCATOR
        Debug.Assert VarPtr(uLocator.DiskNum) + 4 - VarPtr(uLocator.Signature) = LNG_LEN_LOCATOR
        pvVfsWrite uFile, VarPtr(uLocator.Signature), LNG_LEN_LOCATOR
    Else
        uHdr.CenOff = pvToInt32(uHdr.Ext.CenOff)
    End If
    If uHdr.Ext.Entries > LNG_ENTRIES_LIMIT Then
        uHdr.Entries = -1
    Else
        uHdr.Entries = pvToInt32(uHdr.Ext.Entries) And &HFFFF&
    End If
    If uHdr.Ext.TotEntr > LNG_ENTRIES_LIMIT Then
        uHdr.TotEntr = -1
    Else
        uHdr.TotEntr = pvToInt32(uHdr.Ext.TotEntr) And &HFFFF&
    End If
    If uHdr.Ext.CenSize > LNG_SIZE_LIMIT Then
        uHdr.CenSize = -1
    Else
        uHdr.CenSize = pvToInt32(uHdr.Ext.CenSize)
    End If
    baComment = pvToOemString(Comment, m_lCodePage)
    uHdr.Signature = LNG_SIG_END
    uHdr.LenCom = UBound(baComment) + 1
    Debug.Assert VarPtr(uHdr.LenCom) + 2 - VarPtr(uHdr.Signature) = LNG_LEN_END
    pvVfsWrite uFile, VarPtr(uHdr.Signature), LNG_LEN_END
    If uHdr.LenCom > 0 Then
        pvVfsWrite uFile, VarPtr(baComment(0)), uHdr.LenCom
    End If
End Sub

Private Sub pvToDosDateTime(dDate As Date, nDate As Integer, nTime As Integer)
    Dim uSysTime        As SYSTEMTIME
    Dim uFileTime       As FILETIME
    
    With uSysTime
        .wYear = Year(dDate)
        .wMonth = Month(dDate)
        .wDay = Day(dDate)
        .wHour = Hour(dDate)
        .wMinute = Minute(dDate)
        .wSecond = Second(dDate)
    End With
    Call SystemTimeToFileTime(uSysTime, uFileTime)
    Call FileTimeToDosDateTime(uFileTime, VarPtr(nDate), VarPtr(nTime))
End Sub

Private Function pvToOemString(sText As String, ByVal lCodePage As Long) As Byte()
    Dim baRetVal()      As Byte
    Dim lSize           As Long
    
    If Len(sText) <> 0 Then
        lSize = WideCharToMultiByte(lCodePage, 0, StrPtr(sText), Len(sText), 0, 0, 0, 0)
        If lSize > 0 Then
            ReDim baRetVal(0 To lSize - 1) As Byte
            lSize = WideCharToMultiByte(lCodePage, 0, StrPtr(sText), Len(sText), VarPtr(baRetVal(0)), UBound(baRetVal) + 1, 0, 0)
            pvToOemString = baRetVal
        Else
            pvToOemString = vbNullString
        End If
    Else
        pvToOemString = vbNullString
    End If
End Function

#End If ' ImplCompress

#If ImplDecompress Then

Public Function OpenArchive(ArchiveFile As Variant) As Boolean
    Const FUNC_NAME     As String = "OpenArchive"
    Const MAX_END_SEEK  As Long = 10000
    Dim uArchiveFile    As ZipVfsType
    Dim uEndHdr         As ZipEndHeaderType
    Dim lIdx            As Long
    Dim sArchiveComment As String
    Dim uFiles()        As ZipFileInfo
    Dim uCentral        As ZipCentralHeaderType
    Dim baComment()     As Byte
    Dim cOffset         As Currency
    Dim baBuffer()      As Byte
    Dim uLocator        As Zip64EndLocatorType
    
    On Error GoTo EH
    pvSetError
    m_lCurrentFile = -1
    uArchiveFile = pvVfsOpen(ArchiveFile)
    cOffset = pvVfsSeek(uArchiveFile, 0, FILE_END)
    If cOffset >= LNG_LEN_END Then
        For lIdx = 0 To IIf(cOffset - LNG_LEN_END < MAX_END_SEEK, cOffset - LNG_LEN_END, MAX_END_SEEK)
            cOffset = pvVfsSeek(uArchiveFile, -LNG_LEN_END - lIdx, FILE_END)
            Debug.Assert VarPtr(uEndHdr.LenCom) + 2 - VarPtr(uEndHdr.Signature) = LNG_LEN_END
            pvVfsRead uArchiveFile, VarPtr(uEndHdr), LNG_LEN_END
            If uEndHdr.Signature = LNG_SIG_END And uEndHdr.LenCom = lIdx Then
                If uEndHdr.LenCom > 0 Then
                    ReDim baComment(0 To uEndHdr.LenCom - 1) As Byte
                    pvVfsRead uArchiveFile, VarPtr(baComment(0)), uEndHdr.LenCom
                    sArchiveComment = pvFromOemString(baComment, m_lCodePage)
                End If
                If uEndHdr.CenOff = -1 Or uEndHdr.Ext.Entries = -1 Then
                    If cOffset <= LNG_LEN_LOCATOR + LNG_LEN_EOCDR Then
                        Exit For
                    End If
                    pvVfsSeek uArchiveFile, cOffset - LNG_LEN_LOCATOR, FILE_BEGIN
                    pvVfsRead uArchiveFile, VarPtr(uLocator), LNG_LEN_LOCATOR
                    If uLocator.Signature <> LNG_SIG_LOCATOR Or uLocator.DiskEocdr <> uEndHdr.DiskNum Then
                        Exit For
                    End If
                    pvVfsSeek uArchiveFile, uLocator.Offset * 10000@, FILE_BEGIN
                    pvVfsRead uArchiveFile, VarPtr(uEndHdr.Ext), LNG_LEN_EOCDR
                    If uEndHdr.Ext.Signature <> LNG_SIG_EOCDR Then
                        Exit For
                    End If
                    uEndHdr.Ext.CenOff = uEndHdr.Ext.CenOff * 10000@
                    uEndHdr.Ext.Entries = uEndHdr.Ext.Entries * 10000@
                Else
                    uEndHdr.Ext.CenOff = pvToInt64(uEndHdr.CenOff)
                    uEndHdr.Ext.Entries = uEndHdr.Entries And &HFFFF&
                End If
                lIdx = -1
                Exit For
            End If
        Next
    End If
    If lIdx < 0 Then
        '--- note: redim one more (last ignored)
        ReDim uFiles(0 To uEndHdr.Ext.Entries) As ZipFileInfo
        pvVfsSeek uArchiveFile, uEndHdr.Ext.CenOff, FILE_BEGIN
        For lIdx = 0 To uEndHdr.Ext.Entries - 1
            m_lCurrentFile = lIdx
            Debug.Assert VarPtr(uCentral.FDate) + LenB(uCentral.FDate) - VarPtr(uCentral.Signature) + VarPtr(uCentral.AttribI) + 2 - VarPtr(uCentral.Crc32) + VarPtr(uCentral.Offset) + 4 - VarPtr(uCentral.AttribX) = LNG_LEN_CENTRAL
            Debug.Assert VarPtr(uCentral.FDate) + LenB(uCentral.FDate) - VarPtr(uCentral.Signature) = 16
            pvVfsRead uArchiveFile, VarPtr(uCentral.Signature), 16
            Debug.Assert VarPtr(uCentral.AttribI) + LenB(uCentral.AttribI) - VarPtr(uCentral.Crc32) = 22
            pvVfsRead uArchiveFile, VarPtr(uCentral.Crc32), 22
            Debug.Assert VarPtr(uCentral.Offset) + LenB(uCentral.Offset) - VarPtr(uCentral.AttribX) = 8
            pvVfsRead uArchiveFile, VarPtr(uCentral.AttribX), 8
            If uCentral.Signature <> LNG_SIG_CENTRAL Then
                pvSetError MODULE_NAME & "." & FUNC_NAME, ERR_INVALID_ARCHIVE & ". " & Replace(ERR_ENTRY_INVALID_SIG, "%1", lIdx + 1)
                GoTo QH
            End If
            If Not pvLoadFileInfo(uArchiveFile, uCentral, uFiles(lIdx), lIdx + 1) Then
                GoTo QH
            End If
        Next
    Else
        '--- fall-back to sequential lookup from start of file
        pvVfsSeek uArchiveFile, 0, FILE_BEGIN
        uEndHdr.Ext.Entries = 0
        Do
            cOffset = pvVfsSeek(uArchiveFile, 0, FILE_CURRENT)
            If pvVfsRead(uArchiveFile, VarPtr(uCentral.Signature), 4) < 4 Then
                Exit Do
            End If
            Select Case uCentral.Signature
            Case LNG_SIG_LOCAL
                uCentral.Offset = pvFromInt64(cOffset)
                pvVfsRead uArchiveFile, VarPtr(uCentral.VerExt), 10
                pvVfsRead uArchiveFile, VarPtr(uCentral.Crc32), 16
                uEndHdr.Ext.Entries = uEndHdr.Ext.Entries + 1
                If uEndHdr.Ext.Entries = 1 Then
                    ReDim uFiles(0 To 3) As ZipFileInfo
                ElseIf uEndHdr.Ext.Entries > UBound(uFiles) Then
                    ReDim Preserve uFiles(0 To 2 * uEndHdr.Ext.Entries) As ZipFileInfo
                End If
                If Not pvLoadFileInfo(uArchiveFile, uCentral, uFiles(uEndHdr.Ext.Entries - 1), uEndHdr.Ext.Entries) Then
                    GoTo QH
                End If
                If uCentral.Ext.CSize = 0 And uCentral.Ext.USize <> 0 Then
                    cOffset = pvVfsSeek(uArchiveFile, 0, FILE_CURRENT)
                    ReDim baBuffer(0 To uCentral.Ext.USize - LNG_LEN_EXTLOCAL) As Byte
                    pvVfsRead uArchiveFile, VarPtr(baBuffer(0)), UBound(baBuffer) + 1
                    For lIdx = 0 To UBound(baBuffer) - LNG_LEN_EXTLOCAL
                        If Peek(VarPtr(baBuffer(lIdx))) = LNG_SIG_EXTLOCAL Then
                            If Peek(VarPtr(baBuffer(lIdx + 4))) = uCentral.Crc32 Or uCentral.Crc32 = 0 Then
                                uCentral.Crc32 = Peek(VarPtr(baBuffer(lIdx + 4)))
                                uCentral.Ext.CSize = Peek(VarPtr(baBuffer(lIdx + 8)))
                                uFiles(uEndHdr.Ext.Entries - 1).Crc32 = uCentral.Crc32
                                uFiles(uEndHdr.Ext.Entries - 1).CompressedSize = uCentral.Ext.CSize
                                Exit For
                            End If
                        End If
                    Next
                    If uCentral.Ext.CSize = 0 Then
                        Exit Do
                    End If
                    pvVfsSeek uArchiveFile, cOffset, FILE_BEGIN
                End If
                pvVfsSeek uArchiveFile, uCentral.Ext.CSize, FILE_CURRENT
            Case LNG_SIG_EXTLOCAL
                If uEndHdr.Ext.Entries > 0 Then
                    pvVfsSeek uArchiveFile, LNG_LEN_EXTLOCAL - 4, FILE_CURRENT
                End If
            Case Else
                pvVfsSeek uArchiveFile, cOffset + 1, FILE_BEGIN
            End Select
        Loop
        If uEndHdr.Ext.Entries = 0 Then
            pvSetError MODULE_NAME & "." & FUNC_NAME, ERR_INVALID_ARCHIVE
        End If
        '--- note: redim one more (last ignored)
        ReDim Preserve uFiles(0 To uEndHdr.Ext.Entries) As ZipFileInfo
    End If
    '--- commit archive info to member vars
    AssignVariant m_vArchiveFile, ArchiveFile
    m_lFileCount = uEndHdr.Ext.Entries
    m_sComment = sArchiveComment
    m_uFiles = uFiles
    '--- success
    OpenArchive = True
QH:
    pvVfsClose uArchiveFile
    Exit Function
EH:
    pvSetError MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description
    Resume QH
End Function

Private Function pvLoadFileInfo(uArchiveFile As ZipVfsType, uCentral As ZipCentralHeaderType, uFile As ZipFileInfo, ByVal lEntry As Long) As Boolean
    Const FUNC_NAME     As String = "pvLoadFileInfo"
    Dim baFName()       As Byte
    Dim baExtra()       As Byte
    Dim baComment()     As Byte
    Dim lCodePage       As Long
    Dim lPos            As Long
    Dim lIdx            As Long
    Dim aSizes(0 To 2)  As Currency
    
    With uFile
        .Method = uCentral.Method
        .LastModified = pvFromDosDateTime(uCentral.FDate, uCentral.FTime)
        .Crc32 = uCentral.Crc32
        .Attributes = uCentral.AttribX
        .Flags = uCentral.Flags
        lCodePage = IIf((.Flags And zipFlagUseUtf8) <> 0, CP_UTF8, m_lCodePage)
        Debug.Assert uCentral.LenFname > 0
        If uCentral.LenFname > 0 Then
            ReDim baFName(0 To uCentral.LenFname - 1) As Byte
            pvVfsRead uArchiveFile, VarPtr(baFName(0)), uCentral.LenFname
            .FileName = pvToWinFileName(baFName, lCodePage)
        Else
            If pvSetError(MODULE_NAME & "." & FUNC_NAME, ERR_INVALID_ARCHIVE & ". " & Replace(ERR_ENTRY_NO_FILENAME, "%1", lEntry), CanContinue:=True) Then
                GoTo QH
            End If
            .FileName = Replace(STR_ENTRY, "%1", lEntry)
        End If
        uCentral.Ext.USize = pvToInt64(uCentral.USize)
        uCentral.Ext.CSize = pvToInt64(uCentral.CSize)
        uCentral.Ext.Offset = pvToInt64(uCentral.Offset)
        If uCentral.LenExt > 0 Then
            ReDim baExtra(0 To uCentral.LenExt - 1) As Byte
            pvVfsRead uArchiveFile, VarPtr(baExtra(0)), uCentral.LenExt
            .Extra = baExtra
            lPos = 0
            Do While lPos + 3 < uCentral.LenExt
                Call CopyMemory(uCentral.Ext.HeaderId, baExtra(lPos + 0), 4)
                Select Case uCentral.Ext.HeaderId
                Case LNG_EXTRADATA_ZIP64_HEADER
                    If uCentral.Ext.DataSize >= 8 And lPos + uCentral.Ext.DataSize + 3 <= UBound(baExtra) Then
                        Call CopyMemory(aSizes(0), baExtra(lPos + 4), Clamp(uCentral.Ext.DataSize, 0, 24))
                        If uCentral.USize = -1 Then
                            uCentral.Ext.USize = aSizes(lIdx) * 10000@
                            lIdx = lIdx + 1
                        End If
                        If uCentral.CSize = -1 Then
                            uCentral.Ext.CSize = aSizes(lIdx) * 10000@
                            lIdx = lIdx + 1
                        End If
                        If uCentral.Offset = -1 Then
                            uCentral.Ext.Offset = aSizes(lIdx) * 10000@
                            lIdx = lIdx + 1
                        End If
                    End If
                    Exit Do
                Case LNG_EXTRADATA_INFOZIP_UNICODE_PATH
                    If uCentral.Ext.DataSize > 5 And lPos + uCentral.Ext.DataSize + 3 <= UBound(baExtra) Then
                        If baExtra(lPos + 4) = 1 Then
                            ReDim baFName(0 To uCentral.Ext.DataSize - 6) As Byte
                            Call CopyMemory(baFName(0), baExtra(lPos + 9), uCentral.Ext.DataSize - 5)
                            .FileName = pvToWinFileName(baFName, lCodePage)
                        End If
                    End If
                End Select
                lPos = lPos + uCentral.Ext.DataSize + 4
            Loop
        End If
        .CompressedSize = uCentral.Ext.CSize
        .Size = uCentral.Ext.USize
        .Offset = uCentral.Ext.Offset
        If uCentral.LenCom > 0 Then
            ReDim baComment(0 To uCentral.LenCom - 1) As Byte
            pvVfsRead uArchiveFile, VarPtr(baComment(0)), uCentral.LenCom
            .Comment = pvFromOemString(baComment, lCodePage)
        End If
        #If ImplCrypto Then
            If (.Flags And zipFlagEncrypted) <> 0 Then
                If .Method = LNG_METHOD_AES Then
                    lPos = 0
                    Do While lPos + 3 < UBound(baExtra) + 1
                        Call CopyMemory(.Aes.HeaderId, baExtra(lPos + 0), 4)
                        If .Aes.HeaderId = LNG_EXTRADATA_AES_HEADER Then
                            If .Aes.DataSize >= LNG_EXTRADATA_AES_SIZE Then
                                Debug.Assert VarPtr(.Aes.Strength) + 1 - VarPtr(.Aes.HeaderId) = 9
                                Call CopyMemory(.Aes.HeaderId, baExtra(lPos + 0), 9)
                                Call CopyMemory(.Aes.Method, baExtra(lPos + 9), 2)
                                .Method = .Aes.Method
                            End If
                            Exit Do
                        End If
                        lPos = lPos + .Aes.DataSize + 4
                    Loop
                End If
                '--- store for traditional ZipCrypt password check
                If (.Flags And zipFlagHasDataDescriptor) <> 0 Then
                    Call CopyMemory(.DecDat, ByVal UnsignedAdd(VarPtr(uCentral.FDate), -1), 1)
                Else
                    Call CopyMemory(.DecDat, ByVal UnsignedAdd(VarPtr(uCentral.CSize), -1), 1)
                End If
            End If
        #End If ' ImplCrypto
    End With
    '--- success
    pvLoadFileInfo = True
QH:
End Function

Public Function Extract( _
            OutputTarget As Variant, _
            Optional Filter As Variant, _
            Optional Password As String) As Boolean
    Const FUNC_NAME     As String = "Extract"
    Dim baBuffer()      As Byte
    Dim uArchiveFile    As ZipVfsType
    Dim lIdx            As Long
    Dim uLocal          As ZipLocalHeaderType
    Dim baFName()       As Byte
    Dim baExt()         As Byte
    Dim cRead           As Currency
    Dim lSize           As Long
    Dim lResult         As Long
    Dim hCtx            As Long
    Dim uBuf            As ZipBuffersType
    Dim vFileName       As Variant
    Dim uFile           As ZipVfsType
    Dim lCrc32          As Long
    Dim pCrc32          As Long
    Dim bSkip           As Boolean
    Dim sFilterFolder   As String
    Dim sFilterMask     As String
    Dim sOutputFolder   As String
    Dim sOutputFile     As String
    Dim cMatched        As Collection
    Dim sFileName       As String
    Dim vElem           As Variant
#If ImplCrypto Then
    Dim lSaltSize       As Long
    Dim baSalt()        As Byte
    Dim nPassVer        As Integer
#End If ' ImplCrypto
    
    On Error GoTo EH
    pvSetError
    m_lCurrentFile = -1
    If IsEmpty(m_vArchiveFile) Then
        Err.Raise vbObjectError, , ERR_ARCHIVE_NOT_OPEN
    End If
    ReDim baBuffer(0 To LNG_IOBUF_SIZE - 1) As Byte
    uArchiveFile = pvVfsOpen(m_vArchiveFile)
    If VarType(Filter) = vbString Then
        sFilterMask = Replace(Filter, "/", "\")
        lIdx = InStrRev(sFilterMask, "\")
        If lIdx > 0 Then
            sFilterFolder = Left$(sFilterMask, lIdx) & "*"
            sFilterMask = Mid$(sFilterMask, lIdx + 1)
        End If
    End If
    Set cMatched = New Collection
    For lIdx = 0 To m_lFileCount - 1
        If IsMissing(Filter) Then
            '--- do nothing
        ElseIf IsArray(Filter) Then
            If Not At(Filter, lIdx) Then
                GoTo SkipMatch
            End If
        ElseIf VarType(Filter) = vbString Then
            sFileName = m_uFiles(lIdx).FileName
            If LenB(sFilterFolder) <> 0 Then
                If Not sFileName Like sFilterFolder Then
                    GoTo SkipMatch
                End If
                If LenB(sFilterMask) <> 0 Then
                    If sFileName Like sFilterFolder & "\*" Then
                        GoTo SkipMatch
                    End If
                End If
            End If
            If LenB(sFilterMask) <> 0 Then
                If PathMatchSpecW(StrPtr(Mid$(sFileName, InStrRev(sFileName, "\") + 1)), StrPtr(sFilterMask)) = 0 Then
                    GoTo SkipMatch
                End If
            End If
        ElseIf Not lIdx = Filter Then
            GoTo SkipMatch
        End If
        cMatched.Add lIdx
SkipMatch:
    Next
    If VarType(OutputTarget) = vbString Then
        sOutputFolder = OutputTarget
        If LenB(sOutputFolder) <> 0 And cMatched.Count = 1 Then
            If (FileAttribs(sOutputFolder) And vbDirectory) = 0 Then
                sOutputFolder = pvPathGetFolder(sOutputFolder, sOutputFile)
            End If
        End If
    End If
    '--- touch to open for read
    pvVfsRead uArchiveFile, 0, 0
    For Each vElem In cMatched
        m_lCurrentFile = vElem
        With m_uFiles(m_lCurrentFile)
            On Error GoTo EH_Continue
            '--- read local header
            pvVfsSeek uArchiveFile, .Offset, FILE_BEGIN
            Debug.Assert VarPtr(uLocal.FDate) + 2 - VarPtr(uLocal.Signature) + VarPtr(uLocal.LenExt) + 2 - VarPtr(uLocal.Crc32) = LNG_LEN_LOCAL
            Debug.Assert VarPtr(uLocal.FDate) + 2 - VarPtr(uLocal.Signature) = 14
            pvVfsRead uArchiveFile, VarPtr(uLocal.Signature), 14
            Debug.Assert VarPtr(uLocal.LenExt) + 2 - VarPtr(uLocal.Crc32) = 16
            pvVfsRead uArchiveFile, VarPtr(uLocal.Crc32), 16
            Debug.Assert uLocal.LenFname > 0
            If uLocal.LenFname > 0 Then
                ReDim baFName(0 To uLocal.LenFname - 1) As Byte
                pvVfsRead uArchiveFile, VarPtr(baFName(0)), uLocal.LenFname
            End If
            If uLocal.LenExt > 0 Then
                ReDim baExt(0 To uLocal.LenExt - 1) As Byte
                pvVfsRead uArchiveFile, VarPtr(baExt(0)), uLocal.LenExt
            End If
            '--- sanity checks
            If uLocal.Signature <> LNG_SIG_LOCAL Then
                Err.Raise vbObjectError, , ERR_INVALID_LOCALHDR & ". " & Replace(ERR_ENTRY_INVALID_SIG, "%1", m_lCurrentFile + 1)
            End If
'            If pvToWinFileName(baFName, .Flags) <> .FileName Then
'                Err.Raise vbObjectError, , ERR_INVALID_LOCALHDR & " (Filename)"
'            End If
            If .Method <> LNG_METHOD_DEFLATE And .Method <> LNG_METHOD_STORE Then
                If pvSetError(MODULE_NAME & "." & FUNC_NAME, Replace(ERR_UNSUPPORTED_METHOD, "%1", .Method), CanContinue:=True) Then
                    GoTo QH
                Else
                    .Method = LNG_METHOD_STORE
                End If
            End If
            bSkip = False
            If VarType(OutputTarget) = vbString Then
                If LenB(sOutputFile) = 0 Then
                    vFileName = .FileName
                ElseIf InStr(sOutputFile, "*.*") > 0 Then
                    vFileName = Replace(sOutputFile, "*.*", .FileName)
                Else
                    vFileName = Replace(sOutputFile, "*", pvFileGetFName(.FileName))
                End If
                vFileName = PathCombine(sOutputFolder, CStr(vFileName))
                RaiseEvent BeforeExtract(m_lCurrentFile, vFileName, bSkip, m_bCancel)
            Else
                RaiseEvent BeforeExtract(m_lCurrentFile, OutputTarget, bSkip, m_bCancel)
            End If
            If m_bCancel Then
                m_sLastError = ERR_USER_CANCEL
                GoTo QH
            End If
            If bSkip Then
                GoTo SkipFile
            End If
            '--- allow output to buffer, stream object or (customized) filename
            If VarType(OutputTarget) = vbString Then
                uFile = pvVfsCreate(vFileName, .Attributes)
            Else
                If IsArray(OutputTarget) And .Size > 0 Then
                    ReDim OutputTarget(0 To .Size - 1) As Byte
                End If
                uFile = pvVfsCreate(OutputTarget)
            End If
            If .CompressedSize > 0 Then
                lCrc32 = -1
                pCrc32 = VarPtr(lCrc32)
                cRead = 0
                If (.Flags And zipFlagEncrypted) <> 0 Then
                    RaiseEvent Progress(m_lCurrentFile, cRead, .CompressedSize, m_bCancel)
                    If m_bCancel Then
                        m_sLastError = ERR_USER_CANCEL
                        GoTo QH
                    End If
                    #If ImplCrypto Then
                        If (.Flags And zipFlagStrongEncrypted) <> 0 Then
                            Err.Raise vbObjectError, , ERR_UNSUPPORTED_ENCRYPTION
                        End If
                        If LenB(Password) = 0 Then
                            Err.Raise vbObjectError, , ERR_PASSWORD_REQUIRED
                        End If
                        If .Aes.Strength = 0 Then
                            '--- init traditional ZipCrypt
                            If Not pvCryptoTradInit(ToUtf8Array(Password)) Then
                                Err.Raise vbObjectError, , m_sLastError
                            End If
                            lResult = pvVfsRead(uArchiveFile, VarPtr(baBuffer(0)), LNG_ENC_HEADER_SIZE)
                            If lResult = 0 Then
                                Err.Raise vbObjectError, , ERR_READING_ARCHIVE
                            End If
                            If Not pvCryptoTradCrypt(baBuffer, 0, LNG_ENC_HEADER_SIZE, UseAfter:=True) Then
                                Err.Raise vbObjectError, , m_sLastError
                            End If
                            If baBuffer(LNG_ENC_HEADER_SIZE - 1) <> .DecDat Then
                                Err.Raise vbObjectError, , ERR_INVALID_PASSWORD
                            End If
                            cRead = LNG_ENC_HEADER_SIZE
                        Else
                            '--- init WinZip AES
                            lSaltSize = (.Aes.Strength + 1) * 4
                            lResult = pvVfsRead(uArchiveFile, VarPtr(baBuffer(0)), lSaltSize + 2)
                            If lResult < lSaltSize + 2 Then
                                Err.Raise vbObjectError, , ERR_READING_ARCHIVE
                            End If
                            ReDim baSalt(0 To lSaltSize - 1) As Byte
                            Call CopyMemory(baSalt(0), baBuffer(0), UBound(baSalt) + 1)
                            If Not pvCryptoAesInit(ToUtf8Array(Password), baSalt, lSaltSize * 2, nPassVer) Then
                                Err.Raise vbObjectError, , m_sLastError
                            End If
                            If nPassVer <> PeekInt(VarPtr(baBuffer(lSaltSize))) Then
                                Err.Raise vbObjectError, , ERR_INVALID_PASSWORD
                            End If
                            '--- don't calc CRC32 for AE-2
                            If .Aes.Version = 2 Then
                                pCrc32 = 0
                            End If
                            '--- note: LNG_AES_AUTHCODE_SIZE chunk comes after file data
                            cRead = lSaltSize + 2 + LNG_AES_AUTHCODE_SIZE
                        End If
                    #Else
                        Err.Raise vbObjectError, , ERR_NO_CRYPTO_COMPILED
                    #End If ' ImplCrypto
                End If
                If .Method = LNG_METHOD_DEFLATE Then
                    hCtx = CallWindowProc(m_uRtbl.DecompressInit, VarPtr(m_uRtbl))
                    If hCtx = 0 Then
                        Err.Raise vbObjectError, , ERR_INIT_DECOMPRESSOR
                    End If
                End If
                Do ' While cRead < .CompressedSize
                    RaiseEvent Progress(m_lCurrentFile, cRead, .CompressedSize, m_bCancel)
                    If m_bCancel Then
                        m_sLastError = ERR_USER_CANCEL
                        GoTo QH
                    End If
                    If cRead >= .CompressedSize Then
                        Exit Do
                    End If
                    lSize = LNG_IOBUF_SIZE
                    If lSize > .CompressedSize - cRead Then
                        lSize = .CompressedSize - cRead
                    End If
                    lResult = pvVfsRead(uArchiveFile, VarPtr(baBuffer(0)), lSize)
                    If lResult = 0 Then
                        Err.Raise vbObjectError, , ERR_READING_ARCHIVE
                    End If
                    #If ImplCrypto Then
                        If (.Flags And zipFlagEncrypted) <> 0 Then
                            If .Aes.Strength = 0 Then
                                If Not pvCryptoTradCrypt(baBuffer, 0, lResult, UseAfter:=True) Then
                                    Err.Raise vbObjectError, , m_sLastError
                                End If
                            Else
                                If Not pvCryptoAesCrypt(baBuffer, 0, lResult, HashBefore:=True) Then
                                    Err.Raise vbObjectError, , m_sLastError
                                End If
                            End If
                        End If
                    #End If ' ImplCrypto
                    cRead = cRead + lResult
                    If .Method = LNG_METHOD_DEFLATE Then
                        uBuf.InBlock = VarPtr(baBuffer(0))
                        uBuf.InLen = lResult
                        lResult = CallWindowProc(m_uRtbl.DecompressBlock, hCtx, VarPtr(uBuf), pCrc32)
                        If lResult = 0 Or uBuf.OutBlock = 0 Then
                            Err.Raise vbObjectError, , ERR_DECOMPRESSING
                        End If
                        pvVfsWrite uFile, uBuf.OutBlock, uBuf.OutLen
                        Call CoTaskMemFree(uBuf.OutBlock)
                        uBuf.OutBlock = 0
                    Else ' If .Method = LNG_METHOD_STORE Then
                        If pCrc32 <> 0 Then
                            Call CallWindowProc(m_uRtbl.CalcCrc32, VarPtr(m_uRtbl), VarPtr(baBuffer(0)), lResult, pCrc32)
                        End If
                        pvVfsWrite uFile, VarPtr(baBuffer(0)), lResult
                    End If
                Loop
                pvVfsSetEof uFile, .FileName, .LastModified
                #If ImplCrypto Then
                    If .Aes.Strength <> 0 Then
                        '--- check HMAC-SHA1 auth code
                        ReDim baSalt(0 To LNG_AES_AUTHCODE_SIZE - 1) As Byte
                        lResult = pvVfsRead(uArchiveFile, VarPtr(baSalt(0)), LNG_AES_AUTHCODE_SIZE)
                        If lResult < LNG_AES_AUTHCODE_SIZE Then
                            Err.Raise vbObjectError, , ERR_READING_ARCHIVE
                        End If
                        If Not pvArrayEqual(baSalt, pvCryptoAesGetFinalHash(LNG_AES_AUTHCODE_SIZE)) Then
                            Err.Raise vbObjectError, , ERR_INVALID_AUTHCODE
                        End If
                    End If
                #End If ' ImplCrypto
                If pCrc32 <> 0 Then
                    If .Crc32 <> (lCrc32 Xor -1) Then
                        Err.Raise vbObjectError, , ERR_CRC_CHECK
                    End If
                End If
            Else
                If (.Attributes And vbDirectory + vbVolume) = 0 And Right$(.FileName, 1) <> "\" Then
                    pvVfsSetEof uFile, .FileName, .LastModified
                End If
                '--- note: Total=1 to prevent division by zero in client code
                RaiseEvent Progress(m_lCurrentFile, 1, 1, m_bCancel)
                If m_bCancel Then
                    m_sLastError = ERR_USER_CANCEL
                    GoTo QH
                End If
            End If
            If VarType(OutputTarget) = vbString Then
                RaiseEvent ExtractComplete(m_lCurrentFile, vFileName)
            Else
                RaiseEvent ExtractComplete(m_lCurrentFile, OutputTarget)
            End If
            '--- success (entries extracted)
            Extract = True
SkipFile:
            On Error GoTo EH
            pvVfsClose uFile
            If hCtx <> 0 Then
                Call CallWindowProc(m_uRtbl.DecompressCleanup, hCtx)
                hCtx = 0
            End If
        End With
    Next
QH:
    On Error Resume Next
    If uBuf.OutBlock <> 0 Then
        Call CoTaskMemFree(uBuf.OutBlock)
        uBuf.OutBlock = 0
    End If
    If hCtx <> 0 Then
        Call CallWindowProc(m_uRtbl.DecompressCleanup, hCtx)
        hCtx = 0
    End If
    pvVfsClose uArchiveFile
    pvVfsClose uFile
    Exit Function
EH:
    pvSetError MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description
    Resume QH
EH_Continue:
    If pvSetError(MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description, CanContinue:=True) Then
        Resume QH
    Else
        Resume SkipFile
    End If
End Function

Friend Function frCopyCompressed(uOutput As ZipVfsType, ByVal FileIdx As Long, uHdr As ZipLocalHeaderType, oNotify As cZipArchive, ByVal lIdx As Long) As Boolean
    Const FUNC_NAME     As String = "frCopyCompressed"
    Dim baBuffer()      As Byte
    Dim uArchiveFile    As ZipVfsType
    Dim cRead           As Currency
    Dim lSize           As Long

    On Error GoTo EH
    ReDim baBuffer(0 To LNG_IOBUF_SIZE - 1) As Byte
    uArchiveFile = pvVfsOpen(m_vArchiveFile)
    With m_uFiles(FileIdx)
        '--- read local header
        pvVfsSeek uArchiveFile, .Offset, FILE_BEGIN
        pvVfsRead uArchiveFile, VarPtr(uHdr.Signature), 14
        pvVfsRead uArchiveFile, VarPtr(uHdr.Crc32), 16
        '--- sanity checks
        If uHdr.Signature <> LNG_SIG_LOCAL Then
            Err.Raise vbObjectError, , ERR_INVALID_LOCALHDR & ". " & Replace(ERR_ENTRY_INVALID_SIG, "%1", FileIdx + 1)
        End If
        '--- copy local header
        lSize = LNG_LEN_LOCAL + uHdr.LenFname + uHdr.LenExt
        pvVfsSeek uArchiveFile, .Offset, FILE_BEGIN
        pvVfsRead uArchiveFile, VarPtr(baBuffer(0)), lSize
        pvVfsWrite uOutput, VarPtr(baBuffer(0)), lSize
        '--- copy data
        Do ' While cRead < .CompressedSize
            oNotify.frFireProgress lIdx, cRead, .CompressedSize, m_bCancel
            If m_bCancel Then
                m_sLastError = ERR_USER_CANCEL
                GoTo QH
            End If
            If cRead >= .CompressedSize Then
                Exit Do
            End If
            lSize = LNG_IOBUF_SIZE
            If lSize > .CompressedSize - cRead Then
                lSize = .CompressedSize - cRead
            End If
            pvVfsRead uArchiveFile, VarPtr(baBuffer(0)), lSize
            cRead = cRead + lSize
            pvVfsWrite uOutput, VarPtr(baBuffer(0)), lSize
        Loop
    End With
    '--- success
    frCopyCompressed = True
QH:
    pvVfsClose uArchiveFile
    Exit Function
EH:
    pvSetError MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description
    Resume QH
End Function

Friend Sub frFireProgress(ByVal FileIdx As Long, ByVal Current As Currency, ByVal Total As Currency, Cancel As Boolean)
    RaiseEvent Progress(FileIdx, Current, Total, Cancel)
End Sub

Private Function pvFromDosDateTime(ByVal nDate As Integer, ByVal nTime As Integer) As Date
    Dim uSysTime        As SYSTEMTIME
    Dim uFileTime       As FILETIME

    If nDate <> 0 Or nTime <> 0 Then
        Call DosDateTimeToFileTime(nDate, nTime, uFileTime)
        Call FileTimeToSystemTime(uFileTime, uSysTime)
        With uSysTime
            pvFromDosDateTime = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
        End With
    End If
End Function

Private Function pvToWinFileName(baBuffer() As Byte, ByVal lCodePage As Long) As String
    Dim vSplit          As Variant
    Dim lIdx            As Long
    
    vSplit = Split(Replace(pvFromOemString(baBuffer, lCodePage), "/", "\"), "\")
    For lIdx = 0 To UBound(vSplit)
        vSplit(lIdx) = SanitizeFileName(CStr(vSplit(lIdx)), "_")
    Next
    pvToWinFileName = Join(vSplit, "\")
    If Left$(pvToWinFileName, 1) = "\" Then
        Mid$(pvToWinFileName, 1, 1) = "_"
    End If
End Function

#End If ' ImplDecompress

#If ImplInflate Then

Private Function FromBase64Array(sText As String) As Byte()
    Dim lSize           As Long
    Dim dwDummy         As Long
    Dim baOutput()      As Byte

    Call CryptStringToBinary(StrPtr(sText), Len(sText), CRYPT_STRING_BASE64, 0, lSize, 0, dwDummy)
    ReDim baOutput(0 To lSize - 1) As Byte
    Call CryptStringToBinary(StrPtr(sText), Len(sText), CRYPT_STRING_BASE64, VarPtr(baOutput(0)), lSize, 0, dwDummy)
    FromBase64Array = baOutput
End Function

Private Function ToBase64String(baInput() As Byte) As String
    Dim lSize As Long
    If (CryptBinaryToString(VarPtr(baInput(LBound(baInput))), UBound(baInput) - LBound(baInput) + 1, CRYPT_STRING_BASE64, 0&, lSize) <> 0) Then
        ToBase64String = String$(lSize - 1, 0)
        Call CryptBinaryToString(VarPtr(baInput(LBound(baInput))), UBound(baInput) - LBound(baInput) + 1, CRYPT_STRING_BASE64, StrPtr(ToBase64String), lSize)
    End If
End Function

Public Function InflateBase64(sText As String, baOutput() As Byte) As Boolean
    InflateBase64 = Inflate(FromBase64Array(sText), baOutput)
End Function

Public Function Inflate(baBuffer() As Byte, baOutput() As Byte) As Boolean
    Const FUNC_NAME     As String = "Inflate"
    Dim hCtx            As Long
    Dim lOutputPtr      As Long
    Dim lOutputSize     As Long
    
    On Error GoTo EH
    hCtx = InflateInit()
    If Not InflateBlob(hCtx, VarPtr(baBuffer(0)), UBound(baBuffer) + 1, lOutputPtr, lOutputSize) Then
        GoTo QH
    End If
    ReDim baOutput(0 To lOutputSize - 1) As Byte
    Call CopyMemory(baOutput(0), ByVal lOutputPtr, lOutputSize)
    '--- success
    Inflate = True
QH:
    On Error Resume Next
    If lOutputPtr <> 0 Then
        Call CoTaskMemFree(lOutputPtr)
    End If
    If hCtx <> 0 Then
        InflateEnd hCtx
    End If
    Exit Function
EH:
    pvSetError MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description
    Resume QH
End Function

Public Function InflateInit() As Long
    InflateInit = CallWindowProc(m_uRtbl.DecompressInit, VarPtr(m_uRtbl))
    If InflateInit = 0 Then
        Err.Raise vbObjectError, , ERR_INIT_DECOMPRESSOR
    End If
End Function

Public Function InflateBlob(ByVal hCtx As Long, ByVal lInputPtr As Long, ByVal lInputSize As Long, lOutputPtr As Long, lOutputSize As Long) As Boolean
    Const FUNC_NAME     As String = "InflateBlob"
    Dim uBuf            As ZipBuffersType
    Dim lResult         As Long
    
    On Error GoTo EH
    uBuf.InBlock = lInputPtr
    uBuf.InLen = lInputSize
    lResult = CallWindowProc(m_uRtbl.DecompressBlock, hCtx, VarPtr(uBuf))
    If lResult = 0 Or uBuf.OutBlock = 0 Then
        Err.Raise vbObjectError, , ERR_DECOMPRESSING
    End If
    '--- commit
    lOutputPtr = uBuf.OutBlock
    lOutputSize = uBuf.OutLen
    uBuf.OutBlock = 0
    '--- success
    InflateBlob = True
QH:
    On Error Resume Next
    If uBuf.OutBlock <> 0 Then
        Call CoTaskMemFree(uBuf.OutBlock)
        uBuf.OutBlock = 0
    End If
    Exit Function
EH:
    pvSetError MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description
    Resume QH
End Function

Public Sub InflateEnd(ByVal hCtx As Long)
    If hCtx <> 0 Then
        Call CallWindowProc(m_uRtbl.DecompressCleanup, hCtx)
    End If
End Sub

Public Function DeflateBase64(baInput() As Byte, sText As String, Optional ByVal Level As Long = 6) As Boolean
    Dim baOutput()      As Byte
    
    DeflateBase64 = Deflate(baInput, baOutput, Level)
    If DeflateBase64 Then
        sText = ToBase64String(baOutput)
    End If
End Function

Public Function Deflate(baBuffer() As Byte, baOutput() As Byte, Optional ByVal Level As Long = 6) As Boolean
    Const FUNC_NAME     As String = "Deflate"
    Dim hCtx            As Long
    Dim lOutputPtr      As Long
    Dim lOutputSize     As Long
    
    On Error GoTo EH
    hCtx = DeflateInit()
    If Not DeflateBlob(hCtx, VarPtr(baBuffer(0)), UBound(baBuffer) + 1, lOutputPtr, lOutputSize, Level, True) Then
        GoTo QH
    End If
    ReDim baOutput(0 To lOutputSize - 1) As Byte
    Call CopyMemory(baOutput(0), ByVal lOutputPtr, lOutputSize)
    '--- success
    Deflate = True
QH:
    On Error Resume Next
    If lOutputPtr <> 0 Then
        Call CoTaskMemFree(lOutputPtr)
    End If
    If hCtx <> 0 Then
        DeflateEnd hCtx
    End If
    Exit Function
EH:
    pvSetError MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description
    Resume QH
End Function

Public Function DeflateInit() As Long
    DeflateInit = CallWindowProc(m_uRtbl.CompressInit, VarPtr(m_uRtbl))
    If DeflateInit = 0 Then
        Err.Raise vbObjectError, , ERR_INIT_COMPRESSOR
    End If
End Function

Public Function DeflateBlob(ByVal hCtx As Long, ByVal lInputPtr As Long, ByVal lInputSize As Long, lOutputPtr As Long, lOutputSize As Long, Optional ByVal Level As Long = 6, Optional ByVal Final As Boolean) As Boolean
    Const FUNC_NAME     As String = "DeflateBlob"
    Dim uBuf            As ZipBuffersType
    Dim lResult         As Long
    
    On Error GoTo EH
    uBuf.Greedy = (Level <= 4)
    uBuf.MaxMatch = At(Array(0, 2, 6, 12, 24, 8, 16, 32, 64, 1000), Level)
    uBuf.NiceLen = At(Array(0, 8, 10, 14, 24, 30, 65, 130, 200, 32768), Level)
    uBuf.InBlock = lInputPtr
    uBuf.InLen = lInputSize
    uBuf.Final = -Final
    lResult = CallWindowProc(m_uRtbl.CompressBlock, hCtx, VarPtr(uBuf))
    If lResult = 0 Or uBuf.OutBlock = 0 Then
        Err.Raise vbObjectError, , ERR_COMPRESSING
    End If
    '--- commit
    lOutputPtr = uBuf.OutBlock
    lOutputSize = uBuf.OutLen
    uBuf.OutBlock = 0
    '--- success
    DeflateBlob = True
QH:
    On Error Resume Next
    If uBuf.OutBlock <> 0 Then
        Call CoTaskMemFree(uBuf.OutBlock)
        uBuf.OutBlock = 0
    End If
    Exit Function
EH:
    pvSetError MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description
    Resume QH
End Function

Public Sub DeflateEnd(ByVal hCtx As Long)
    If hCtx <> 0 Then
        Call CallWindowProc(m_uRtbl.CompressCleanup, hCtx)
    End If
End Sub

#End If ' ImplInflate

'= Crypto ================================================================

#If ImplCrypto Then

Private Sub pvCryptoGetRandom(ByVal lSize As Long, baOutput() As Byte)
    ReDim baOutput(0 To lSize - 1) As Byte
    Call RtlGenRandom(baOutput(0), lSize)
End Sub

Private Function pvCryptoAesInit(baPass() As Byte, baSalt() As Byte, ByVal lKeyLen As Long, nPassVer As Integer) As Boolean
    Dim baDerivedKey()  As Byte
    Dim lDummy          As Long '--- discarded
    Dim hResult         As Long
    Dim sApiSource      As String
    
    '--- init member vars
    m_uCrypto.Nonce(0) = 0
    m_uCrypto.Nonce(1) = 0
    m_uCrypto.EncrData = vbNullString
    m_uCrypto.EncrPos = 0
    '--- generate RFC 2898 based derived key
    If m_uCrypto.hPbkdf2Alg = 0 Then
        '--- CNG API missing on XP
        On Error GoTo EH_Unsupported
        hResult = BCryptOpenAlgorithmProvider(m_uCrypto.hPbkdf2Alg, StrPtr(BCRYPT_SHA1_ALGORITHM), StrPtr(MS_PRIMITIVE_PROVIDER), BCRYPT_ALG_HANDLE_HMAC_FLAG)
        If hResult < 0 Then
            sApiSource = "BCryptOpenAlgorithmProvider(" & BCRYPT_SHA1_ALGORITHM & ")"
            GoTo QH
        End If
        On Error GoTo 0
    End If
    ReDim baDerivedKey(0 To 2 * lKeyLen + 1) As Byte
    '--- PBKDF2 API missing on Vista
    On Error GoTo EH_Unsupported
    hResult = BCryptDeriveKeyPBKDF2(m_uCrypto.hPbkdf2Alg, baPass(0), UBound(baPass) + 1, baSalt(0), UBound(baSalt) + 1, 1000, 0, baDerivedKey(0), UBound(baDerivedKey) + 1, 0)
    If hResult < 0 Then
        sApiSource = "BCryptDeriveKeyPBKDF2"
        GoTo QH
    End If
    On Error GoTo 0
    '--- extract Password Verification Value from last 2 bytes of derived key
    Call CopyMemory(nPassVer, baDerivedKey(2 * lKeyLen), 2)
    '--- init AES key from first half of the derived key
    If m_uCrypto.hAesAlg = 0 Then
        hResult = BCryptOpenAlgorithmProvider(m_uCrypto.hAesAlg, StrPtr(BCRYPT_AES_ALGORITHM), StrPtr(MS_PRIMITIVE_PROVIDER), 0)
        If hResult < 0 Then
            sApiSource = "BCryptOpenAlgorithmProvider(" & BCRYPT_AES_ALGORITHM & ")"
            GoTo QH
        End If
        hResult = BCryptGetProperty(m_uCrypto.hAesAlg, StrPtr(BCRYPT_OBJECT_LENGTH), m_uCrypto.AesKeyObjLen, 4, lDummy, 0)
        If hResult < 0 Then
            sApiSource = "BCryptGetProperty(" & BCRYPT_OBJECT_LENGTH & ")"
            GoTo QH
        End If
        hResult = BCryptSetProperty(m_uCrypto.hAesAlg, StrPtr(BCRYPT_CHAINING_MODE), StrPtr(BCRYPT_CHAIN_MODE_ECB), LenB(BCRYPT_CHAIN_MODE_ECB), 0)
        If hResult < 0 Then
            sApiSource = "BCryptSetProperty(" & BCRYPT_CHAINING_MODE & ")"
            GoTo QH
        End If
    End If
    If m_uCrypto.hAesKey <> 0 Then
        Call BCryptDestroyKey(m_uCrypto.hAesKey)
        m_uCrypto.hAesKey = 0
    End If
    ReDim m_uCrypto.AesKeyObjData(0 To m_uCrypto.AesKeyObjLen - 1) As Byte
    hResult = BCryptGenerateSymmetricKey(m_uCrypto.hAesAlg, m_uCrypto.hAesKey, m_uCrypto.AesKeyObjData(0), m_uCrypto.AesKeyObjLen, baDerivedKey(0), lKeyLen, 0)
    If hResult < 0 Then
        sApiSource = "BCryptGenerateSymmetricKey"
        GoTo QH
    End If
    '--- init HMAC key from second half of the derived key
    If m_uCrypto.hHmacAlg = 0 Then
        hResult = BCryptOpenAlgorithmProvider(m_uCrypto.hHmacAlg, StrPtr(BCRYPT_SHA1_ALGORITHM), StrPtr(MS_PRIMITIVE_PROVIDER), BCRYPT_ALG_HANDLE_HMAC_FLAG)
        If hResult < 0 Then
            sApiSource = "BCryptOpenAlgorithmProvider(" & BCRYPT_SHA1_ALGORITHM & ")"
            GoTo QH
        End If
        hResult = BCryptGetProperty(m_uCrypto.hHmacAlg, StrPtr(BCRYPT_HASH_LENGTH), m_uCrypto.HmacLen, 4, lDummy, 0)
        If hResult < 0 Then
            sApiSource = "BCryptGetProperty(" & BCRYPT_HASH_LENGTH & ")"
            GoTo QH
        End If
    End If
    If m_uCrypto.hHmacHash <> 0 Then
        Call BCryptDestroyHash(m_uCrypto.hHmacHash)
        m_uCrypto.hHmacHash = 0
    End If
    hResult = BCryptCreateHash(m_uCrypto.hHmacAlg, m_uCrypto.hHmacHash, 0, 0, baDerivedKey(lKeyLen), lKeyLen, 0)
    If hResult < 0 Then
        sApiSource = "BCryptCreateHash"
        GoTo QH
    End If
    '--- success
    pvCryptoAesInit = True
    Exit Function
QH:
    If hResult < 0 Then
        m_sLastError = GetSystemMessage(hResult) & " [" & sApiSource & "]"
    ElseIf Err.LastDllError <> 0 Then
        m_sLastError = GetSystemMessage(Err.LastDllError)
    End If
    Exit Function
EH_Unsupported:
    m_sLastError = ERR_UNSUPPORTED_ENCRYPTION
End Function

Private Sub pvCryptoTerminate()
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

Private Function pvCryptoAesCrypt( _
            baData() As Byte, _
            Optional ByVal Offset As Long, _
            Optional ByVal Size As Long, _
            Optional ByVal HashBefore As Boolean, _
            Optional ByVal HashAfter As Boolean) As Boolean
    Dim lIdx            As Long
    Dim lPadSize        As Long
    Dim lDummy          As Long '--- discarded
    Dim hResult         As Long
    Dim sApiSource      As String
    
    If Size < 0 Then
        Size = UBound(baData) + 1 - Offset
    End If
    If HashBefore Then
        hResult = BCryptHashData(m_uCrypto.hHmacHash, baData(Offset), Size, 0)
        If hResult < 0 Then
            sApiSource = "BCryptHashData"
            GoTo QH
        End If
    End If
    With m_uCrypto
        '--- reuse EncrData from prev call until next LNG_AES_BLOCK_SIZE boundary
        For lIdx = Offset To Offset + Size - 1
            If (.EncrPos And (LNG_AES_BLOCK_SIZE - 1)) = 0 Then
                Exit For
            End If
            baData(lIdx) = baData(lIdx) Xor .EncrData(.EncrPos)
            .EncrPos = .EncrPos + 1
        Next
        If lIdx < Offset + Size Then
            '--- pad remaining input size to LNG_AES_BLOCK_SIZE
            lPadSize = (Offset + Size - lIdx + LNG_AES_BLOCK_SIZE - 1) And -LNG_AES_BLOCK_SIZE
            If UBound(.EncrData) + 1 < lPadSize Then
                ReDim .EncrData(0 To lPadSize - 1) As Byte
            End If
            '--- encrypt incremental nonces in EncrData
            Call CallWindowProc(m_uRtbl.MemNonce, VarPtr(.EncrData(0)), VarPtr(.Nonce(0)), lPadSize)
            hResult = BCryptEncrypt(.hAesKey, .EncrData(0), lPadSize, 0, 0, 0, .EncrData(0), lPadSize, lDummy, 0)
            If hResult < 0 Then
                sApiSource = "BCryptEncrypt"
                GoTo QH
            End If
            '--- xor remaining input and leave anything extra of EncrData for reuse
            .EncrPos = Offset + Size - lIdx
            Call CallWindowProc(m_uRtbl.MemXor, VarPtr(.EncrData(0)), VarPtr(baData(lIdx)), .EncrPos)
        End If
    End With
    If HashAfter Then
        hResult = BCryptHashData(m_uCrypto.hHmacHash, baData(Offset), Size, 0)
        If hResult < 0 Then
            sApiSource = "BCryptHashData"
            GoTo QH
        End If
    End If
    '--- success
    pvCryptoAesCrypt = True
    Exit Function
QH:
    If hResult < 0 Then
        m_sLastError = GetSystemMessage(hResult) & " [" & sApiSource & "]"
    ElseIf Err.LastDllError <> 0 Then
        m_sLastError = GetSystemMessage(Err.LastDllError)
    End If
End Function

Private Function pvCryptoAesGetFinalHash(ByVal lSize As Long) As Byte()
    Dim baResult()      As Byte
    
    ReDim baResult(0 To m_uCrypto.HmacLen - 1) As Byte
    Call BCryptFinishHash(m_uCrypto.hHmacHash, baResult(0), m_uCrypto.HmacLen, 0)
    ReDim Preserve baResult(0 To lSize - 1) As Byte
    pvCryptoAesGetFinalHash = baResult
End Function

Private Function pvCryptoTradInit(baPass() As Byte) As Boolean
    With m_uCrypto
        .TradKey(0) = &H12345678
        .TradKey(1) = &H23456789
        .TradKey(2) = &H34567890
        .TradKey(3) = m_uRtbl.Crc32Table
        Call CallWindowProc(m_uRtbl.ZipCrypt, VarPtr(.TradKey(0)), VarPtr(baPass(0)), UBound(baPass) + 1)
    End With
    '--- success
    pvCryptoTradInit = True
End Function

Private Function pvCryptoTradCrypt( _
            baData() As Byte, _
            Optional ByVal Offset As Long, _
            Optional ByVal Size As Long = -1, _
            Optional ByVal UseAfter As Boolean) As Boolean
    If Size < 0 Then
        Size = UBound(baData) + 1 - Offset
    End If
    Call CallWindowProc(m_uRtbl.ZipCrypt, VarPtr(m_uCrypto.TradKey(0)), VarPtr(baData(Offset)), Size, IIf(UseAfter, 3, 1))
    '--- success
    pvCryptoTradCrypt = True
End Function

#End If ' ImplCrypto

'= VFS ===================================================================

Private Function pvVfsOpen(File As Variant) As ZipVfsType
    Dim hFind           As Long
    
    If IsArray(File) Then
        If VarType(File) = vbArray + vbByte Then
            pvVfsOpen.FileName = STR_BUFFER
            If (Peek(VarPtr(File)) And VT_BYREF) <> 0 Then
                Call CopyMemory(pvVfsOpen.BufferArray, File, 16)
            Else
                pvVfsOpen.BufferArray = File
            End If
            pvVfsOpen.BufferBase = pvArrPtr(File)
            pvVfsOpen.BufferPtr = pvVfsOpen.BufferBase
            If pvVfsOpen.BufferBase <> 0 Then
                pvVfsOpen.BufferSize = UBound(File) + 1
            End If
            pvVfsOpen.Data.nFileSizeLow = pvFromInt64(pvVfsOpen.BufferSize, pvVfsOpen.Data.nFileSizeHigh)
            pvVfsOpen.Data.dwFileAttributes = vbArchive
            pvVfsOpen.Data.ftLastWriteTime = pvToFileTime(VBA.Now)
        Else
            Debug.Assert VarType(File) = vbArray + vbVariant
            pvVfsOpen.FileName = STR_ARCHIVE
            Set pvVfsOpen.SourceArchive = File(0)
            pvVfsOpen.SourceFileIdx = File(1)
            pvVfsOpen.SourceFileInfo = pvVfsOpen.SourceArchive.FileInfo(pvVfsOpen.SourceFileIdx)
            pvVfsOpen.Data.nFileSizeLow = pvFromInt64(pvVfsOpen.SourceFileInfo(zipIdxSize), pvVfsOpen.Data.nFileSizeHigh)
            pvVfsOpen.Data.dwFileAttributes = pvVfsOpen.SourceFileInfo(zipIdxAttributes)
            pvVfsOpen.Data.ftLastWriteTime = pvToFileTime(pvVfsOpen.SourceFileInfo(zipIdxLastModified))
        End If
    ElseIf IsObject(File) Then
        pvVfsOpen.FileName = STR_STREAM
        Set pvVfsOpen.Stream = File
        If Not pvVfsOpen.Stream Is Nothing Then
            pvVfsOpen.Data.nFileSizeLow = pvFromInt64(pvVfsOpen.Stream.VfsSetFilePointer(0, FILE_END), pvVfsOpen.Data.nFileSizeHigh)
        End If
        pvVfsOpen.Data.dwFileAttributes = vbArchive
        pvVfsOpen.Data.ftLastWriteTime = pvToFileTime(VBA.Now)
    Else
        pvVfsOpen.FileName = File
        If Right$(File, 1) = ":" Then
            pvVfsOpen.Data.dwFileAttributes = vbArchive
            pvVfsOpen.Data.ftLastWriteTime = pvToFileTime(VBA.Now)
        ElseIf Right$(File, 1) <> "\" Then
            hFind = FindFirstFile(StrPtr(File), VarPtr(pvVfsOpen.Data))
        Else
            hFind = FindFirstFile(StrPtr(Left$(File, Len(File) - 1)), VarPtr(pvVfsOpen.Data))
        End If
        If hFind = INVALID_HANDLE_VALUE Then
            Err.Raise vbObjectError, "FindFirstFile", GetSystemMessage(Err.LastDllError) & " (" & File & ")"
        End If
        Call FindClose(hFind)
    End If
End Function

Private Function pvCreateFile(ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal NoSecurity As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
    Select Case LCase$(lpFileName)
    Case "stdin:"
        pvCreateFile = GetStdHandle(STD_INPUT_HANDLE)
    Case "stdout:"
        pvCreateFile = GetStdHandle(STD_OUTPUT_HANDLE)
    Case "stderr:"
        pvCreateFile = GetStdHandle(STD_ERROR_HANDLE)
    Case Else
        pvCreateFile = CreateFile(StrPtr(lpFileName), dwDesiredAccess, dwShareMode, NoSecurity, dwCreationDisposition, dwFlagsAndAttributes, hTemplateFile)
    End Select
End Function

Private Function pvVfsCreate(File As Variant, Optional ByVal eAttrib As VbFileAttribute = vbArchive) As ZipVfsType
    Dim sError          As String
    
    If IsArray(File) Then
        pvVfsCreate.FileName = STR_BUFFER
        If (Peek(VarPtr(File)) And VT_BYREF) <> 0 Then
            Call CopyMemory(pvVfsCreate.BufferArray, File, 16)
        Else
            Err.Raise vbObjectError, , ERR_INVALID_BYTEARRAY
        End If
        pvVfsCreate.BufferBase = pvArrPtr(pvVfsCreate.BufferArray)
        pvVfsCreate.BufferPtr = pvVfsCreate.BufferBase
        If pvVfsCreate.BufferBase <> 0 Then
            pvVfsCreate.BufferSize = UBound(pvVfsCreate.BufferArray) + 1
        End If
    ElseIf IsObject(File) Then
        pvVfsCreate.FileName = STR_STREAM
        Set pvVfsCreate.Stream = File
    Else
        pvVfsCreate.FileName = File
        If InStrRev(File, "\") > 0 Then
            If Not MkPath(Left$(File, InStrRev(File, "\") - 1), sError) Then
                Err.Raise vbObjectError, , sError
            End If
        End If
        If Right$(File, 1) <> "\" Then
            pvVfsCreate.Handle = pvCreateFile(File, GENERIC_WRITE, FILE_SHARE_READ, 0, CREATE_ALWAYS, eAttrib, 0)
            If pvVfsCreate.Handle = INVALID_HANDLE_VALUE Then
                Err.Raise vbObjectError, "CreateFile", GetSystemMessage(Err.LastDllError) & " (" & File & ")"
            End If
        Else
            Call SetFileAttributes(StrPtr(Left$(File, Len(File) - 1)), eAttrib)
        End If
    End If
End Function

Private Sub pvVfsClose(uFile As ZipVfsType)
    Dim vEmpty          As Variant
    
    If Not IsEmpty(uFile.BufferArray) Then
        If (Peek(VarPtr(uFile.BufferArray)) And VT_BYREF) <> 0 Then
            Call CopyMemory(uFile.BufferArray, vEmpty, 16)
        Else
            uFile.BufferArray = Empty
        End If
    End If
    If uFile.Handle <> 0 And uFile.Handle <> INVALID_HANDLE_VALUE Then
        Call CloseHandle(uFile.Handle)
        uFile.Handle = 0
    End If
    Set uFile.Stream = Nothing
    Set uFile.SourceArchive = Nothing
End Sub

Private Function pvVfsRead(uFile As ZipVfsType, ByVal lPtr As Long, ByVal lSize As Long) As Long
    If Not IsEmpty(uFile.BufferArray) Then
        pvVfsRead = Clamp(lSize, 0, uFile.BufferSize - (uFile.BufferPtr - uFile.BufferBase))
        Call CopyMemory(ByVal lPtr, ByVal uFile.BufferPtr, pvVfsRead)
        uFile.BufferPtr = uFile.BufferPtr + pvVfsRead
    ElseIf Not uFile.Stream Is Nothing Then
        If uFile.Handle = 0 Then
            uFile.Handle = INVALID_HANDLE_VALUE
            uFile.Stream.VfsSetFilePointer 0, FILE_BEGIN
        End If
        pvVfsRead = uFile.Stream.VfsReadFile(lPtr, lSize)
    Else
        If uFile.Handle = 0 Or uFile.Handle = INVALID_HANDLE_VALUE Then
            uFile.Handle = pvCreateFile(uFile.FileName, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, 0, 0)
            If uFile.Handle = INVALID_HANDLE_VALUE Then
                Err.Raise vbObjectError, "CreateFile", GetSystemMessage(Err.LastDllError) & " (" & uFile.FileName & ")"
            End If
        End If
        If ReadFile(uFile.Handle, ByVal lPtr, lSize, pvVfsRead, 0) = 0 Then
            If Err.LastDllError <> LNG_PIPE_ENDED_ERROR Then
                Err.Raise vbObjectError, "ReadFile", GetSystemMessage(Err.LastDllError) & " (" & uFile.FileName & ")"
            End If
        End If
    End If
End Function

Private Function pvVfsWrite(uFile As ZipVfsType, ByVal lPtr As Long, ByVal lSize As Long) As Long
    Const MAX_STEP      As Long = 2 ^ 24 '--- 16MB
    Dim lOffset         As Long
    
    If Not IsEmpty(uFile.BufferArray) Then
        lOffset = uFile.BufferPtr - uFile.BufferBase
        If lOffset + lSize > uFile.BufferSize Then
            If lOffset + lSize <= MAX_STEP Then
                uFile.BufferSize = Clamp(2 ^ Int(Log(lOffset + lSize) / Log(2) + 1), 8192)
            Else
                uFile.BufferSize = (lOffset + lSize + MAX_STEP - 1) And -MAX_STEP
            End If
            ReDim Preserve uFile.BufferArray(0 To uFile.BufferSize - 1) As Byte
            uFile.BufferBase = pvArrPtr(uFile.BufferArray)
            uFile.BufferPtr = uFile.BufferBase + lOffset
        End If
        pvVfsWrite = lSize
        Call CopyMemory(ByVal uFile.BufferPtr, ByVal lPtr, pvVfsWrite)
        uFile.BufferPtr = uFile.BufferPtr + pvVfsWrite
    ElseIf Not uFile.Stream Is Nothing Then
        If uFile.Handle = 0 Then
            uFile.Handle = INVALID_HANDLE_VALUE
            uFile.Stream.VfsSetFilePointer 0, FILE_BEGIN
        End If
        pvVfsWrite = uFile.Stream.VfsWriteFile(lPtr, lSize)
    Else
        If uFile.Handle = 0 Or uFile.Handle = INVALID_HANDLE_VALUE Then
            uFile.Handle = pvCreateFile(uFile.FileName, GENERIC_WRITE, FILE_SHARE_READ, 0, CREATE_ALWAYS, vbArchive, 0)
            If uFile.Handle = INVALID_HANDLE_VALUE Then
                Err.Raise vbObjectError, "CreateFile", GetSystemMessage(Err.LastDllError) & " (" & uFile.FileName & ")"
            End If
        End If
        If WriteFile(uFile.Handle, ByVal lPtr, lSize, pvVfsWrite, 0) = 0 Then
            Err.Raise vbObjectError, "WriteFile", GetSystemMessage(Err.LastDllError) & " (" & uFile.FileName & ")"
        End If
    End If
End Function

Private Function pvVfsSeek(uFile As ZipVfsType, ByVal cPosition As Currency, ByVal lMoveMethod As Long) As Currency
    If Not IsEmpty(uFile.BufferArray) Then
        Select Case lMoveMethod
        Case FILE_BEGIN
            pvVfsSeek = cPosition
        Case FILE_CURRENT
            pvVfsSeek = uFile.BufferPtr - uFile.BufferBase + cPosition
        Case FILE_END
            pvVfsSeek = uFile.BufferSize + cPosition
        End Select
        uFile.BufferPtr = uFile.BufferBase + pvVfsSeek
    ElseIf Not uFile.Stream Is Nothing Then
        pvVfsSeek = uFile.Stream.VfsSetFilePointer(cPosition, lMoveMethod)
    Else
        If uFile.Handle = 0 Or uFile.Handle = INVALID_HANDLE_VALUE Then
            '--- touch to open for read
            pvVfsRead uFile, 0, 0
        End If
        If SetFilePointerEx(uFile.Handle, cPosition / 10000@, pvVfsSeek, lMoveMethod) = 0 Then
            Err.Raise vbObjectError, "SetFilePointerEx", GetSystemMessage(Err.LastDllError) & " (" & uFile.FileName & ")"
        End If
        pvVfsSeek = pvVfsSeek * 10000@
    End If
End Function

Private Sub pvVfsSetEof(uFile As ZipVfsType, sMetaData As String, ByVal dLastWriteTime As Date)
    If Not IsEmpty(uFile.BufferArray) Then
        If uFile.BufferSize <> uFile.BufferPtr - uFile.BufferBase Then
            uFile.BufferSize = uFile.BufferPtr - uFile.BufferBase
            If uFile.BufferSize > 0 Then
                ReDim Preserve uFile.BufferArray(0 To uFile.BufferSize - 1) As Byte
            Else
                Erase uFile.BufferArray
            End If
            uFile.BufferBase = pvArrPtr(uFile.BufferArray)
            uFile.BufferPtr = uFile.BufferBase + uFile.BufferSize
        End If
    ElseIf Not uFile.Stream Is Nothing Then
        uFile.Stream.VfsSetEndOfFile sMetaData ' , dLastWriteTime
    ElseIf Right$(uFile.FileName, 1) <> ":" Then
        If SetEndOfFile(uFile.Handle) = 0 Then
            Err.Raise vbObjectError, "SetEndOfFile", GetSystemMessage(Err.LastDllError) & " (" & uFile.FileName & ")"
        End If
        If dLastWriteTime <> 0 Then
            If SetFileTime(uFile.Handle, ByVal 0, ByVal 0, pvToFileTime(dLastWriteTime)) = 0 Then
                Err.Raise vbObjectError, "SetFileTime", GetSystemMessage(Err.LastDllError) & " (" & uFile.FileName & ")"
            End If
        End If
    End If
End Sub

Private Function pvArrPtr(vArray As Variant) As Long
    Dim lPtr            As Long
    
    lPtr = Peek(UnsignedAdd(VarPtr(vArray), 8))
    If (Peek(VarPtr(vArray)) And VT_BYREF) <> 0 Then
        lPtr = Peek(lPtr)
    End If
    If lPtr <> 0 Then
        pvArrPtr = Peek(UnsignedAdd(lPtr, 12))
    End If
End Function

'= common ================================================================

Private Function pvFromOemString(baBuffer() As Byte, ByVal lCodePage As Long) As String
    Dim lSize           As Long
    
    If UBound(baBuffer) >= 0 Then
        lSize = MultiByteToWideChar(lCodePage, 0, baBuffer(0), UBound(baBuffer) + 1, 0, 0)
        If lSize > 0 Then
            pvFromOemString = String$(lSize, 0)
            lSize = MultiByteToWideChar(lCodePage, 0, baBuffer(0), UBound(baBuffer) + 1, ByVal StrPtr(pvFromOemString), lSize)
        End If
    End If
End Function

Private Function pvFromFileTime(uFileTime As FILETIME) As Date
    Dim uLocalTime      As FILETIME
    Dim uSysTime        As SYSTEMTIME
    
    If FileTimeToLocalFileTime(uFileTime, uLocalTime) <> 0 Then
        If FileTimeToSystemTime(uLocalTime, uSysTime) <> 0 Then
            Call SystemTimeToVariantTime(uSysTime, pvFromFileTime)
        End If
    End If
End Function

Private Function pvToFileTime(ByVal dDate As Date) As FILETIME
    Dim uSysTime        As SYSTEMTIME
    Dim uLocalTime      As FILETIME

    If VariantTimeToSystemTime(dDate, uSysTime) <> 0 Then
        If SystemTimeToFileTime(uSysTime, uLocalTime) <> 0 Then
            Call LocalFileTimeToFileTime(uLocalTime, pvToFileTime)
        End If
    End If
End Function

Private Function pvFromInt64(ByVal cFileSize As Currency, Optional HiPart As Long) As Long
    cFileSize = cFileSize / 10000@
    Call CopyMemory(pvFromInt64, cFileSize, 4)
    Call CopyMemory(HiPart, ByVal UnsignedAdd(VarPtr(cFileSize), 4), 4)
End Function

Private Function pvToInt64(ByVal lLowPart As Long, Optional ByVal lHiPart As Long) As Currency
    Call CopyMemory(pvToInt64, lLowPart, 4)
    If lHiPart <> 0 Then
        Call CopyMemory(ByVal UnsignedAdd(VarPtr(pvToInt64), 4), lHiPart, 4)
    End If
    pvToInt64 = pvToInt64 * 10000@
End Function

Private Function pvToInt32(ByVal cValue As Currency) As Long
    cValue = cValue / 10000@
    Call CopyMemory(pvToInt32, cValue, 4)
End Function

Private Function pvSetError(Optional Source As String, Optional Description As String, Optional ByVal CanContinue As Boolean) As Boolean
    If StrPtr(Source) = 0 And StrPtr(Description) = 0 Then
        m_sLastError = vbNullString
        m_bCancel = False
    Else
        #If ImplLogging Then
            If m_lCurrentFile >= 0 Then
                Debug.Print "Error in " & Source & ": " & Description, m_uFiles(m_lCurrentFile).FileName, Timer
            Else
                Debug.Print "Error in " & Source & ": " & Description, Timer
            End If
        #End If ' ImplLogging
        m_sLastError = Description
        If Not CanContinue Then
            m_bCancel = True
        End If
        RaiseEvent Error(m_lCurrentFile, Source, m_sLastError, m_bCancel)
        pvSetError = m_bCancel
    End If
End Function

Private Function pvPathGetFolder(sPath As String, Optional sFile As String) As String
    Dim lPos            As Long
    
    lPos = InStrRev(sPath, "\")
    If lPos > 0 Then
        pvPathGetFolder = Left$(sPath, lPos - 1)
        sFile = Mid$(sPath, lPos + 1)
    Else
        pvPathGetFolder = sPath
        sFile = vbNullString
    End If
End Function

Private Function pvFileGetFName(sFile As String) As String
    Dim lPos            As Long
    
    lPos = InStrRev(sFile, ".")
    If lPos > InStrRev(sFile, "\") Then
        pvFileGetFName = Left$(sFile, lPos - 1)
    Else
        pvFileGetFName = sFile
    End If
End Function

Private Function pvArrayEqual(baFirst() As Byte, baSecond() As Byte) As Boolean
    If UBound(baFirst) = UBound(baSecond) Then
        pvArrayEqual = (InStrB(baFirst, baSecond) = 1)
    End If
End Function

'= zlib thunks ===========================================================

Private Function pvInitRelocTable(uRtbl As ZipRelocTableType) As Long
    Dim lpThunk         As Long
    Dim vSplit          As Variant
    
    lpThunk = pvGetThunkAddress()
    vSplit = Split(STR_THUNK_OFFSETS, "|")
    With uRtbl
        .CompressInit = lpThunk + vSplit(zipRelCompressInit)
        .CompressCleanup = lpThunk + vSplit(zipRelCompressCleanup)
        .CompressBlock = lpThunk + vSplit(zipRelCompressBlock)
        .DecompressInit = lpThunk + vSplit(zipRelDecompressInit)
        .DecompressCleanup = lpThunk + vSplit(zipRelDecompressCleanup)
        .DecompressBlock = lpThunk + vSplit(zipRelDecompressBlock)
        .CalcCrc32 = lpThunk + vSplit(zipRelCalcCrc32)
        .MemNonce = lpThunk + vSplit(zipRelMemNonce)
        .MemXor = lpThunk + vSplit(zipRelMemXor)
        .ZipCrypt = lpThunk + vSplit(zipRelZipCrypt)
        .MallocImpl = GetProcAddress(GetModuleHandle(StrPtr("ole32")), "CoTaskMemAlloc")
        .ReallocImpl = GetProcAddress(GetModuleHandle(StrPtr("ole32")), "CoTaskMemRealloc")
        .FreeImpl = GetProcAddress(GetModuleHandle(StrPtr("ole32")), "CoTaskMemFree")
        .LenCodes = lpThunk + vSplit(zipRelLenCodes)
        .DistCodes = lpThunk + vSplit(zipRelDistCodes)
        .MirrorBytes = lpThunk + vSplit(zipRelMirrorBytes)
        .LenLenMap = lpThunk + vSplit(zipRelLenLenMap)
        .Crc32Table = pvGetCrc32Table()
    End With
End Function

Private Function pvGetThunkAddress() As Long
    Static lpThunk      As Long
    Dim baThunk()       As Byte
    Dim sBuffer         As String
    
    If lpThunk = 0 Then
        sBuffer = String$(50, 0)
        Call GetEnvironmentVariable(StrPtr("_ZIP_THUNK_" & GetCurrentProcessId() & "_" & STR_THUNK_BUILDDATE), StrPtr(sBuffer), Len(sBuffer) - 1)
        lpThunk = Val(sBuffer)
        If lpThunk = 0 Then
            pvGetThunkData baThunk
            lpThunk = VirtualAlloc(0, UBound(baThunk) + 1, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
            Call CopyMemory(ByVal lpThunk, baThunk(0), UBound(baThunk) + 1)
            Call SetEnvironmentVariable(StrPtr("_ZIP_THUNK_" & GetCurrentProcessId() & "_" & STR_THUNK_BUILDDATE), StrPtr(lpThunk))
        End If
    End If
    pvGetThunkAddress = lpThunk
End Function

Private Function pvGetCrc32Table() As Long
    Static aTable()     As Long
    Static bIsInit      As Boolean
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim lReminder       As Long
    Dim lValue          As Long

    If Not bIsInit Then
        '--- table mem allocated: 4KB
        ReDim aTable(0 To &H3FF) As Long
        For lIdx = 0 To &H3FF
            If lIdx < &H100 Then
                lReminder = 0
                lValue = lIdx
            Else
                lReminder = aTable(lIdx - &H100)
                lValue = 0
            End If
            For lJdx = 1 To 8
                If ((lReminder Xor lValue) And 1) <> 0 Then
                  lReminder = (lReminder And &HFFFFFFFE) \ 2 And &H7FFFFFFF Xor &HEDB88320
                Else
                  lReminder = (lReminder And &HFFFFFFFE) \ 2 And &H7FFFFFFF
                End If
                lValue = (lValue And &HFFFFFFFE) \ 2 And &H7FFFFFFF
            Next
            aTable(lIdx) = lReminder
        Next
        bIsInit = True
    End If
    pvGetCrc32Table = VarPtr(aTable(0))
End Function

Public Function CalcCrc32Array(baData() As Byte) As Long
    CalcCrc32Array = -1
    Call CallWindowProc(m_uRtbl.CalcCrc32, VarPtr(m_uRtbl), VarPtr(baData(0)), UBound(baData) + 1, VarPtr(CalcCrc32Array))
    CalcCrc32Array = CalcCrc32Array Xor -1
End Function

Public Sub CalcCrc32Ptr(ByVal lPtr As Long, ByVal lSize As Long, lCrc32 As Long)
    Call CallWindowProc(m_uRtbl.CalcCrc32, VarPtr(m_uRtbl), lPtr, lSize, VarPtr(lCrc32))
End Sub

Private Sub pvAppendBuffer(ByVal a01 As Long, ByVal a02 As Long, ByVal a03 As Long, ByVal a04 As Long, ByVal a05 As Long, ByVal a06 As Long, ByVal a07 As Long, ByVal a08 As Long, ByVal a09 As Long, ByVal a10 As Long, ByVal a11 As Long, ByVal a12 As Long, ByVal a13 As Long, ByVal a14 As Long, ByVal a15 As Long, ByVal a16 As Long, ByVal a17 As Long, ByVal a18 As Long, ByVal a19 As Long, ByVal a20 As Long, ByVal a21 As Long, ByVal a22 As Long, ByVal a23 As Long, ByVal a24 As Long, ByVal a25 As Long, ByVal a26 As Long, ByVal a27 As Long, ByVal a28 As Long, ByVal a29 As Long, ByVal a30 As Long, ByVal a31 As Long, ByVal a32 As Long)
    #If a01 And a02 And a03 And a04 And a05 And a06 And a07 And a08 And a09 And a10 And a11 And a12 And a13 And a14 And a15 And a16 And a17 And a18 And a19 And a20 And a21 And a22 And a23 And a24 And a25 And a26 And a27 And a28 And a29 And a30 And a31 And a32 Then '--- touch args
    #End If
    Call CopyMemory(m_baBuffer(m_lBuffIdx), a01, 4 * 32)
    m_lBuffIdx = m_lBuffIdx + 4 * 32
End Sub

Private Sub pvGetThunkData(baBuffer() As Byte)
    ReDim m_baBuffer(0 To 7680 - 1) As Byte
    m_lBuffIdx = 0
    '--- begin thunk data
    pvAppendBuffer &H24448B51, &H5C8B5308, &H8B551024, &H5618246C, &H44788B57, &H448BEB03, &H6C892424, &H8B1024, &H7403C3F6, &H74DD3B1C, &H13B60F14, &HC8B60F43, &HE8C1D133, &H97043308, &H7503C3F6, &H245C89E8, &HC7D58B1C, &H182444, &H2B000000, &HFCE283D3, &H831A0C8D, &HEAC103C2, &H1BCB3B02, &H23D1F7C9, &H244C89CA, &H8B797420, &H24A48DE9, &H0&, &HCB8B1B8B, &HF08E9C1, &HC88BF1B6, &HF08E9C1
    pvAppendBuffer &HF133C9B6, &HE9C1CB8B, &HD1B60F10, &HE9C1C88B, &HB7B48B10, &H800&, &H33C9B60F, &HC1C88BD1, &HB43318E9, &H40097, &HC1D38B00, &HD13318EA, &HFC8B60F, &H5C8BC3B6, &HC8331C24, &H3304C383, &H5C899734, &HB4331C24, &HC008F, &H244C8B00, &H41C68B18, &H18244C89, &H9475CD3B, &H10246C8B, &HF633CD8B, &HEB3BCB2B, &HD5F7ED1B, &H1674E923, &H8D13B60F, &HB60F015B, &HD13346C8, &H3308E8C1
    pvAppendBuffer &HF53B9704, &H4C8BEA75, &H5E5F2424, &H5B01895D, &H10C259, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H24748B56, &H7EF68510, &H24548B35, &H448B4E0C, &HEEC10824, &H2834604, &H750A8B01, &H442FF03, &H4A8B0889, &H4488904, &H840C7, &HC7000000, &HC40&, &HC0830000, &H1EE8310, &HC25ED875, &HCCCC0010, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &HC24548B, &H1A74D285, &H824448B, &H24748B56
    pvAppendBuffer &H8AF02B08, &H408D060C, &HFF483001, &H7501EA83, &H10C25EF2, &HCCCCCC00, &HCCCCCCCC, &HCCCCCCCC, &HC247C83, &H24448B00, &HC488B04, &H4244C89, &H8E8E0F, &H55530000, &H24748B56, &H7C8B5714, &HEF8B2024, &H8301E583, &H9B8D02E7, &H0&, &HED851E8A, &H488B1574, &H2C98308, &HF283D18B, &HD1AF0F01, &H3208EAC1, &H851688D3, &H8A0274FF, &H10B60F1E, &HCBB60F46, &H14245C8B, &H88BD133
    pvAppendBuffer &H8B08E9C1, &H588B9314, &HFD13308, &H4803CAB6, &H5C96904, &H89080884, &HD3B60F10, &H4108EBC1, &HF044889, &H330748B6, &H244C8BD1, &H911C3314, &H1C244C8B, &H8588949, &H1C244C89, &H907FC985, &H5B5D5E5F, &HCC0010C2, &HCCCCCCCC, &HCCCCCCCC, &H24748B56, &HC6A5708, &HFF28468B, &H57F88BD0, &H9E83789, &H8B000009, &H146A284E, &H40C7D1FF, &H10&, &HC40C700, &H0&, &H8B084789
    pvAppendBuffer &HC25E5FC7, &HCCCC0010, &HCCCCCCCC, &HCCCCCCCC, &H24748B56, &H3E8B5708, &H8B0876FF, &HD0FF3047, &H8B0476FF, &HD0FF3047, &H5630478B, &H5E5FD0FF, &HCC0010C2, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &HC24448B, &H24748B56, &H7C8B570C, &HC0850C24, &HFF500D74, &H36FF0476, &H21E837FF, &HFFFFFFFD, &H468D1C76, &H1876FF0C, &HFF1476FF, &H8D501076, &HFF500846, &H36FF0476, &H922E857, &H5E5F0000
    pvAppendBuffer &HCC0010C2, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H42474FF, &HFE7E8, &H10C200, &HCCCCCCCC, &H24748B56, &HC4E8B08, &H570C468D, &HC9853E8B, &H4E3B0C74, &H50077404, &H1182E856, &H4E8B0000, &H10468D10, &HC74C985, &H74084E3B, &HE8565007, &H116C&, &H147E83, &H7414468D, &HE8565007, &H115C&, &H5004468D, &H1152E856, &H468D0000, &HE8565008, &H1148&, &H5630478B, &H5E5FD0FF
    pvAppendBuffer &HCC0010C2, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H824448B, &H6C8B5553, &H57560C24, &H530C588D, &H5708788D, &HFF0470FF, &H21E85530, &H8B000009, &H8B1C244C, &H74C985F0, &H74F68518, &H85078B14, &H510C74C0, &HFF5033FF, &H31E80075, &H8BFFFFFC, &H5D5E5FC6, &H10C25B, &HCCCCCCCC, &HCCCCCCCC, &H24748B56, &HFFCA8308, &H8004868B, &HB70F0004, &H8D02C64C, &H3B66C604, &HF0974CA, &H8966C1BF
    pvAppendBuffer &H10EBC614, &H8304408B, &H874FFF8, &H46948966, &H48008, &H8004868B, &H4C8B0004, &H4C891024, &H868B04C6, &H48004, &HC6548966, &H4E148D02, &H80048E8B, &HB70F0004, &H4800882, &H4896600, &H4868BCE, &HF000480, &HFC60CBF, &H8966C0B7, &H4800882, &HFFF98300, &H8B660C74, &H4800486, &H44896600, &H8E8B02CE, &H48004, &HC24448A, &HE8488, &H8B000400, &H4800486, &HFF254000
    pvAppendBuffer &H8900007F, &H4800486, &HCC25E00, &HCCCCCC00, &H530CEC83, &H246C8B55, &H8B575620, &H3320247C, &H47F8BF6, &H18247C89, &H800CB739, &H8E0F0006, &HE2&, &H800C8F8B, &HC18B0006, &HC503C62B, &HF04F883, &H9C8C&, &H24548B00, &H7DF13B24, &H84B60F0A, &H680083E, &H8B08EB00, &HFC12BC6, &H8D1004B6, &H4488025E, &H438D1424, &H7DC13BFF, &H84B60F0A, &H680093E, &H8B09EB00, &HFC12BC6
    pvAppendBuffer &H11044B6, &H15244488, &HA7DD93B, &H3E84B60F, &H6800A, &HC68B09EB, &HB60FC12B, &H88021044, &H8D162444, &HC13B0143, &HB60F0A7D, &H800B3E84, &H9EB0006, &HC12BC68B, &H1044B60F, &H24448803, &H24448B17, &HBDC86914, &HC11E35A7, &H505110E9, &HFE8AE857, &H3B46FFFF, &H6800CB7, &H518C0F00, &HEBFFFFFF, &H3BCE8B31, &H6800CB7, &H8D277D00, &H6800897, &H8D09EB00, &H24A4&, &HFF8B0000
    pvAppendBuffer &H80F848A, &H8D000680, &H42880152, &H8F3B41FF, &H6800C, &HB729EA7C, &H6800C, &H8E0FED85, &H1F5&, &HA48D0AEB, &H24&, &H498D00, &H24245C8B, &HF04FD83, &H1598C, &H878B00, &H33000400, &H8789ED, &H69000480, &H35A7BD03, &H10E8C11E, &H4794BF0F, &H48008, &HFFFFA83, &H13184, &H48F8B00, &H8B000480, &HBEC12BC2, &H8000&, &H7FFF25, &H2BF02B00, &HFFE181CE
    pvAppendBuffer &H8B00007F, &H3984&, &H33B0004, &HB4890E75, &H68010AF, &H6C3B4500, &H97D2C24, &HD714BF0F, &H75FFFA83, &HFED85C0, &HE98E&, &H24448B00, &H24443928, &H89067E30, &HEB302444, &H24448B04, &H4BE30, &HC63B0000, &H9B8D717E, &H0&, &H331E048A, &H88DB33C9, &H85132444, &H90587EED, &H108F948B, &H8B000680, &H79C22BC6, &H4878B18, &H2B000480, &H25C603C2, &H7FFF&, &H38848A
    pvAppendBuffer &HEB000400, &H247C8B0B, &H38048A24, &H18247C8B, &H13244438, &H94890875, &H680109F, &H3B414300, &H83BD7CCD, &HF7E01FB, &H3024448B, &H8BEB8B46, &H3B24245C, &H8B997CF0, &H8B24245C, &H3B28246C, &H8B3D7DF5, &H6801097, &H2BCE8B00, &H1A2C8DCA, &H1879C985, &H8004878B, &HC22B0004, &HFF25C603, &H8A00007F, &H3884&, &H3EB0004, &H3819048A, &H8752904, &H743B4146, &HD47C2824, &H28246C8B
    pvAppendBuffer &H10B7FF56, &HFF000680, &HE8282474, &HEA0&, &H7D7EF685, &H4C8B38EB, &H38A2024, &H118B8F3C, &HF08718B, &H428BC8B6, &HF09773C, &H300844B6, &HDEB086A, &H804B60F, &H48D096A, &H145&, &H52565000, &H1247E8, &H1BE00, &H6C8B0000, &HFD832824, &H69167C04, &H35A7BD03, &H10E8C11E, &H3B60F50, &H75E85750, &HEBFFFFFC, &HC8F8B15, &H8A000680, &HF848803, &H68008, &H800C87FF
    pvAppendBuffer &H434E0006, &H7FF6854D, &H246C89C9, &H245C8928, &HFED8524, &HFFFE178F, &H5D5E5FFF, &HCC4835B, &HCC0014C2, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H5314EC83, &H246C8B55, &H8B575628, &H3328247C, &H47F8BF6, &H20247C89, &H800CB739, &H8E0F0006, &HE2&, &H800C8F8B, &HC18B0006, &HC503C62B, &HF04F883, &H9C8C&, &H24548B00, &H7DF13B2C, &H84B60F0A, &H680083E, &H8B08EB00, &HFC12BC6
    pvAppendBuffer &H8D1004B6, &H4488025E, &H438D1824, &H7DC13BFF, &H84B60F0A, &H680093E, &H8B09EB00, &HFC12BC6, &H11044B6, &H19244488, &HA7DD93B, &H3E84B60F, &H6800A, &HC68B09EB, &HB60FC12B, &H88021044, &H8D1A2444, &HC13B0143, &HB60F0A7D, &H800B3E84, &H9EB0006, &HC12BC68B, &H1044B60F, &H24448803, &H24448B1B, &HBDC86918, &HC11E35A7, &H505110E9, &HFB6AE857, &H3B46FFFF, &H6800CB7, &H518C0F00
    pvAppendBuffer &HEBFFFFFF, &H3BCE8B31, &H6800CB7, &H8D277D00, &H6800897, &H8D09EB00, &H24A4&, &HFF8B0000, &H80F848A, &H8D000680, &H42880152, &H8F3B41FF, &H6800C, &HB729EA7C, &H6800C, &HC933C033, &H14244489, &H18244C89, &H1C244489, &H8E0FED85, &H209&, &HA48D0AEB, &H24&, &H498D00, &H2C245C8B, &HF04FD83, &H1C08C, &H878B00, &H33000400, &H8789ED, &H69000480, &H35A7BD03
    pvAppendBuffer &H10E8C11E, &H4794BF0F, &H48008, &HFFFFA83, &H19884, &H48F8B00, &H8B000480, &HBEC12BC2, &H8000&, &H7FFF25, &H2BF02B00, &HFFE181CE, &H8B00007F, &H3984&, &H33B0004, &HB4890E75, &H68010AF, &H6C3B4500, &H97D3424, &HD714BF0F, &H75FFFA83, &HFED85C0, &H14C8E, &H24548B00, &H24448B30, &H7EC23B38, &H89C28B06, &HBE382444, &H4&, &H777EC63B, &H331E048A, &H88DB33C9
    pvAppendBuffer &H85132444, &HEB607EED, &H24A48D07, &H0&, &H108F948B, &H8B000680, &H79C22BC6, &H4878B18, &H2B000480, &H25C603C2, &H7FFF&, &H38848A, &HEB000400, &H247C8B0B, &H38048A2C, &H20247C8B, &H13244438, &H94890875, &H680109F, &H3B414300, &H83BD7CCD, &HF7E01FB, &H3824448B, &H8BEB8B46, &H3B2C245C, &H8B917CF0, &H8B2C245C, &H3B302454, &H8B397DF2, &H6801097, &H2BCE8B00, &H1A2C8DCA
    pvAppendBuffer &H1879C985, &H8004878B, &HC22B0004, &HFF25C603, &H8A00007F, &H3884&, &H3EB0004, &H3819048A, &H8752904, &H743B4146, &HD47C3024, &H14246C8B, &H5E7EED85, &H3B01458D, &H8B4F7EF0, &H8B282444, &H8688B10, &H1C24448B, &HB60F8F3C, &H3C428BC8, &HB60F0977, &H6A300844, &HF0DEB08, &H6A0804B6, &H45048D09, &H1&, &HE8525550, &HF20&, &H8B03B60F, &H680108F, &H244C8900, &H24748918
    pvAppendBuffer &H24448914, &H90E91C, &H448B0000, &H50551824, &H878B25EB, &H68010, &H18244489, &H8903B60F, &H89142474, &HEB1C2444, &H244C8B71, &H246C8B18, &H7EED8514, &HFF515532, &HE8302474, &HAF0&, &H758DC033, &H244489FF, &H7FF68514, &H246C8B52, &H244C8B30, &HFED8518, &HFFFE038F, &H5D5E5FFF, &H14C4835B, &H8B0014C2, &H8A28244C, &H8B8F3C03, &H8718B11, &H8BC8B60F, &H9773C42, &H844B60F
    pvAppendBuffer &HEB086A30, &H4B60F0D, &H8D096A08, &H14504, &H56500000, &HE76E852, &H1BE0000, &H8B000000, &H8330246C, &H167C04FD, &HA7BD0369, &HE8C11E35, &HB60F5010, &HE8575003, &HFFFFF8A4, &H8F8B15EB, &H6800C, &H8488038A, &H680080F, &HC87FF00, &H4E000680, &HF6854D43, &H6C89C97F, &H5C893024, &H6AE92C24, &HCCFFFFFF, &HCCCCCCCC, &HCCCCCCCC, &H7C8B5756, &HB0680C24, &H8B00068F, &H28408B07
    pvAppendBuffer &HF08BD0FF, &H575F685, &H4C25E5F, &H4778900, &HBA024E8D, &H8000&, &HEBFFC883, &H498D03, &HFF0241C7, &H8DFFFFFF, &H89660849, &H8966F841, &HEA83F641, &H8DE97501, &H48008BE, &H8000B900, &HABF30000, &H496895F, &H8D000480, &H96890142, &H6800C, &H4C25E, &HCCCCCCCC, &HCCCCCCCC, &H7C8B5756, &H12B90C24, &HBE000000, &H55E75000, &HE72990B8, &H8BA5F355, &H8310244C, &H5E5F60C0
    pvAppendBuffer &H100001C7, &H418955E7, &H841C704, &H55E740A0, &H780C41C7, &HC255E744, &HCCCC0008, &HCCCCCCCC, &H245C8B53, &H68575608, &H8000&, &H738B3B8B, &H28478B08, &H689D0FF, &H4439C033, &H36A2424, &HC7C0950F, &H80000846, &HC0830000, &H446C702, &H0&, &HE8575650, &HD34&, &H28247C83, &H2474FF00, &H2474FF30, &H2474FF30, &H2474FF20, &H7745320, &HFFF817E8, &HE805EBFF, &HFFFFFB30
    pvAppendBuffer &H6A076A, &H5E85756, &H8300000D, &H24247C, &H4E8B1874, &H74C98510, &H8B811, &HC12B0000, &H56006A50, &HCE6E857, &H448B0000, &HE8B1C24, &H8B08895F, &H8B1C2444, &H5B5E044E, &H1B80889, &HC2000000, &HCCCC0024, &HCCCCCCCC, &HCCCCCCCC, &H56555351, &H1424748B, &H80006857, &H1E8B0000, &H1C245C89, &HFF28438B, &H246C8BD0, &HA4868920, &HC7000081, &H81AC86, &H800000, &HA886C700
    pvAppendBuffer &H81&, &H85000000, &H830D7FED, &H19CBE, &H8E0F0000, &H5B6&, &H19CBE83, &H8D180000, &H19CBE, &H8B2A7D00, &HEB1C2454, &H498D03, &H197EED85, &H4D02B60F, &HD3420F8B, &H988609E0, &H8D000001, &H7890841, &H7C18F883, &H245489E3, &H18468B1C, &H975C085, &H11846C7, &HEB000000, &H1F883A6, &H178B7375, &HF03FA83, &H55D8C, &H98868B00, &H83000001, &HE8D1FDC2, &H1789C88B
    pvAppendBuffer &H8902E8C1, &H19886, &H3E18300, &HCA8B1D75, &HA1846C7, &H83000000, &HD12B07E1, &H1789E8D3, &H1988689, &H5CE90000, &H83FFFFFF, &H187501F9, &H8904468B, &H468B0C46, &H10468908, &H61846C7, &HE9000000, &HFFFFFF3F, &HF02F983, &HFFFF3685, &H184E89FF, &HFFFF2EE9, &H2F883FF, &H178B6E75, &HF0EFA83, &H4E58C, &H988E8B00, &H8B000001, &H1FE083C1, &H505E9C1, &H101&, &H2C46C7
    pvAppendBuffer &H89000000, &HC18B2046, &HC11FE083, &HC74005E9, &H31846, &H46890000, &H83C18B24, &HE9C10FE0, &H4C08304, &H1988E89, &H46890000, &HF2428D28, &HC0330789, &H89444689, &H46894846, &H5046894C, &H54468966, &HE9564688, &HFFFFFEBB, &H7503F883, &HF07396F, &H4758C, &H2C468B00, &H7D28463B, &H33F8331, &H4B8B2C7C, &H2C468B40, &H198968A, &HE2800000, &H4B60F07, &H30548801, &H2C46FF44
    pvAppendBuffer &H832C468B, &HAEC1FD07, &H198&, &H28463B03, &H468BCF7C, &H28463B2C, &HFE69850F, &H136AFFFF, &H5044468D, &H9C6E856, &H46890000, &H1846C714, &H4&, &H2C46C7, &HE9000000, &HFFFFFE47, &HF04F883, &HE285&, &H24468B00, &H3204E8B, &H2C4639C1, &H8D513F7C, &H56505746, &H98FE8, &H2476FF00, &H8B0C4689, &HC0832046, &H50C60357, &H97AE856, &H7E8D0000, &H10468914, &HBDE85657
    pvAppendBuffer &HC7000005, &H7&, &H1846C700, &H6&, &HFFFDF2E9, &H1476FFFF, &H198868D, &H50570000, &H61BE8, &H83C88B00, &H840FFFF9, &H39E&, &HFFEF983, &H36284, &H10F98300, &H468B0F7D, &H304C882C, &H2C46FF57, &HFFFDBAE9, &HB80775FF, &H2&, &HC0330FEB, &HF11F983, &H48DC095, &H385&, &H30468900, &HF983C033, &HC0940F12, &H3C5048D, &H89000000, &HF9833446, &H8B1B7510
    pvAppendBuffer &HC0852C46, &HB60F147E, &H89563044, &H46C73C46, &H518&, &HFD6DE900, &HC033FFFF, &H51846C7, &H89000000, &H5CE93C46, &H83FFFFFD, &H5C7505F8, &H4E8B1F8B, &HFD93B30, &H3118C, &H98868B00, &HBA000001, &H1&, &HD92BE2D3, &H231F894A, &H3E8D3D0, &H86893456, &H198&, &H1C7ED285, &H324468B, &H4E8B2046, &H7DC83B2C, &H3C468A0F, &H3144884A, &H2C46FF57, &HE47FD285, &H18245C8B
    pvAppendBuffer &H41846C7, &HE9000000, &HFFFFFCFB, &HF06F883, &H9C85&, &HC76FF00, &H570C5E8D, &H198868D, &HE8500000, &H518&, &HFFFF883, &H29D84, &HFEF88300, &H25D840F, &H3D0000, &H7D000001, &HE8565010, &H408&, &H18245C8B, &HFFFCB6E9, &H8B4075FF, &H1846C703, &H1&, &H7404463B, &HE856530D, &H458&, &H3C7&, &H468B0000, &H107E8D10, &H18245C8B, &HF08463B, &HFFFC8684
    pvAppendBuffer &HE85657FF, &H438&, &H7C7&, &H74E90000, &H8BFFFFFC, &H3D18245C, &H11E&, &HFC658D0F, &H46C7FFFF, &H718&, &H1C468900, &HFFFC56E9, &H7F883FF, &H468B5475, &H1012D1C, &HC8D0000, &H34438B40, &HF881C8D, &H39024BBF, &HFB8C0F0F, &H8B000001, &H19896, &H1B800, &HE0D30000, &H3C22348, &H46890443, &H43BF0F38, &H8A072902, &H5C8B024B, &HEAD31824, &H1989689, &H46C70000
    pvAppendBuffer &H818&, &HFBFDE900, &HF883FFFF, &HFF3A7508, &H868D1076, &H198&, &H21E85057, &H83000004, &H840FFFF8, &H1A6&, &HFFEF883, &H16A84, &H1EF88300, &H1618D0F, &H46C70000, &H918&, &H1C468900, &HFFFBBEE9, &H9F883FF, &H8F850F, &H468B0000, &H8D178B1C, &H438B400C, &H88048D38, &H248BF0F, &H20244489, &H10244C89, &H8C0FD13B, &H15A&, &H198868B, &H1BB0000, &H2B000000
    pvAppendBuffer &HD3102454, &H244C8BE3, &HD8234B20, &H89045903, &H2498A17, &H7E83E8D3, &H86890038, &H198&, &H61846C7, &H74000000, &H24A48D2C, &H0&, &H81A0868B, &H4EFF0000, &H25C32B38, &H7FFF&, &H3084B60F, &H1A0&, &H81E85650, &H83000002, &H7500387E, &H384EFFDB, &H18245C8B, &HFFFB26E9, &HAF883FF, &H178B3175, &HF10FA83, &HDD8C&, &H988E8B00, &HF000001, &H4689C1B7, &HF0428D40
    pvAppendBuffer &H8910E9C1, &H988E8907, &HC7000001, &HB1846, &HF0E90000, &H83FFFFFA, &H3F750BF8, &HF883078B, &HA78C0F10, &H8B000000, &H1988E, &HF0C08300, &H89D1B70F, &HFFF28107, &H8B0000FF, &HE9C14046, &H988E8910, &H3B000001, &HF75075C2, &H83C01BD8, &H89400BE0, &HACE91846, &H83FFFFFA, &H850F0CF8, &HFFFFFAA3, &H7C083F83, &H86B60F65, &H198&, &HD9E85650, &H83000001, &HAEC1F807, &H198&
    pvAppendBuffer &H40468308, &H7C850FFF, &HC7FFFFFA, &H11846, &H70E90000, &H8BFFFFFA, &HFF18245C, &H81A4B6, &H30438B00, &H448BD0FF, &H86C72424, &H81A4&, &H0&, &HC75E5F, &H0&, &H2024448B, &HC75B5D, &H0&, &HC259C033, &H4C8B0014, &H1B82424, &H8B000000, &H81A496, &H11895F00, &H81A8968B, &H4C8B0000, &H5D5E2424, &H5911895B, &HCC0014C2, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC
    pvAppendBuffer &H120EC81, &H56530000, &H2C24B48B, &H57000001, &H81B068, &H28468B00, &HD88BD0FF, &HC247C8D, &H80808B8, &H24B908, &H20680000, &H89000001, &HB8ABF333, &H9090909, &HA024BC8D, &HB9000000, &H1C&, &H448DABF3, &H84C71024, &H11024, &H7070700, &HC7535007, &H11C2484, &H7070000, &H84C70707, &H12024, &H7070700, &H2484C707, &H124&, &H7070707, &H282484C7, &H7000001
    pvAppendBuffer &HC7070707, &H12C2484, &H7070000, &H84C70707, &H13024, &H8080800, &H2484C708, &H134&, &H8080808, &H4C7E8, &H4438900, &HC24448D, &H5350206A, &H182444C7, &H5050505, &H1C2444C7, &H5050505, &H202444C7, &H5050505, &H242444C7, &H5050505, &H282444C7, &H5050505, &H2C2444C7, &H5050505, &H302444C7, &H5050505, &H342444C7, &H5050505, &H477E8, &H43895F00, &H5EC38B08
    pvAppendBuffer &H1843C7, &HC7000000, &H1443&, &H43C70000, &H10&, &HC43C700, &H0&, &H19883C7, &H0&, &H83C70000, &H19C&, &H0&, &H81A083C7, &H0&, &H815B0000, &H120C4, &H4C200, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H245C8B53, &H748B560C, &H868B0C24, &H81A0&, &H9C880E8B, &H1A030, &HA0868B00, &H40000081, &H7FFF25, &HA0868900, &H8B000081, &H81A886
    pvAppendBuffer &HAC863B00, &H7C000081, &H50C0031A, &H81A4B6FF, &H86890000, &H81AC&, &HFF2C418B, &HA48689D0, &H8B000081, &H81A886, &HA48E8B00, &H88000081, &H86FF011C, &H81A8&, &H8C25B5E, &HCCCCCC00, &H245C8B53, &H6C8B5508, &H38B1024, &HC244489, &H875ED85, &HFFC8835D, &H8C25B, &H7D8B57, &H4B74FF85, &H39F63356, &HEB237C37, &H498D03, &H8D04478B, &HC083F004, &H388304, &H53500774
    pvAppendBuffer &HFFFFBBE8, &H373B46FF, &H448BE67E, &H77FF1424, &H30408B04, &H448BD0FF, &H47C71424, &H4&, &H408B5700, &HC7D0FF30, &H45&, &H5F5E0000, &H5BC0335D, &HCC0008C2, &HCCCCCCCC, &HCCCCCCCC, &HC24448B, &H245C8B53, &H88B560C, &H8B04408B, &H7C8B5733, &H178B1024, &H48DCA23, &H8B60FC8, &H267FCE3B, &HD308B60F, &H66F12BEA, &HFF027883, &H408B2C75, &H74C08504, &H8B088B1A, &HCA230440
    pvAppendBuffer &HFC8048D, &HCE3B08B6, &H5E5FDA7E, &H5BFFC883, &H5F000CC2, &HFFFEB85E, &HC25BFFFF, &H1789000C, &HBF0F3389, &H5E5F0240, &HCC25B, &HCCCCCCCC, &H8B08EC83, &H550C2444, &H408B288B, &H24448908, &H24448B10, &H246C8918, &HFC08504, &H12D8E, &H57565300, &H1043D, &HBE077E00, &H102&, &HF08B0CEB, &H1023D, &H8D037E00, &HC62BFD70, &H1DBB&, &H24448900, &HFFCF8324, &H8B34458B
    pvAppendBuffer &H244489E8, &H3B048D14, &HD1C22B99, &H400C8DF8, &H48D743B, &HD88B047D, &H743BEBEB, &H47E088D, &HE1EBF88B, &H1C245C8B, &H8D40048D, &H853C&, &H17B90000, &HF000001, &H32F04B7, &H246C8BFD, &HC13B6610, &H8BC8BF0F, &HC7F3C45, &H4884B60F, &HFFFFFE00, &H7EB076A, &H144B60F, &H50086AA8, &H29E85553, &HF000003, &H660247B7, &HD74C085, &H9804772B, &H55535650, &H313E8, &H38458B00
    pvAppendBuffer &H8BFFCE83, &HBF20245C, &H1E&, &HFF8BE88B, &H9937048D, &HF8D1C22B, &H3B400C8D, &H7D048D5C, &HEBF88B04, &H8D5C3BEB, &H8B047E08, &H8BE1EBF0, &H8D1C247C, &H348D4004, &H85&, &H8BF50300, &H6A10246C, &HEBF0F05, &HF3C458B, &H50C804B6, &HB9E85557, &HF000002, &H660246B7, &HF74C085, &HC38B5098, &H5004462B, &HA1E85557, &H8B000002, &H85242444, &HD98F0FC0, &H5FFFFFFE, &H835D5B5E
    pvAppendBuffer &HCC208C4, &HCCCCCC00, &HCCCCCCCC, &HCCCCCCCC, &H6C8B5553, &H57560C24, &H458B086A, &H8BD0FF28, &H8B28245C, &HB8CB8BF8, &H1&, &H1BEE0D3, &H8B000000, &H482C244C, &H4489E6D3, &H48D2824, &HF5&, &H458B5000, &H89D0FF28, &HC9330447, &H83FF468D, &H789FFCA, &H2378C085, &H24648D, &H6604478B, &H2C85489, &HC604478B, &H8B00C804, &H44C70447, &H4C8&, &H3B410000, &H33E17E0F
    pvAppendBuffer &H246C39ED, &H8B697E20, &HEB1C2474, &H498D03, &H2E04B60F, &H517EC33B, &H1824448B, &H8BA8048B, &H244C23C8, &H244C3B28, &H8B3E7524, &H23F8D3CB, &H7F073B07, &H44F8B34, &HC16C8966, &H44F8B02, &H2E14B60F, &H348DD32B, &HEB60FC1, &H27DCA3B, &H748B1688, &H1BA1C24, &HF000000, &H2B2E0CB6, &H3E2D3CB, &H7E073BC2, &H6C3B45CC, &HA07C2024, &H3739F633, &H548B677C, &H6C8B2C24, &HFF8B1C24
    pvAppendBuffer &H8D04478B, &HB60FF00C, &H7EC23B01, &HFFC8834B, &H2418966, &H8D04478B, &HB60FF00C, &H83C22B01, &H57E07F8, &H7B8&, &H48D5000, &H50118813, &HC68BCB8B, &H440BE0D3, &HFF502C24, &H552C2474, &H2C2474FF, &H2C2474FF, &HFFFEC3E8, &H44F8BFF, &H2C24548B, &H4F14489, &H7E373B46, &H5FC78BA3, &HC25B5D5E, &HCCCC001C, &HCCCCCCCC, &HCCCCCCCC, &H504EC81, &HC0330000, &H8B55D233, &H51424AC
    pvAppendBuffer &H54890000, &H44890424, &H44890C24, &H44891024, &H44891424, &H44891824, &H44891C24, &H44892024, &H44892424, &H44892824, &H44892C24, &H44893024, &H44893424, &H44893824, &H44893C24, &H44894024, &H8B564424, &H51424B4, &HED850000, &HFF8B197E, &H300CB60F, &HC8C44FF, &H27DD13B, &H3B40D18B, &H89ED7CC5, &H33082454, &HEBC933D2, &H498D03, &H500C5489, &H100C5403, &H304C183, &H3CF983D2
    pvAppendBuffer &H3357EE7C, &H7EED85FF, &H9C8D5349, &H9424&, &H34B60F00, &H3C737, &H8B000000, &H8D54B454, &H44890142, &HF68554B4, &HC0331874, &H24648D, &HC003CA8B, &HD101E183, &H83C10BFA, &HF07501EE, &HB48B0389, &H51C24, &HC3834700, &H7CFD3B04, &H448B5BC0, &H835F0C24, &H57C09F8, &H9B8&, &H6A5000, &HFF55006A, &H52424B4, &H848D0000, &HA024&, &H848B5000, &H52824, &HE830FF00
    pvAppendBuffer &HFFFFFD8C, &HC4815D5E, &H504&, &HCC000CC2, &HC24448B, &H24748B56, &H104E8B0C, &H4609E0D3, &H24448B0C, &H89C10314, &HF8831046, &H57397C08, &HC247C8B, &H3908468B, &HF7C0446, &HFF50C003, &H8468936, &HFF2C478B, &H8B0689D0, &HE8B0456, &H880C468A, &H46FF0A04, &HC6EC104, &H10468308, &H107E83F8, &H5FCD7D08, &H10C25E, &H8BEC8B55, &HC0408000, &HE060A020, &HD0509010, &HF070B030
    pvAppendBuffer &HC8488808, &HE868A828, &HD8589818, &HF878B838, &HC4448404, &HE464A424, &HD4549414, &HF474B434, &HCC4C8C0C, &HEC6CAC2C, &HDC5C9C1C, &HFC7CBC3C, &HC2428202, &HE262A222, &HD2529212, &HF272B232, &HCA4A8A0A, &HEA6AAA2A, &HDA5A9A1A, &HFA7ABA3A, &HC6468606, &HE666A626, &HD6569616, &HF676B636, &HCE4E8E0E, &HEE6EAE2E, &HDE5E9E1E, &HFE7EBE3E, &HC1418101, &HE161A121, &HD1519111, &HF171B131
    pvAppendBuffer &HC9498909, &HE969A929, &HD9599919, &HF979B939, &HC5458505, &HE565A525, &HD5559515, &HF575B535, &HCD4D8D0D, &HED6DAD2D, &HDD5D9D1D, &HFD7DBD3D, &HC3438303, &HE363A323, &HD3539313, &HF373B333, &HCB4B8B0B, &HEB6BAB2B, &HDB5B9B1B, &HFB7BBB3B, &HC7478707, &HE767A727, &HD7579717, &HF777B737, &HCF4F8F0F, &HEF6FAF2F, &HDF5F9F1F, &HFF7FBF3F, &H101&, &H3&, &H3&, &H102&
    pvAppendBuffer &H4&, &H4&, &H103&, &H5&, &H5&, &H104&, &H6&, &H6&, &H105&, &H7&, &H7&, &H106&, &H8&, &H8&, &H107&, &H9&, &H9&, &H108&, &HA&, &HA&, &H10109, &HB&, &HC&, &H1010A, &HD&, &HE&, &H1010B, &HF&, &H10&, &H1010C, &H11&, &H12&
    pvAppendBuffer &H2010D, &H13&, &H16&, &H2010E, &H17&, &H1A&, &H2010F, &H1B&, &H1E&, &H20110, &H1F&, &H22&, &H30111, &H23&, &H2A&, &H30112, &H2B&, &H32&, &H30113, &H33&, &H3A&, &H30114, &H3B&, &H42&, &H40115, &H43&, &H52&, &H40116, &H53&, &H62&, &H40117, &H63&
    pvAppendBuffer &H72&, &H40118, &H73&, &H82&, &H50119, &H83&, &HA2&, &H5011A, &HA3&, &HC2&, &H5011B, &HC3&, &HE2&, &H5011C, &HE3&, &H101&, &H11D&, &H102&, &H102&, &H121110, &H6090708, &H40B050A, &H20D030C, &HF010E, &H0&, &H1&, &H1&, &H1&, &H2&, &H2&, &H2&, &H3&
    pvAppendBuffer &H3&, &H3&, &H4&, &H4&, &H10004, &H5&, &H6&, &H10005, &H7&, &H8&, &H20006, &H9&, &HC&, &H20007, &HD&, &H10&, &H30008, &H11&, &H18&, &H30009, &H19&, &H20&, &H4000A, &H21&, &H30&, &H4000B, &H31&, &H40&, &H5000C, &H41&, &H60&, &H5000D
    pvAppendBuffer &H61&, &H80&, &H6000E, &H81&, &HC0&, &H6000F, &HC1&, &H100&, &H70010, &H101&, &H180&, &H70011, &H181&, &H200&, &H80012, &H201&, &H300&, &H80013, &H301&, &H400&, &H90014, &H401&, &H600&, &H90015, &H601&, &H800&, &HA0016, &H801&, &HC00&, &HA0017, &HC01&, &H1000&
    pvAppendBuffer &HB0018, &H1001&, &H1800&, &HB0019, &H1801&, &H2000&, &HC001A, &H2001&, &H3000&, &HC001B, &H3001&, &H4000&, &HD001C, &H4001&, &H6000&, &HD001D, &H6001&, &H8000&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&
    '--- end thunk data
    ReDim baBuffer(0 To 7624 - 1) As Byte
    Call CopyMemory(baBuffer(0), m_baBuffer(0), UBound(baBuffer) + 1)
    Erase m_baBuffer
End Sub

'= shared ================================================================

#If ImplUseShared = 0 Then

Private Function FileAttribs(sFile As String) As VbFileAttribute
    FileAttribs = GetFileAttributes(StrPtr(sFile))
    If FileAttribs = -1 Then
        FileAttribs = &H8000&
    End If
End Function

Private Function MkPath(sPath As String, sError As String) As Boolean
    Dim lAttrib         As Long
    
    lAttrib = GetFileAttributes(StrPtr(sPath))
    If lAttrib = -1 Then
        If InStrRev(sPath, "\") > 0 Then
            If Not MkPath(Left$(sPath, InStrRev(sPath, "\") - 1), sError) Then
                Exit Function
            End If
        End If
        If CreateDirectory(StrPtr(sPath), 0) = 0 Then
            sError = GetSystemMessage(Err.LastDllError) & " (" & sPath & ")"
            Exit Function
        End If
    ElseIf (lAttrib And vbDirectory + vbVolume) = 0 Then
        sError = "File already exists (" & sPath & ")"
        Exit Function
    End If
    '--- success
    MkPath = True
End Function

Private Function GetSystemMessage(ByVal lLastDllError As Long) As String
    Dim lSize            As Long
   
    GetSystemMessage = String$(2000, 0)
    lSize = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, lLastDllError, 0, StrPtr(GetSystemMessage), Len(GetSystemMessage), 0)
    If lSize > 2 Then
        If Mid$(GetSystemMessage, lSize - 1, 2) = vbCrLf Then
            lSize = lSize - 2
        End If
    End If
    GetSystemMessage = Left$(GetSystemMessage, lSize) & " &H" & Hex$(lLastDllError)
End Function

Private Sub AssignVariant(vDest As Variant, vSrc As Variant)
    If IsObject(vSrc) Then
        Set vDest = vSrc
    Else
        vDest = vSrc
    End If
End Sub

Private Function PathCombine(sPath As String, sFile As String) As String
    PathCombine = sPath & IIf(LenB(sPath) <> 0 And Right$(sPath, 1) <> "\" And LenB(sFile) <> 0, "\", vbNullString) & sFile
End Function

Private Function At(vArray As Variant, ByVal lIdx As Long) As Variant
    On Error GoTo QH
    At = vArray(lIdx)
QH:
End Function

Private Function Clamp( _
            ByVal lValue As Long, _
            Optional ByVal Min As Long = -2147483647, _
            Optional ByVal Max As Long = 2147483647) As Long
    If lValue < Min Then
        Clamp = Min
    ElseIf lValue > Max Then
        Clamp = Max
    Else
        Clamp = lValue
    End If
End Function

Private Function Peek(ByVal lPtr As Long) As Long
    Call CopyMemory(Peek, ByVal lPtr, 4)
End Function

Private Function PeekInt(ByVal lPtr As Long) As Integer
    Call CopyMemory(PeekInt, ByVal lPtr, 2)
End Function

Public Function ToUtf8Array(sText As String) As Byte()
    Dim baRetVal()      As Byte
    Dim lSize           As Long
    
    lSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), 0, 0, 0, 0)
    If lSize > 0 Then
        ReDim baRetVal(0 To lSize - 1) As Byte
        Call WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), VarPtr(baRetVal(0)), lSize, 0, 0)
    Else
        baRetVal = vbNullString
    End If
    ToUtf8Array = baRetVal
End Function

Public Function FromUtf8Array(baText() As Byte) As String
    Dim lSize           As Long
    
    If UBound(baText) >= 0 Then
        FromUtf8Array = String$(2 * (UBound(baText) + 1), 0)
        lSize = MultiByteToWideChar(CP_UTF8, 0, baText(0), UBound(baText) + 1, ByVal StrPtr(FromUtf8Array), Len(FromUtf8Array))
        FromUtf8Array = Left$(FromUtf8Array, lSize)
    End If
End Function

Private Function UnsignedAdd(ByVal lUnsignedPtr As Long, ByVal lSignedOffset As Long) As Long
    '--- note: safely add *signed* offset to *unsigned* ptr for *unsigned* retval w/o overflow in LARGEADDRESSAWARE processes
    UnsignedAdd = ((lUnsignedPtr Xor &H80000000) + lSignedOffset) Xor &H80000000
End Function

Private Function C_Obj(Value As Variant) As Object
    On Error Resume Next
    Set C_Obj = Value
    On Error GoTo 0
End Function

Private Function SanitizeFileName(sFileName As String, Optional InvalidChar As String = " ") As String
    Const STR_INVALID   As String = "[\x00-\x1F""<>\|:\*\?\\/]"
    
    With InitRegExp(STR_INVALID)
        SanitizeFileName = .Replace(sFileName, InvalidChar)
    End With
End Function

Private Function InitRegExp(sPattern As String) As Object
    Static oCache       As Object
    
    If Not oCache Is Nothing Then
        If oCache.Pattern = sPattern Then
            GoTo QH
        End If
    End If
    Set oCache = VBA.CreateObject("VBScript.RegExp")
    oCache.Global = True
    oCache.Pattern = sPattern
QH:
    Set InitRegExp = oCache
End Function

#End If ' Not ImplUseShared

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    pvInitRelocTable m_uRtbl
    m_lCurrentFile = -1
    m_lCodePage = CP_OEMCP
End Sub

#If ImplCrypto Then
    Private Sub Class_Terminate()
        pvCryptoTerminate
    End Sub
#End If ' ImplCrypto

