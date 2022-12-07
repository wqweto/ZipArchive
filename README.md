## cZipArchive
A single-class pure VB6 library for zip archives management

### Usage

Just include `cZipArchive.cls` to your project and start using instances of the class like this:

#### Simple compression

    With New cZipArchive
        .AddFile App.Path & "\your_file"
        .CompressArchive App.Path & "\test.zip"
    End With
    
#### Compress all files and sub-folders

    With New cZipArchive
        .AddFromFolder "C:\Path\To\*.*", Recursive:=True
        .CompressArchive App.Path & "archive.zip"
    End With

#### Decompress all files from archive

    With New cZipArchive
        .OpenArchive App.Path & "\test.zip"
        .Extract "C:\Path\To\extract_folder"
    End With
    
Method `Extract` can optionally filter on file mask (e.g. `Filter:="*.doc"`), file index (e.g. `Filter:=15`) or array of booleans with each entry to decompress index set to `True`.

#### Extract single file to target filename

`OutputTarget` can include a target `new_filename` to be used when extracting a specific file from the archive.

    With New cZipArchive
        .OpenArchive App.Path & "\test.zip"
        .Extract "C:\Path\To\extract_folder\new_filename", Filter:="your_file"
    End With
    
#### Get archive entry uncompressed size

By using `FileInfo` property keyed on entry filename in first parameter and `zipIdxSize` like this

    With New cZipArchive
        .OpenArchive App.Path & "\test.zip"
        Debug.Print .FileInfo("report.pdf", zipIdxSize)
    End With
    
#### List files in zip archive

By using `FileInfo` propery keyed on entry numeric index in first parameter like this

    Dim lIdx            As Long
    With New cZipArchive
        .OpenArchive App.Path & "\test.zip"
        For lIdx = 0 To .FileCount - 1
            Debug.Print "FileName=" & .FileInfo(lIdx, zipIdxFileName) & ", Size=" & .FileInfo(lIdx, zipIdxSize)
        Next
    End With

Here is a list of available values for the second parameter of `FileInfo`:

<span> | Name
---    | ---
`0`    | `zipIdxFileName`
`1`    | `zipIdxAttributes`
`2`    | `zipIdxCrc32`
`3`    | `zipIdxSize`
`4`    | `zipIdxCompressedSize`
`5`    | `zipIdxComment`
`6`    | `zipIdxLastModified`
`7`    | `zipIdxMethod`
`8`    | `zipIdxOffset`
`9`    | `zipIdxFlags`

#### Encryption support

Make sure to set Conditional Compilation in Make tab in project's properties dialog to include `ZIP_CRYPTO = 1` setting for crypto support to get compiled from sources. By default crypto support is not compiled to reduce footprint on the final executable size.

    With New cZipArchive
        .OpenArchive App.Path & "\test.zip"
        .Extract App.Path & "\test", Password:="123456"
    End With
    
Use `Password` parameter on `AddFile` method together with `EncrStrength` parameter to set crypto used when creating archive.

EncrStrength | Mode
---          | ---
`0`          | ZipCrypto (default)
`1`          | AES-128
`2`          | AES-192
`3`          | AES-256 (recommended)

Note that default ZipCrypto encryption is weak but this is the only option which is compatible with Windows Explorer built-in zipfolders support.

#### In-memory operations

Sample utility function `ReadBinaryFile` in `/test/basic/Form1.frm` returns byte array with file's content. 

    Dim baZip() As Byte
    With New cZipArchive
        .AddFile ReadBinaryFile("sample.pdf"), "report.pdf"
        .CompressArchive baZip
    End With
    WriteBinaryFile "test.zip", baZip

Method `Extract` accepts byte array target too.
    
    Dim baOutput() As Byte
    With New cZipArchive
        .OpenArchive ReadBinaryFile("test.zip")
        .Extract baOutput, Filter:=0    '--- archive's first file only
    End With
    
### ToDo (not supported yet)

    - Deflate64 (de)compressor
    - VBA7 (x64) support
