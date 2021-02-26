## cZipArchive
A single-class pure VB6 library for zip archives management

### Simple compress

    With New cZipArchive
        .AddFile "your_file"
        .CompressArchive "test.zip"
    End With
    
### Compress all files and sub-folders

    With New cZipArchive
        .AddFromFolder "C:\Path\To\*.*", Recursive:=True
        .CompressArchive "archive.zip"
    End With

### Extract

    With New cZipArchive
        .OpenArchive "test.zip"
        .Extract "extract_folder"
    End With
    
Method `Extract` can optionally filter on file mask (e.g. `Filter:="*.doc"`), file index (e.g. `Filter:=15`) or array of booleans with each entry to decompress index set to `True`.

### Password support

Make sure to set Conditional Compilation in Make tab in project's properties dialog to include `ZIP_CRYPTO = 1` setting for crypto support to get compiled from sources. By default crypto support is not compiled to reduce footprint on the final executable size.

    With New cZipArchive
        .OpenArchive App.Path & "\test.zip"
        .Extract App.Path & "\test", Password:="123456"
    End With
    
Use `Password` parameter on `AddFile` method together with `EncrStrength` parameter to set crypto used when creating archive. Use `EncrStrength:=0` for ZipCrypto (default), `EncrStrength:=1` for AES-128, `EncrStrength:=2` for AES-192, etc.

Note that default ZipCrypto encryption is weak but this is the only option which is compatible with Windows Explorer built-in zip folders support.

### In-memory operations

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
    - Zip64 format extension (file sizes above 2GB+)
    - VBA7 (x64) support
