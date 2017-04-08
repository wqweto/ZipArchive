## cZipArchive
A single-class pure VB6 library for zip archives management

### Simple compress

    With New cZipArchive
        .AddFile "your_file"
        .CompressArchive "test.zip"
    End With

### Extract

    With New cZipArchive
        .OpenArchive "test.zip"
        .Extract "extract_folder"
    End With
    
### ToDo (not supported yet)

    - Deflate64 (de)compressor
    - Zip64 format extension (sizes above 2GB+, unicode filenames)
