''' <summary>
''' Specifies how the solid memory block is configured when creating an archive.
''' </summary>
''' <remarks>
''' Disabled and Enabled map to 7-Zip solid mode on/off behavior.
''' Size-based values (Bytes, KiB, MiB, GiB, TiB) require a positive size and
''' specify the solid block size using the selected unit.
''' </remarks>
<Description("Defines the solid memory block mode used by the archive. Can disable solid mode, enable it using defaults, or specify an explicit block size in bytes or binary units.")>
Public Enum MemoryBlockMode
    Disabled = 0
    Enabled = 1
    Bytes = 10
    KiB = 11
    MiB = 12
    GiB = 13
    TiB = 14
End Enum


''' <summary>
''' Controls output verbosity for archive operations.
''' </summary>
<Description("Controls output verbosity for 7-Zip archive operations.")>
Public Enum SevenZipOutputLogLevel
    ''' <summary>No logging (default).</summary>
    <Description("Disable log (default).")>
    Disabled = 0

    ''' <summary>Show processed file names.</summary>
    <Description("Show names of processed files.")>
    Files = 1

    ''' <summary>Show internally processed files.</summary>
    <Description("Show names of internally processed files in solid archives.")>
    InternalFiles = 2

    ''' <summary>Show additional operations.</summary>
    <Description("Show additional operations such as Analyze and Replicate.")>
    Operations = 3
End Enum