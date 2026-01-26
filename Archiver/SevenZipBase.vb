Imports System.IO
Imports System.Text

Public Interface ISevenZipBase
    <Description("Get/Set the output log level")>
    Property OutputLogLevel As SevenZipOutputLogLevel
    <Description("Sets level of compression. Valid values are 0 (no compression) to 9 (maximum compression).")>
    Property CompressionLevel As Integer
    <Description("Sets level of analysis. Valid values are 0 (no analysis) to 9 (maximum analysis).")>
    Property AnalysisLevel As Integer

    <Description("Adds a single file to the archive.")>
    Sub AddFile(ByVal Filepath As String)
    <Description("Adds multiple files to the archive using an Array(...). If {BaseDirectory} is supplied, each file will be prefixed with this path.")>
    Sub AddFiles(ByVal FilepathArr As Object, Optional BaseDirectory As String = Nothing)
    Sub SetMemoryBlockSize(Mode As MemoryBlockMode, Optional Size As Integer = 0)

End Interface

Public MustInherit Class SevenZipBase

    Implements ISevenZipBase

    Protected _archivePath As String
    Protected _analysisLevel As Integer = 5
    Protected _compressionLevel As Integer = 5
    Protected _outputLogLevel As SevenZipOutputLogLevel = SevenZipOutputLogLevel.Disabled
    Protected _filesToAdd As New List(Of String)
    Protected _memoryBlock As New MemoryBlock

    Protected Sub New(ByVal ArchivePath As String)
        _archivePath = ArchivePath
    End Sub

    ''' <summary>
    ''' Sets the output log level for Delete, Add, Update, and Extract operations.
    ''' </summary>
    ''' <returns>An integer representing the selected log verbosity level.</returns>
    Public Property OutputLogLevel As SevenZipOutputLogLevel Implements ISevenZipBase.OutputLogLevel
        Get
            Return _outputLogLevel
        End Get
        Set(value As SevenZipOutputLogLevel)
            _outputLogLevel = value
        End Set
    End Property

    ''' <summary>
    ''' Sets the compression level for Add and Update operations.
    ''' Valid values are 0 (no compression) to 9 (maximum compression).
    ''' </summary>
    ''' <returns>An integer representing the compression level.</returns>
    Public Property CompressionLevel As Integer Implements ISevenZipBase.CompressionLevel
        Get
            Return _compressionLevel
        End Get
        Set(value As Integer)
            If value < 0 OrElse value > 9 Then
                Throw New ArgumentOutOfRangeException(NameOf(CompressionLevel), "Compression level must be between 0 and 9.")
            End If
            _compressionLevel = value
        End Set
    End Property

    ''' <summary>
    ''' Sets the analysis level for Add and Update operations.
    ''' Valid values are 0 (no analysis) to 9 (maximum analysis).
    ''' </summary>
    ''' <returns>An integer representing the analysis level.</returns>
    Public Property AnalysisLevel As Integer Implements ISevenZipBase.AnalysisLevel
        Get
            Return _analysisLevel
        End Get
        Set(value As Integer)
            If value < 0 OrElse value > 9 Then
                Throw New ArgumentOutOfRangeException(NameOf(AnalysisLevel), "Analysis level must be between 0 and 9.")
            End If
            _analysisLevel = value
        End Set
    End Property

    Public Sub AddFile(Filepath As String) Implements ISevenZipBase.AddFile
        _filesToAdd.Add(Filepath)
    End Sub

    Public Sub AddFiles(ByVal FilepathArr As Object, Optional ByVal BaseDirectory As String = Nothing) Implements ISevenZipBase.AddFiles

        Dim paths = CoerceToStringArray(FilepathArr)
        Dim baseDir = If(String.IsNullOrWhiteSpace(BaseDirectory), Nothing, BaseDirectory)

        _filesToAdd.AddRange(
            paths.Select(
                Function(p)
                    If baseDir Is Nothing OrElse Path.IsPathRooted(p) Then
                        Return p
                    End If
                    Return Path.Combine(baseDir, p)
                End Function
            )
        )
    End Sub

    Public Sub SetMemoryBlockSize(Mode As MemoryBlockMode, Optional Size As Integer = 0) Implements ISevenZipBase.SetMemoryBlockSize
        _memoryBlock.Mode = Mode
        Select Case Mode
            Case MemoryBlockMode.Disabled, MemoryBlockMode.Enabled
                ' No size needed
            Case Else
                If Size <= 0 Then
                    Throw New ArgumentOutOfRangeException(NameOf(Size), "Size must be greater than zero when specifying a memory block size.")
                Else
                    _memoryBlock.Size = Size
                End If
        End Select
    End Sub

    Protected Shared Function CoerceToStringArray(ByVal value As Object) As String()
        If value Is Nothing Then
            Throw New ArgumentNullException(NameOf(value))
        End If

        ' If already a String()
        Dim direct = TryCast(value, String())
        If direct IsNot Nothing Then Return direct

        Dim arr = TryCast(value, System.Array)
        If arr Is Nothing Then
            Throw New ArgumentException("FilePaths must be an array (e.g., VBScript: Array(""File1"",""File2"")).", NameOf(value))
        End If

        Dim list As New List(Of String)(arr.Length)
        For Each item In arr
            If item Is Nothing Then Continue For
            Dim s As String = Convert.ToString(item, Globalization.CultureInfo.InvariantCulture)
            If Not String.IsNullOrWhiteSpace(s) Then list.Add(s)
        Next

        Return list.ToArray()
    End Function

    Protected Shared Function Write7ZipListFile(ByVal paths As IEnumerable(Of String)) As String
        Dim listPath = Path.Combine(Path.GetTempPath(), $"7z_filelist_{Guid.NewGuid():N}.txt")

        Dim utf8NoBom As New UTF8Encoding(encoderShouldEmitUTF8Identifier:=False)

        File.WriteAllLines(listPath, paths, utf8NoBom)
        Return listPath
    End Function

    Protected Class MemoryBlock
        Property Mode As MemoryBlockMode = MemoryBlockMode.Disabled
        Property Size() As Integer = 0

        Public Function ToArg() As String
            Select Case Mode
                Case MemoryBlockMode.Disabled
                    Return Nothing
                Case MemoryBlockMode.Enabled
                    Return " -ms=on"
                Case MemoryBlockMode.Bytes
                    Return " -ms=" & Size & "b"
                Case MemoryBlockMode.KiB
                    Return " -ms=" & Size & "k"
                Case MemoryBlockMode.MiB
                    Return " -ms=" & Size & "m"
                Case MemoryBlockMode.GiB
                    Return " -ms=" & Size & "g"
                Case MemoryBlockMode.TiB
                    Return " -ms=" & Size & "t"
                Case Else
                    Throw New InvalidOperationException("Unknown MemoryBlockMode.")
            End Select
        End Function

    End Class


    Protected Function GetCommandString(Mode As UpdateMode) As String
        Dim sb As New System.Text.StringBuilder(ConsolePath)

        Select Case Mode
            Case UpdateMode.AddToArchive
                sb.Append(" a")
            Case UpdateMode.UpdateArchive
                sb.Append(" u")
            Case Else
                Throw New ArgumentOutOfRangeException(NameOf(Mode), "Invalid update mode specified.")
        End Select

        Console.WriteLine($"ArchivePath: ""{_archivePath}""")
        sb.Append($" ""{_archivePath}""")

        ' Write file list to temp file
        Dim listFilePath = Write7ZipListFile(_filesToAdd)
        sb.Append(" @""").Append(listFilePath).Append("""")

        Console.WriteLine($"OutputLogLevel: {CInt(OutputLogLevel)} - {OutputLogLevel.ToString}")
        If OutputLogLevel >= SevenZipOutputLogLevel.Files AndAlso OutputLogLevel <= SevenZipOutputLogLevel.Operations Then
            sb.Append(" -bb").Append(CInt(OutputLogLevel))
        End If

        Console.WriteLine($"CompressionLevel: {CompressionLevel}")
        sb.Append(" -m").Append("x" & CompressionLevel.ToString())

        Console.WriteLine($"MemoryBlockMode: {_memoryBlock.Mode.ToString()} Size: {_memoryBlock.Size}")
        Dim arg = _memoryBlock.ToArg()
        If arg IsNot Nothing Then sb.Append(arg)

        Return sb.ToString()
    End Function

    Protected Enum UpdateMode
        AddToArchive
        UpdateArchive
    End Enum


End Class

