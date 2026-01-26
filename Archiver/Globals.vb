Friend Module Globals

    Private ReadOnly Property ConsoleName As String = "7z.exe"
    Property Directory As String = "C:\Program Files\7-Zip"

    ''' <summary>
    ''' Full path to the 7-Zip console executable.
    ''' </summary>
    ''' <returns>Full path to 7z.exe</returns>
    ReadOnly Property ConsolePath As String
        Get
            Return System.IO.Path.Combine(Directory, ConsoleName)
        End Get
    End Property


End Module
