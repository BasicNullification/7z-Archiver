Public Interface IUpdate
    Inherits ISevenZipBase
End Interface

Public Class Update

    Inherits SevenZipBase

    Friend Sub New(ByVal ArchivePath As String)
        MyBase.New(ArchivePath)
    End Sub
End Class