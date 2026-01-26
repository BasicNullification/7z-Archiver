Public Interface IAdd
    Inherits ISevenZipBase
End Interface
Public Class Add
    Inherits SevenZipBase
    Friend Sub New(ByVal ArchivePath As String)
        MyBase.New(ArchivePath)
    End Sub

End Class