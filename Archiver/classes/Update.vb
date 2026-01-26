<ComVisible(True), InterfaceType(ComInterfaceType.InterfaceIsDual)>
<Guid("F78D7E0B-06B8-4AE9-A041-EF1430546489")>
Public Interface IUpdate
    Inherits ISevenZipBase
End Interface

<ComVisible(True), ClassInterface(ClassInterfaceType.None)>
<Guid("E7C3102B-95EC-4F3B-AC63-A23C7AE6FF6F")>
Public Class Update

    Inherits SevenZipBase

    Friend Sub New(ByVal ArchivePath As String)
        MyBase.New(ArchivePath)
    End Sub
End Class