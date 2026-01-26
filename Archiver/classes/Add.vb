<ComVisible(True), InterfaceType(ComInterfaceType.InterfaceIsDual)>
<Guid("052830F7-6AEB-495E-A209-3AB33AEEAFA1")>
Public Interface IAdd
    Inherits ISevenZipBase
End Interface

<ComVisible(True), ClassInterface(ClassInterfaceType.None)>
<Guid("1D8FF0DD-46A3-425F-B903-5479B417CD00")>
Public Class Add
    Inherits SevenZipBase
    Friend Sub New(ByVal ArchivePath As String)
        MyBase.New(ArchivePath)
    End Sub

End Class