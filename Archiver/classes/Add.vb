'<ComVisible(True), InterfaceType(ComInterfaceType.InterfaceIsDual)>
'<Guid("052830F7-6AEB-495E-A209-3AB33AEEAFA1")>
'Public Interface IAdd
'    Inherits ISevenZipBase
'
'End Interface

<ProgId("Archiver.Add"), ComVisible(True), ClassInterface(ClassInterfaceType.None)>
<Guid("1D8FF0DD-46A3-425F-B903-5479B417CD00")>
Public Class Add

    Inherits SevenZipBase
    Implements ISevenZipBase

    Protected Overrides ReadOnly Property OperationMode As UpdateMode
        Get
            Return UpdateMode.AddToArchive
        End Get
    End Property

    Friend Sub New(ByVal ArchivePath As String)
        MyBase.New(ArchivePath)
    End Sub

End Class