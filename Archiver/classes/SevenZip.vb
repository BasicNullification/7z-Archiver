<ComVisible(True), InterfaceType(ComInterfaceType.InterfaceIsDual)>
<Guid("7B4239A2-56D4-484A-99D0-C748DA5293F6")>
Public Interface ISevenZip
    <Description("Location of the 7-Zip program directory. Defaults to 'C:\Program Files\7-Zip'.")>
    Property ProgDirectory As String

    <Description("Update older files in the archive and add files that are not already in the archive.")>
    Function Update(ByVal ArchivePath As String) As IUpdate

    <Description("Add files to the archive.")>
    Function Add(ByVal ArchivePath As String) As IAdd
End Interface

<ProgId("Archiver.7z"), ComVisible(True), ClassInterface(ClassInterfaceType.None)>
<Guid("7B4239A2-56D4-484A-99D0-C748DA5293F6")>
Public Class SevenZip

    Implements ISevenZip

    Public Property ProgDirectory As String Implements ISevenZip.ProgDirectory
        Get
            Return Globals.Directory
        End Get
        Set(value As String)
            Globals.Directory = value
        End Set
    End Property

    Public Function Update(ByVal ArchivePath As String) As IUpdate Implements ISevenZip.Update
        Return New Update(ArchivePath)
    End Function

    Public Function Add(ByVal ArchivePath As String) As IAdd Implements ISevenZip.Add
        Return New Add(ArchivePath)
    End Function

End Class
