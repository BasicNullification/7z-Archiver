<ComVisible(True), InterfaceType(ComInterfaceType.InterfaceIsDual)>
<Guid("7B4239A2-56D4-484A-99D0-C748DA5293F6")>
Public Interface ISevenZip
    <Description("Location of the 7-Zip program directory. Defaults to 'C:\Program Files\7-Zip'.")>
    Property ProgDirectory As String

    <Description("Update older files in the archive and add files that are not already in the archive.")>
    Function Update(ByVal ArchivePath As String) As Update

    <Description("Add files to the archive.")>
    Function Add(ByVal ArchivePath As String) As Add
End Interface

<ProgId("Archiver.7z"), ComVisible(True), ClassInterface(ClassInterfaceType.None)>
<Guid("2CB7944F-88B3-43D5-B081-C3098A8ACC69")>
Public Class SevenZip

    Implements ISevenZip

    Public Sub New()

    End Sub

    ''' <summary>
    ''' Location of the 7-Zip program directory. Defaults to 'C:\Program Files\7-Zip'.
    ''' </summary>
    ''' <returns></returns>
    Public Property ProgDirectory As String Implements ISevenZip.ProgDirectory
        Get
            Return Globals.Directory
        End Get
        Set(value As String)
            Globals.Directory = value
        End Set
    End Property

    ''' <summary>
    ''' Update older files in the archive and add files that are not already in the archive.
    ''' </summary>
    ''' <param name="ArchivePath"></param>
    ''' <returns></returns>
    Public Function Update(ByVal ArchivePath As String) As Update Implements ISevenZip.Update
        Return New Update(ArchivePath)
    End Function

    ''' <summary>
    ''' Add new files to the archive.
    ''' </summary>
    ''' <param name="ArchivePath"></param>
    ''' <returns></returns>
    Public Function Add(ByVal ArchivePath As String) As Add Implements ISevenZip.Add
        Return New Add(ArchivePath)
    End Function

End Class
