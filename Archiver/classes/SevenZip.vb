<ComVisible(True), InterfaceType(ComInterfaceType.InterfaceIsDual)>
<Guid("7B4239A2-56D4-484A-99D0-C748DA5293F6")>
Public Interface ISevenZip
    Property ProgDirectory As String
    Function Update() As IUpdate
    Function Add() As IAdd
End Interface

<ProgId("Archiver.7z"), ComVisible(True), ClassInterface(ClassInterfaceType.None)>
<Guid("7B4239A2-56D4-484A-99D0-C748DA5293F6")>
Public Class SevenZip

    Implements ISevenZip

    Private _dir As String = "C:\Program Files\7-Zip"
    Private Const _consoleName As String = "7z.exe"

    Private ReadOnly Property ConsolePath As String
        Get
            Return System.IO.Path.Combine(_dir, _consoleName)
        End Get
    End Property

    Public Property ProgDirectory As String Implements ISevenZip.ProgDirectory
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As String)
            Throw New NotImplementedException()
        End Set
    End Property

    Public Function Update() As IUpdate Implements ISevenZip.Update
        Return New Update
    End Function

    Public Function Add() As IAdd Implements ISevenZip.Add
        Return New Add
    End Function

End Class

Public MustInherit Class SevenZipBase
    Protected Function GetConsolePath() As String
        Dim archiver As New SevenZip
        Return archiver.ConsolePath
    End Function

End Class