Imports MicroStationSCE.SCEAccess

Public Class SCEFavouriteClient

    Public Shared Client As String
    Public Shared Region As String
    Public Shared MSProject As String
    Public Shared CustomName As String 'Description
    Public Shared Software As String
    Public Shared SoftwareVersion As String
    Public Shared SoftwareAddon As String
    Public Shared ToolSet As String

    Shared Sub LoadFavourite(ByVal sFavourite As String)
        'Load specified favourite information
        Dim dsFavourite As DataSet

        dsFavourite = GetRowDB(SCECore.UserDatabase, "LUFavourites", "MSProject", "CustomName", sFavourite, "", "", False, True)
        Client = dsFavourite.Tables(0).Rows(0).Item("Client")
        Region = dsFavourite.Tables(0).Rows(0).Item("Region")
        MSProject = dsFavourite.Tables(0).Rows(0).Item("MSProject")
        CustomName = dsFavourite.Tables(0).Rows(0).Item("CustomName")

        If Not dsFavourite.Tables(0).Rows(0).Item("SoftwareAddon") Is System.DBNull.Value Then
            SoftwareAddon = dsFavourite.Tables(0).Rows(0).Item("SoftwareAddon")
        End If

        If Not dsFavourite.Tables(0).Rows(0).Item("ToolSet") Is System.DBNull.Value Then
            ToolSet = dsFavourite.Tables(0).Rows(0).Item("ToolSet")
        End If
    End Sub

End Class
