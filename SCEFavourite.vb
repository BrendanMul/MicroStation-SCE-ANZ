Public Class SCEFavourite

    Public Shared dsFavouriteClients As DataSet
    Public Shared dsFavouriteCells As DataSet
    Public Shared ClientsLoaded As Boolean
    Public Shared ClientCount As Integer
    Public Shared CellsLoaded As Boolean
    Public Shared CellCount As Integer

    Public Shared Sub LoadClients()
        'load favourite clients into memory

        If JEGCore.bDebug = True Then JEGCore.RecordSequence(System.Reflection.MethodBase.GetCurrentMethod.Module.Name, System.Reflection.MethodBase.GetCurrentMethod.Name())

        If ClientsLoaded = True Then
            GoTo ExitSub
        End If

        dsFavouriteClients = SCEAccess.GetAllDB(JEGCore.UserDatabase, "LUFavouriteClients", "Client")
        ClientsLoaded = True

        If SCEVB.IsNullDataSet(dsFavouriteClients) = True Then
            ClientCount = 0
        Else
            ClientCount = dsFavouriteClients.Tables(0).Rows.Count
        End If
ExitSub:
    End Sub

    Public Shared Sub LoadCells()
        'load favourite cells into memory
        If CellsLoaded = True Then
            GoTo ExitSub
        End If
        dsFavouriteCells = SCEAccess.GetAllDB(JEGCore.UserDatabase, "LUFavouriteCells", "CellName")
        CellsLoaded = True

        If SCEVB.IsNullDataSet(dsFavouriteCells) = True Then
            CellCount = 0
        Else
            CellCount = dsFavouriteCells.Tables(0).Rows.Count
        End If
ExitSub:
    End Sub

    Public Shared Function GetIndex(ByVal sCustomName As String) As Integer
        'return index of specified favourite
        Dim i As Integer
        Dim iReturn As Integer = -1

        For i = 0 To dsFavouriteClients.Tables(0).Rows.Count - 1
            If UCase(sCustomName) = UCase(dsFavouriteClients.Tables(0).Rows(i).Item("CustomName")) Then
                iReturn = i
                Exit For
            End If
        Next

        Return iReturn
    End Function
End Class
