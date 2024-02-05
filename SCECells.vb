
Public Class SCECells

    Public Shared dsCells As DataSet

    Public Shared Sub Initialise(ByVal sRegion As String, ByVal sClient As String)
        'load specified client cells into memory if not already
        Dim dsDataSet As DataSet

        If SCEVB.IsNullDataSet(dsCells) = True Then
            GoTo Load
        End If

        'Check if client cells are already loaded
        dsDataSet = SCEDataSet.GetRows(dsCells, "RGCellCatalogue", "Client", sClient)
        If SCEVB.IsNullDataSet(dsDataSet) = False Then
            GoTo ExitSub
        End If
Load:
        'Load cell catalogue
        dsDataSet = SCEAccess.GetAllDB(JEGCore.ClientPath & "\" & sRegion & "\" & sClient & "\Database\SCE Core - " & sClient & ".accdb", "RGCellCatalogue", "CellName")
        If SCEVB.IsNullDataSet(dsDataSet) = False Then
            GoTo ExitSub
        End If

        'Merge client cells into memory
        dsCells = SCEDataSet.Merge(dsCells, dsDataSet)
ExitSub:
    End Sub

End Class
