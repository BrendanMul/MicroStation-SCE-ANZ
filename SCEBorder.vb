Public Class SCEBorder

    Public Shared oClientBorder As DataTable
    Public Shared BorderCount As Integer
    Public Shared BorderNames As Array

    Public Shared Sub Initialise(ByVal sBorderName As String)

    End Sub

    Public Shared Function GetIndex(ByVal sBorderName As String) As Integer
        'Return index of specified border
        Dim iReturn As Integer = -1

        Return iReturn
    End Function

    Public Shared Function GetClientBorderTable(ByVal sRegion As String, ByVal sClient As String, ByVal sBorderName As String) As DataTable
        'return client border data
        Dim oReturn As DataTable = Nothing
        Dim dsDataSet As DataSet

        dsDataSet = SCEAccess.GetAllDB(JEGCore.ClientDBFullName(sRegion, sClient), "TB" & sBorderName, "Tag")

        If SCEVB.IsNullDataSet(dsDataSet) = False Then
            oReturn = dsDataSet.Tables(0)
        End If

        Return oReturn
    End Function

    Public Shared Function MergeBorderInfo(ByVal oBorderTable As DataTable, ByVal oTagTable As DataTable) As DataTable
        'return merged tables
        Dim i As Integer
        Dim j As Integer

        If oBorderTable Is Nothing Then
            GoTo ExitFunction
        End If
        If oTagTable Is Nothing Then
            oBorderTable = Nothing
            GoTo ExitFunction
        End If

        oBorderTable.Columns.Add("Value")
        For i = 0 To oBorderTable.Rows.Count - 1
            For j = 0 To oTagTable.Rows.Count - 1
                If UCase(oBorderTable.Rows(i).Item("Tag")) = UCase(oTagTable.Rows(j).Item("Tag")) Then
                    oBorderTable.Rows(i).Item("Value") = oTagTable.Rows(j).Item("Value")
                End If
            Next
        Next
ExitFunction:
        Return oBorderTable
    End Function

End Class
