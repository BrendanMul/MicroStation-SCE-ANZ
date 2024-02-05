Public Class SCEExcel

    Function GetCell(ByVal sFullFileName As String, ByVal sSearchColumn As String, ByVal sFilterString As String, ByVal sReturnColumn As String) As String
        'Return value from specified column
        Dim oExcel As Object 'Microsoft.Office.Interop.Excel.Application
        Dim i As Integer
        Dim sReturn As String = ""
        Dim dsCellLib As Object

        oExcel = CreateObject("Excel.Application")
        oExcel.Workbooks.Open(sFullFileName)

        If My.Computer.FileSystem.FileExists(sFullFileName) = True Then
            ''Dim dsCellLib As Microsoft.Office.Interop.Excel.Range
            dsCellLib = oExcel.Columns.EntireColumn(sSearchColumn)

            For i = 1 To dsCellLib.Cells.Count
                If dsCellLib.Range(sSearchColumn & i).Value = "" Then Exit For 'stop searching when empty field found
                If dsCellLib.Range(sSearchColumn & i).Value = sFilterString Then
                    sReturn = dsCellLib.Range(sReturnColumn & i).Value
                    Exit For
                End If
            Next
        End If

        Return sReturn
    End Function

End Class
