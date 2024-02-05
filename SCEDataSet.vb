Public Class SCEDataSet
    'Functions to automate interrogation of datasets

    Public Shared Function CreateDataset(ByVal sName As String, Optional ByVal sColumns As String = "", Optional ByVal bGetStandardColumns As Boolean = False, Optional ByVal sTableName As String = "", Optional ByVal sDBFullName As String = "", Optional ByVal sPassword As String = "") As DataSet
        'Return new dataset
        Dim dsDataSet As New DataSet
        Dim aSplit As Array = Nothing
        Dim i As Integer

        If sDBFullName = "" Then
            sDBFullName = JEGCore.SCECoreDB
        End If

        If sPassword = "" Then
            sPassword = JEGCore.Password
        End If

        If bGetStandardColumns = True Then
            sColumns = SCEAccess.GetFieldDB(sDBFullName, "LUTables", "TableName", sTableName, "Columns", sPassword)
        End If
        'MsgBox(SCECore.SCECoreDB)
        'MsgBox(sTableName)

        aSplit = Split(sColumns, "|")
        dsDataSet.Tables.Add()
        dsDataSet.Tables(0).TableName = sName
        For i = 0 To UBound(aSplit)
            dsDataSet.Tables(0).Columns.Add(aSplit(i))
        Next

        Return dsDataSet
    End Function

    Public Shared Function GetColumn(ByVal oTable As DataTable, ByVal sColumnName As String, Optional ByVal bDistinct As Boolean = False) As Array
        'return column contents
        Dim aReturn() As String = Nothing
        Dim i As Integer
        Dim j As Integer
        Dim iIndex As Integer = -1

        iIndex = GetColumnIndex(oTable, sColumnName)

        If iIndex > -1 Then
            For i = 0 To oTable.Rows.Count - 1
                If Not oTable.Rows(i).Item(iIndex) Is System.DBNull.Value Then
                    ReDim Preserve aReturn(i)
                    aReturn(i) = oTable.Rows(i).Item(iIndex)
                End If
            Next
        End If

        If bDistinct = True Then
            Dim aReturn1() As String = Nothing
            For i = 0 To UBound(aReturn)
                If i = 0 Then
                    ReDim aReturn1(0)
                    aReturn1(0) = aReturn(0)
                Else
                    For j = 0 To UBound(aReturn1)
                        If UCase(aReturn(i)) = UCase(aReturn1(j)) Then
                            GoTo NextRegion
                        End If
                    Next
                    ReDim Preserve aReturn1(UBound(aReturn1) + 1)
                    aReturn1(UBound(aReturn1)) = aReturn(i)
                End If
NextRegion:
            Next
            aReturn = aReturn1
        End If

        Return aReturn
    End Function

    Public Shared Function GetColumnIndex(ByVal oTable As DataTable, ByVal sColumnName As String) As Integer
        'return index of column
        Dim iReturn As Integer = -1
        Dim i As Integer

        If oTable Is Nothing Then
            GoTo ExitFunction
        End If

        For i = 0 To oTable.Columns.Count - 1
            If UCase(sColumnName) = UCase(oTable.Columns.Item(i).ColumnName) Then
                iReturn = i
                Exit For
            End If
        Next
ExitFunction:
        Return iReturn
    End Function

    'Public Shared Function GetDistinct(ByVal oTable As DataTable, ByVal sColumn As String) As DataTable
    '    Dim dv As DataView

    '    dv = New DataView(oTable)
    '    oTable = dv.ToTable(True, "CellName", "Location")
    '    Return oTable
    'End Function


    Public Shared Function GetField(ByVal oTable As DataTable, ByVal sFilterColumn As String, ByVal sFilterValue As String, ByVal sColumn As String, Optional ByVal bLike As Boolean = False) As String
        'return field from table
        Dim sReturn As String = ""
        Dim oRow As DataRow = Nothing

        oRow = GetRow(oTable, sFilterColumn, sFilterValue, bLike)

        If Not oRow Is Nothing Then
            If Not oRow.Item(sColumn) Is System.DBNull.Value Then
                sReturn = oRow.Item(sColumn)
            End If
        End If

        Return sReturn
    End Function

    Public Shared Function GetRow(ByVal oTable As DataTable, ByVal sColumn As String, ByVal sValue As String, Optional ByVal bLike As Boolean = False) As DataRow
        'Dim table As DataTable = dsDataSet.Tables(sTableName)
        Dim expression As String
        Dim oRows() As DataRow
        Dim oReturn As DataRow = Nothing

        If bLike = True Then
            expression = sColumn & " LIKE " & "'%" & Replace(sValue, "'", "''") & "%'"
        Else
            expression = sColumn & "='" & Replace(sValue, "'", "''") & "'"
        End If

        'expression = "Date > #1/1/00#"

        ' Use the Select method to find all rows matching the filter.
        oRows = oTable.Select(expression)
        'If oRows.GetUpperBound(0) > 0 Then
        If Not oRows Is Nothing Then
            Try
                oReturn = oRows(0)
            Catch
            End Try
        End If

        Return oReturn
    End Function

    Public Shared Function GetRows(ByVal dsDataSet As DataSet, ByVal sTableName As String, ByVal sColumn As String, ByVal sValue As String, Optional ByVal bLike As Boolean = False)
        Dim table As DataTable = dsDataSet.Tables(sTableName)
        Dim expression As String
        Dim oReturn() As DataRow

        If bLike = True Then
            expression = sColumn & " LIKE '%" & sValue & "%'"
        Else
            expression = sColumn & "='" & sValue & "'"
        End If

        'expression = "Date > #1/1/00#"

        ' Use the Select method to find all rows matching the filter.
        oReturn = table.Select(expression)

        Return oReturn
    End Function

    Public Shared Function DeleteRows(ByVal dsDataSet As DataSet, ByVal sTableName As String, ByVal sColumn As String, ByVal sValue As String) As Integer
        Dim table As DataTable = dsDataSet.Tables(sTableName)
        Dim expression As String
        Dim oRows() As DataRow
        Dim iReturn As Integer = 0

        'expression = sColumn & " LIKE '%" & sValue & "%'"
        expression = sColumn & "='" & sValue & "'"
        'expression = "Date > #1/1/00#"

        ' Use the Select method to find all rows matching the filter.
        oRows = table.Select(expression)
        For Each oDelRow As DataRow In oRows
            iReturn = iReturn + 1
            oDelRow.Delete()
            oDelRow.AcceptChanges()
        Next

        Return iReturn
    End Function

    Public Shared Function Merge(ByVal dsDataSet1 As DataSet, ByVal dsDataSet2 As DataSet) As DataSet
        'merge two datasets
        Dim dsReturn As DataSet = Nothing

        If SCEVB.IsNullDataSet(dsDataSet1) = True Then
            If SCEVB.IsNullDataSet(dsDataSet2) = False Then
                dsReturn = dsDataSet2
            End If
        Else
            If SCEVB.IsNullDataSet(dsDataSet2) = True Then
                dsReturn = dsDataSet1
            Else
                'Merge datasets
                dsReturn = dsDataSet1
                dsReturn.Merge(dsDataSet2)
            End If
        End If

        Return dsReturn
    End Function

    Public Shared Function CreateColumns(ByVal oTable As DataTable, ByVal sColumns As String, Optional ByVal bAppend As Boolean = False) As DataTable
        'create columns in specified table
        Dim i As Integer
        Dim aSplit As Array
        Dim oColumn As DataColumn

        If bAppend = False Then
            oTable = New DataTable
        End If

        aSplit = Split(sColumns, "|")

        If SCEVB.IsNullArray(aSplit) = True Then
            GoTo ExitFunction
        End If

        'Add columns to table
        For i = 0 To UBound(aSplit)
            oColumn = New DataColumn
            oColumn.ColumnName = aSplit(i)
            oTable.Columns.Add(oColumn)
        Next
ExitFunction:
        Return oTable
    End Function

    Public Shared Function GetColumnsAsString(ByVal oTable As DataTable) As String
        'return "|" delimited column string
        Dim i As Integer
        Dim sReturn As String = ""

        For i = 0 To oTable.Columns.Count - 1
            If sReturn = "" Then
                sReturn = oTable.Columns(0).ColumnName
            Else
                sReturn = sReturn & "|" & oTable.Columns(i).ColumnName
            End If
        Next

        Return sReturn
    End Function

    Public Shared Function GetRowAsString(ByVal oRow As DataRow) As String
        'return "|" delimited row item string
        Dim i As Integer
        Dim sReturn As String = ""

        For i = 0 To UBound(oRow.ItemArray)
            If sReturn = "" Then
                sReturn = oRow.Item(i)
            Else
                sReturn = sReturn & "|" & oRow.Item(i)
            End If
        Next

        Return sReturn
    End Function

    Public Shared Function FileToDataSet(ByVal oReturn As DataTable, ByVal sFileName As String, Optional ByVal sSeparator As String = "|") As DataTable
        'Return datatable of file contents
        'Dim oReturn As New DataTable
        Dim sLine As String = ""
        Dim aSplit As Array = Nothing
        Dim i As Integer
        Dim oRow As DataRow = Nothing

        If My.Computer.FileSystem.FileExists(sFileName) = False Then
            GoTo ExitFunction
        End If

        FileOpen(1, sFileName, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
        Do While Not EOF(1)
            sLine = Trim(LineInput(1))
            aSplit = Split(sLine, sSeparator)
            'Create columns
            If oReturn.Columns.Count = 0 Then
                For i = 0 To UBound(aSplit)
                    oReturn.Columns.Add(aSplit(i))
                Next
                GoTo NextLine
            End If

            'Add Rows
            oRow = oReturn.Rows.Add
            For i = 0 To UBound(aSplit)
                If UBound(oRow.ItemArray) < i Then
                    Exit For
                End If
                oRow.Item(i) = aSplit(i)
            Next
            oRow.AcceptChanges()
NextLine:
        Loop
        FileClose(1)

ExitFunction:
        Return oReturn
    End Function
End Class
