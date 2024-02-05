Public Class SCEVB

    Shared Function IsNullArray(ByVal aArray As Array) As Boolean
        'Return boolean if array is empty
        Dim bReturn As Boolean = False

        If aArray Is Nothing Then
            bReturn = True
            GoTo Exitfunction
        End If

        If aArray(0) = "" And UBound(aArray) = 0 Then
            bReturn = True
        End If
ExitFunction:
        Return bReturn
    End Function

    Shared Function IsNullDataSet(ByVal dsDataSet As DataSet) As Boolean
        'Return boolean if dataset is empty
        Dim bReturn As Boolean = False

        If dsDataSet Is Nothing Then
            bReturn = True
        Else
            If dsDataSet.Tables.Count = 0 Then
                Return True
            ElseIf dsDataSet.Tables(0).Rows.Count = 0 Then
                Return True
            End If
        End If

        Return bReturn
    End Function

    Shared Function CommentChars() As String
        'Return comment characters
        Return "';#"
    End Function

    Shared Function IsCommented(ByVal sString As String) As Boolean
        'Return if specified line is commented
        Dim bReturn As Boolean = False

        If InStr(CommentChars, Microsoft.VisualBasic.Left(Trim(sString), 1)) > 0 Then
            bReturn = True
        End If

        Return bReturn
    End Function

    Shared Function DataTableToArray(ByVal oTable As DataTable, ByVal sColumnName As String, Optional ByVal bRemoveEmpties As Boolean = False) As Array
        'Return array of values from specified database column
        Dim sArray() As String = Nothing
        Dim i As Integer
        Dim j As Integer = -1

        If Not oTable Is Nothing Then
            For i = 0 To oTable.Rows.Count - 1
                If bRemoveEmpties = True Then
                    If Trim(oTable.Rows(i).Item(sColumnName)) = "" Then GoTo NextItem
                End If
                j = j + 1
                ReDim Preserve sArray(j)
                sArray(j) = oTable.Rows(i).Item(sColumnName)
NextItem:
            Next
        End If

        If j = -1 Then
            ReDim sArray(0)
        End If

        Return sArray
    End Function

    Shared Function DataRowToString(ByVal oRow As DataRow) As String
        'Return string of values from specified database row
        Dim sReturn As String = ""
        Dim i As Integer
        Dim j As Integer = -1

        If Not oRow Is Nothing Then
            For i = 0 To UBound(oRow.ItemArray)
                If oRow(i) Is System.DBNull.Value Then
                    If sReturn = "" Then
                        sReturn = " "
                    Else
                        sReturn = sReturn & "|"
                    End If
                    GoTo NextItem
                End If
                If Trim(oRow.Item(i)) = "" Then
                    sReturn = sReturn & "|"
                    GoTo NextItem
                End If
                If sReturn = "" Then
                    sReturn = oRow(i)
                Else
                    sReturn = sReturn & "|" & oRow(i)
                End If
NextItem:
            Next
        End If

        Return sReturn
    End Function

    Public Shared Function ArrayToString(ByVal aArray As Array, ByVal sSeparator As String) As String
        'convert specified array to string 
        Dim sReturn As String = ""
        Dim sString As String

        For Each sString In aArray
            If sReturn = "" Then
                sReturn = sString
            Else
                sReturn = sReturn & sSeparator & sString
            End If
        Next

        Return sReturn
    End Function

    Public Shared Function AppendArray(ByVal aArray1 As Object, ByVal aArray2 As Object) As Array
        'Append one variant onto end of another (both variants are specified)
        Dim i As Integer
        Dim j As Integer
        Dim aReturn() As String

        If IsNullArray(aArray1) = True Then
            aReturn = aArray2
        ElseIf IsNullArray(aArray2) = True Then
            aReturn = aArray1
        Else
            ReDim Preserve aReturn(UBound(aArray1))
            For j = 0 To UBound(aArray1)
                aReturn(j) = aArray1(j)
            Next
            j = UBound(aArray1) + UBound(aArray2) + 1

            ReDim Preserve aReturn(j)
            j = UBound(aArray1)
            For i = 0 To UBound(aArray2)
                j = j + 1
                aReturn(j) = aArray2(i)
            Next
        End If

        Return aReturn
    End Function

    Public Shared Function TrimArray(ByVal aString() As String) As Array
        'Return array without spaces
        Dim i As Integer
        Dim j As Integer = -1
        Dim aResult() As String = Nothing

        For i = 0 To UBound(aString)
            If aString(i) <> "" Then
                j = j + 1
                ReDim Preserve aResult(j)
                aResult(j) = aString(i)
            End If
        Next

        Return aResult
    End Function

    Function DataSetToVar(ByVal dsDataSet As DataSet, ByVal sColumn As String) As Array
        'convert dataset column data to array
        Dim i As Integer
        Dim aReturn() As String = Nothing

        If SCEVB.IsNullDataSet(dsDataSet) = True Then
            ReDim aReturn(0)
            GoTo ExitFunction
        End If

        For i = 0 To dsDataSet.Tables(0).Rows.Count - 1
            If Not dsDataSet.Tables(0).Rows(i).Item(sColumn) Is System.DBNull.Value Then
                ReDim Preserve aReturn(i)
                aReturn(i) = Trim(dsDataSet.Tables(0).Rows(i).Item(sColumn))
            End If
        Next
ExitFunction:
        Return aReturn
    End Function

    Public Shared Function UniqueTimeStamp() As String
        'return a unique string based on curren datetime stamp
        Return Format(Now, "yyyy.MM.dd hh:mm tt")
    End Function

    Public Shared Function UniqueDateStamp() As String
        'return a unique string based on curren date stamp
        Return Format(Now, "yyyyMMdd")
    End Function

End Class
