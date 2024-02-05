Imports JMECore
Imports System.Collections.Generic
imports System.Drawing

Public Class SCEForms

    Public Shared Function GetDGVColumnIndex(ByVal oDGV As DataGridView, ByVal sColumnName As String)
        'Get value from specified row
        Dim i As Integer
        Dim iReturn As Integer = -1

        For i = 0 To oDGV.ColumnCount - 1
            If oDGV.Columns(i).HeaderText = sColumnName Then
                iReturn = i
                Exit For
            End If
        Next

        Return iReturn
    End Function

    Public Shared Function GetFieldFromRow(ByVal oDGV As DataGridView, ByVal iRowIndex As Integer, ByVal sColumnName As String)
        'Get value from specified row
        Dim iIndex As Integer = GetDGVColumnIndex(oDGV, sColumnName)
        Dim sReturn As String = ""

        If iIndex > -1 Then
            If Not oDGV.Rows(iRowIndex).Cells(iIndex).Value Is System.DBNull.Value Then
                sReturn = oDGV.Rows(iRowIndex).Cells(iIndex).Value
            End If
        End If

        Return sReturn
    End Function

    Public Shared Function SortTable(ByVal oTable As DataTable, ByVal sColumn As String) As DataTable
        'sort datatable
        Dim view As DataView = oTable.DefaultView

        ' By default, the first column sorted ascending.
        view.Sort = sColumn

        Return oTable
    End Function

    Public Shared Sub PopulateCBO(ByVal oCombo As ComboBox, ByVal aOptions As Array)
        Dim i As Integer
        Dim j As Integer

        oCombo.Items.Clear()
        If aOptions(0) <> "" Then
            'Add all data to combo box
            For i = 0 To UBound(aOptions)
                If i > 0 Then
                    For j = 0 To i
                        If i <> j Then
                            If UCase(aOptions(i)) = UCase(aOptions(j)) Then
                                GoTo NextItem
                            End If
                        End If
                    Next
                End If
                oCombo.Items.Add(aOptions(i))
NextItem:
            Next
        End If
    End Sub

    Public Shared Sub PopulateListBox(ByVal oListBox As ListBox, ByVal aData As Array)
        'populate listbox
        Dim i As Integer

        oListBox.Items.Clear()
        For i = 0 To UBound(aData)
            oListBox.Items.Add(aData(i))
        Next
    End Sub

    Public Shared Sub CBOItem(ByRef oCombo As System.Windows.Forms.ComboBox, ByVal sItem As String)
        Dim iIndex As Integer

        iIndex = oCombo.Items.IndexOf(sItem)
        If iIndex > -1 Then
            oCombo.SelectedIndex = iIndex
        End If
    End Sub

    Public Shared Function ConvertToLogicalDate(ByVal odate As Date) As String
        'convert date to logical date e.g. Today  07:50 AM, 18.05.2012  07.50 AM
        Dim sReturn As String = ""

        If Format(odate, "dd.MM.yyyy") = Format(Now, "dd.MM.yyyy") Then
            sReturn = "Today" & " at " & Format(odate, "hh:mm tt")
        ElseIf Format(odate, "dd.MM.yyyy") = Format(Now.AddDays(-1), "dd.MM.yyyy") Then
            sReturn = "Yesterday" & " at " & Format(odate, "hh:mm tt")
        Else
            sReturn = Format(odate, "dd MMM yyyy  hh:mm tt")
        End If

        Return sReturn
    End Function

    '    Public Shared Sub PopulateDetailColumns(ByVal sDBFullName As String, ByVal oDGV As DataGridView, ByVal oColumns As DataGridViewColumnCollection, ByVal oTable As DataTable, Optional ByVal sPassword As String = "")
    '        'set readonly access and visible status to datagridview
    '        Dim i As Integer
    '        Dim j As Integer
    '        Dim oColumn As DataGridViewColumn = Nothing

    '        If sPassword = "" Then
    '            sPassword = SCECore.Password
    '        End If

    '        Dim dsDataSet As DataSet = SCEAccess.GetAllDB(sDBFullName, "LUColumnDefinition", "Field", sPassword)

    '        If SCEVB.IsNullDataSet(dsDataSet) = True Then
    '            GoTo ExitSub
    '        End If

    '        oDGV.Columns.Clear()
    '        oDGV.AlternatingRowsDefaultCellStyle = Nothing

    '        For i = 0 To oColumns.Count - 1
    '            For j = 0 To dsDataSet.Tables(0).Rows.Count - 1
    '                If UCase(dsDataSet.Tables(0).Rows(j).Item("Field")) = UCase(oColumns.Item(i).Name) Then
    '                    Select Case UCase(dsDataSet.Tables(0).Rows(j).Item("ContentType"))
    '                        Case "TEXT"
    '                            oColumn = New DataGridViewColumn
    '                        Case "DATE"
    '                            oColumn = New DataGridViewColumn
    '                        Case "PERSONNEL"
    '                            oColumn = New DataGridViewColumn
    '                        Case "BOOLEAN"
    '                            oColumn = New DataGridViewComboBoxColumn
    '                        Case Else
    '                            oColumn = New DataGridViewColumn
    '                    End Select
    '                    oColumn.Tag = UCase(dsDataSet.Tables(0).Rows(j).Item("ContentType"))
    '                End If
    '            Next
    '            oDGV.Columns.Add(oColumns(i))
    '        Next

    '        For i = 0 To oDGV.ColumnCount - 1
    '            For j = 0 To dsDataSet.Tables(0).Rows.Count - 1
    '                If UCase(oDGV.Columns(i).Name) = UCase(dsDataSet.Tables(0).Rows(j).Item("Field")) Then
    '                    Try
    '                        oDGV.Columns(i).ReadOnly = dsDataSet.Tables(0).Rows(j).Item("IsReadOnly")
    '                        oDGV.Columns(i).Visible = dsDataSet.Tables(0).Rows(j).Item("IsVisible")
    '                        If oDGV.Columns(i).ReadOnly = True Then
    '                            oDGV.Columns(i).DefaultCellStyle.BackColor = Drawing.Color.LightGray
    '                        Else
    '                            oDGV.Columns(i).DefaultCellStyle.BackColor = Drawing.Color.White
    '                        End If
    '                        oDGV.Columns(i).Tag = UCase(dsDataSet.Tables(0).Rows(j).Item("ContentType"))
    '                    Catch
    '                    End Try
    '                End If
    '            Next
    '        Next

    '        'populate datagridview
    '        oDGV.DataSource = oTable
    'ExitSub:
    '    End Sub

    Public Shared Sub ColumnCheck(ByVal sDBFullName As String, ByVal oDGV As DataGridView, Optional ByVal sPassword As String = "", Optional ByVal bDisableAlternatingRowColour As Boolean = False, Optional ByVal sTableName As String = "LUColumnDefinition", Optional ByVal bNoStyleColourModification As Boolean = False)
        'set readonly access and visible status to datagridview
        Dim i As Integer
        Dim j As Integer

        If sPassword = "" Then
            sPassword = JEGCore.Password
        End If

        Dim dsDataSet As DataSet = SCEAccess.GetAllDB(sDBFullName, sTableName, "Field", sPassword)

        If SCEVB.IsNullDataSet(dsDataSet) = True Then
            GoTo ExitSub
        End If

        If bDisableAlternatingRowColour = True Then
            oDGV.AlternatingRowsDefaultCellStyle = Nothing
        End If

        For i = 0 To oDGV.ColumnCount - 1
            oDGV.Columns(i).Visible = False
        Next

        For i = 0 To oDGV.ColumnCount - 1
            For j = 0 To dsDataSet.Tables(0).Rows.Count - 1
                If UCase(oDGV.Columns(i).Name) = UCase(dsDataSet.Tables(0).Rows(j).Item("Field")) Then
                    If InStr(UCase(dsDataSet.Tables(0).Rows(j).Item("DataGrid")), UCase(oDGV.Name)) > 0 Then
                        Try
                            If Not dsDataSet.Tables(0).Rows(j).Item("Title") Is System.DBNull.Value Then
                                oDGV.Columns(i).HeaderText = dsDataSet.Tables(0).Rows(j).Item("Title")
                            End If
                            If Not dsDataSet.Tables(0).Rows(j).Item("IsReadOnly") Is System.DBNull.Value Then
                                oDGV.Columns(i).ReadOnly = dsDataSet.Tables(0).Rows(j).Item("IsReadOnly")
                            End If
                            If Not dsDataSet.Tables(0).Rows(j).Item("IsVisible") Is System.DBNull.Value Then
                                oDGV.Columns(i).Visible = dsDataSet.Tables(0).Rows(j).Item("IsVisible")
                            End If
                            If bNoStyleColourModification = False Then
                                If dsDataSet.Tables(0).Rows(j).Item("BackColour") Is System.DBNull.Value Then
                                    If oDGV.Columns(i).ReadOnly = True Then
                                        oDGV.Columns(i).DefaultCellStyle.BackColor = Drawing.Color.LightGray
                                    Else
                                        'oDGV.Columns(i).DefaultCellStyle.BackColor = Drawing.Color.White
                                    End If
                                Else
                                    oDGV.Columns(i).DefaultCellStyle.BackColor = Drawing.Color.FromName(dsDataSet.Tables(0).Rows(j).Item("BackColour"))
                                End If
                            End If
                            If Not dsDataSet.Tables(0).Rows(j).Item("ForeColour") Is System.DBNull.Value Then
                                oDGV.Columns(i).DefaultCellStyle.ForeColor = Drawing.Color.FromName(dsDataSet.Tables(0).Rows(j).Item("ForeColour"))
                            End If
                            If Not dsDataSet.Tables(0).Rows(j).Item("ContentType") Is System.DBNull.Value Then
                                oDGV.Columns(i).Tag = Trim(UCase(dsDataSet.Tables(0).Rows(j).Item("ContentType")))
                            End If
                            If Not dsDataSet.Tables(0).Rows(j).Item("SortIndex") Is System.DBNull.Value Then
                                oDGV.Columns(i).DisplayIndex = dsDataSet.Tables(0).Rows(j).Item("SortIndex")
                            End If
                            'If Not dsDataSet.Tables(0).Rows(j).Item("SortColumn") Is System.DBNull.Value Then
                            '    oDGV.Sort(oDGV.Columns(i), System.ComponentModel.ListSortDirection.Ascending)
                            'End If
                        Catch
                        End Try
                    End If
                End If
NextRow:
            Next
        Next
ExitSub:
    End Sub

    Public Shared Function GetColors() As List(Of String)
        'create a generic list of strings
        Dim colors As New List(Of String)()
        'get the color names from the Known color enum
        Dim colorNames As String() = [Enum].GetNames(GetType(KnownColor))
        'iterate thru each string in the colorNames array
        For Each colorName As String In colorNames
            'cast the colorName into a KnownColor
            Dim knownColor As KnownColor = DirectCast([Enum].Parse(GetType(KnownColor), colorName), KnownColor)
            'check if the knownColor variable is a System color
            If knownColor > KnownColor.Transparent Then
                'add it to our list
                colors.Add(colorName)
            End If
        Next
        'return the color list
        Return colors
    End Function

    Public Shared Function PopulateDataGridView(ByVal sDBFullName As String, ByVal sTable As String, ByVal dsDataSet As DataSet, ByVal dgvDataGrid As DataGridView, Optional ByVal sPassword As String = "")
        'populate datagridview with content and configure display
        Dim oColumn As DataGridViewColumn
        Dim odsColumn As DataColumn
        Dim odsRow As DataRow
        Dim oNewColumn As DataColumn
        Dim bFound As Boolean = False

        If sPassword = "" Then
            sPassword = JEGCore.Password
        End If

        Dim dsColumnDefinition As DataSet = SCEAccess.GetRowsDB(sDBFullName, sTable, "Field", "DataGrid", dgvDataGrid.Name, False, True, False, sPassword)

        'check if any columns need to be added to dataset 
        For Each odsRow In dsColumnDefinition.Tables(0).Rows
            bFound = False
            For Each odsColumn In dsDataSet.Tables(0).Columns
                If UCase(odsColumn.ColumnName) = UCase(odsRow.Item("Field")) Then
                    bFound = True
                End If
            Next
            If bFound = False Then
                'Create Column
                oNewColumn = New DataColumn
                oNewColumn.ColumnName = odsRow.Item("Field")
                dsDataSet.Tables(0).Columns.Add(oNewColumn)
            End If
        Next

        dgvDataGrid.DataSource = dsDataSet.Tables(0)
        For Each oColumn In dgvDataGrid.Columns
            oColumn.Visible = False
        Next

        'Populate dataset
        For Each odsRow In dsColumnDefinition.Tables(0).Rows
            For Each oColumn In dgvDataGrid.Columns
                If UCase(oColumn.HeaderText) = UCase(odsRow.Item("Field")) Then
                    'Title
                    If Not odsRow.Item("Title") Is System.DBNull.Value Then
                        oColumn.Tag = oColumn.HeaderText
                        oColumn.HeaderText = odsRow.Item("Title")
                    End If
                    'Read only
                    If odsRow.Item("IsReadOnly") Is System.DBNull.Value Then
                        oColumn.ReadOnly = True
                    Else
                        oColumn.ReadOnly = odsRow.Item("IsReadOnly")
                    End If
                    'Visible
                    If Not odsRow.Item("IsVisible") Is System.DBNull.Value Then
                        oColumn.Visible = odsRow.Item("IsVisible")
                    End If
                    'ForeColour
                    If odsRow.Item("ForeColour") Is System.DBNull.Value Then
                        oColumn.DefaultCellStyle.ForeColor = Color.Black
                    Else
                        oColumn.DefaultCellStyle.ForeColor = System.Drawing.Color.FromName(odsRow.Item("ForeColour"))
                    End If
                    'Index
                    If Not odsRow.Item("SortIndex") Is System.DBNull.Value Then
                        oColumn.DisplayIndex = odsRow.Item("SortIndex")
                    End If
                End If
            Next
        Next

        Return dgvDataGrid
    End Function
End Class
