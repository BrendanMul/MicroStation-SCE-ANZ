Imports JMECore.SCEFile

Public Class SCEAccess
    Public Shared Function GetColumnDB(ByVal sDBFullName As String, ByVal sTable As String, ByVal sColumn As String, ByVal sFilterColumns As String, ByVal sFilterStrings As String, _
                Optional ByVal bSort As Boolean = False, Optional ByVal bWholeWord As Boolean = False, _
                Optional ByVal bDistinct As Boolean = False, Optional ByVal sPassword As String = "") As DataSet
        'Return specified column from table with query
        Dim cnSQL As OleDbConnection
        Dim sSQL As String = ""
        Dim dsReturn As New DataSet
        Dim da As OleDb.OleDbDataAdapter
        Dim sDistinct As String
        Dim sComparison As String
        Dim sEndComparison As String
        Dim aColumns As Array
        Dim aValues As Array
        Dim sSort As String = ""
        Dim i As Integer

        If sPassword = "" Then
            sPassword = JEGCore.Password
        End If

        Try
            aColumns = Split(sFilterColumns, "|")
            aValues = Split(sFilterStrings, "|")

            'Set comparison variable
            If bSort = True Then
                sSort = " ORDER BY " & sColumn
            End If

            If bWholeWord = True Then
                sComparison = " ='"
                sEndComparison = ""
            Else
                sComparison = " LIKE '%"
                sEndComparison = "%"
            End If

            If bDistinct = True Then
                sDistinct = "DISTINCT " 'DISTINCT only returns different values (ignores multiple occurances of the same value)
            Else
                sDistinct = ""
            End If

            'Execute the select statement
            If aColumns(0) = "" Then
                sSQL = "SELECT " & sDistinct & "* FROM [" & sTable & "]"
            Else
                For i = 0 To UBound(aColumns)
                    If i = 0 Then
                        sSQL = "SELECT " & sDistinct & sColumn & " FROM [" & sTable & "] WHERE " & aColumns(i) & sComparison & aValues(i) & sEndComparison & "'"
                    Else
                        sSQL = sSQL & " AND " & aColumns(i) & sComparison & aValues(i) & sEndComparison & "'"
                    End If
                Next
            End If

            sSQL = sSQL & sSort

            'Try
            cnSQL = New OleDbConnection("Provider=" & My.Settings.DBProvider & ";Data Source=" & sDBFullName & ";Jet OLEDB:Database Password=" & sPassword & ";Persist Security Info=False")
            cnSQL.Open()
            'Catch
            '    cnSQL = New OleDbConnection("Provider=" & My.Settings.DBProvider & ";Data Source=" & sDBFullName & ";Jet OLEDB:Database Password=norwich;Persist Security Info=False")
            '    cnSQL.Open()
            'End Try
            da = New OleDb.OleDbDataAdapter(sSQL, cnSQL)
            da.Fill(dsReturn, sTable)

            Return dsReturn

            cnSQL.Close()
            da = Nothing
            dsReturn = Nothing
            cnSQL.Dispose()
            cnSQL = Nothing
            Exit Function
        Catch e As OleDbException
            MsgBox(e.Message, MsgBoxStyle.Critical, "oledb Error")
        Catch e As Exception
            MsgBox(e.Message, MsgBoxStyle.Critical, "General Error")
        End Try
        Return Nothing
    End Function

    Shared Function GetRowDB(ByVal sDBFullName As String, ByVal sTable As String, ByVal sColumn As String, ByVal sFilterColumn As String, _
        ByVal sFilterString As String, Optional ByVal sFilterColumn1 As String = "", Optional ByVal sFilterString1 As String = "", _
        Optional ByVal bSort As Boolean = False, Optional ByVal bWholeWord As Boolean = False, Optional ByVal bDistinct As Boolean = False, _
        Optional ByVal sPassword As String = "") As DataRow
        'Return specified column from table with query
        Dim cnSQL As OleDbConnection
        Dim sSQL As String
        Dim dsDataSet As New DataSet
        Dim oReturn As DataRow = Nothing
        Dim da As OleDb.OleDbDataAdapter
        Dim sComparison As String = ""
        Dim sEndComparison As String = ""

        If sPassword = "" Then
            sPassword = JEGCore.Password
        End If

        Try
            'Set comparison variable
            If bWholeWord = True Then
                sComparison = "='"
                sEndComparison = ""
            Else
                sComparison = " LIKE '%"
                sEndComparison = "%"
            End If

            'Execute the select statement
            If bSort = True Then
                If sFilterColumn = "" Then
                    sSQL = "SELECT * FROM [" & sTable & "] ORDER BY " & sColumn
                Else
                    sSQL = "SELECT * FROM [" & sTable & "] WHERE " & sFilterColumn & sComparison & sFilterString & sEndComparison & "'"
                    If sFilterColumn1 <> "" Then
                        sSQL = sSQL & " AND " & sFilterColumn1 & sComparison & sFilterString1 & sEndComparison & "'"
                    End If
                    sSQL = sSQL & " ORDER BY " & sColumn
                End If
            Else
                If sFilterColumn = "" Then
                    sSQL = "SELECT * FROM [" & sTable & "]"
                Else
                    sSQL = "SELECT * FROM [" & sTable & "] WHERE " & sFilterColumn & sComparison & sFilterString & sEndComparison & "'"
                    If sFilterColumn1 <> "" Then
                        sSQL = sSQL & " AND " & sFilterColumn1 & sComparison & sFilterString1 & sEndComparison & "'"
                    End If
                End If
            End If

            Try
                cnSQL = New OleDbConnection("Provider=" & My.Settings.DBProvider & ";Data Source=" & sDBFullName & ";Jet OLEDB:Database Password=" & JEGCore.Password & ";Persist Security Info=False")
                cnSQL.Open()
            Catch
                cnSQL = New OleDbConnection("Provider=" & My.Settings.DBProvider & ";Data Source=" & sDBFullName & ";Jet OLEDB:Database Password=norwich;Persist Security Info=False")
                cnSQL.Open()
            End Try

            da = New OleDb.OleDbDataAdapter(sSQL, cnSQL)
            da.Fill(dsDataSet, sTable)

            If dsDataSet.Tables(0).Rows.Count > 0 Then
                oReturn = dsDataSet.Tables(0).Rows(0)
            End If

            cnSQL.Close()
            cnSQL.Dispose()
            da.Dispose()
            dsDataSet.Dispose()

            da = Nothing
            dsDataSet = Nothing
            cnSQL = Nothing
            GoTo ExitFunction
        Catch e As OleDbException
            'MsgBox(e.Message, MsgBoxStyle.Critical, "oledb Error")
        Catch e As Exception
            MsgBox(e.Message, MsgBoxStyle.Critical, "General Error")
        End Try
ExitFunction:
        Return oReturn
    End Function

    Shared Function GetRowsDB(ByVal sDBFullName As String, ByVal sTable As String, ByVal sColumn As String, ByVal sFilterColumns As String, _
            ByVal sFilterValues As String, Optional ByVal bSort As Boolean = False, Optional ByVal bWholeWord As Boolean = False, _
            Optional ByVal bDistinct As Boolean = False, Optional ByVal sPassword As String = "", Optional ByVal bOR As Boolean = False) As DataSet
        'Return multiple rows
        'Dim cnSQL As OleDbConnection
        Dim cnSQL As OdbcConnection
        Dim sSQL As String = ""
        Dim dsReturn As New DataSet
        'Dim da As OleDb.OleDbDataAdapter
        Dim sComparison As String = ""
        Dim sEndComparison As String = ""
        Dim aColumns As Array
        Dim aValues As Array
        Dim i As Integer
        Dim sSort As String = ""
        Dim sOR As String = "AND"

        If sPassword = "" Then
            sPassword = JEGCore.Password
        End If

        Try
            aColumns = Split(sFilterColumns, "|")
            aValues = Split(sFilterValues, "|")

            'Set comparison variable
            If bSort = True Then
                sSort = " ORDER BY " & sColumn
            End If

            If bWholeWord = True Then
                sComparison = "='"
                sEndComparison = ""
            Else
                sComparison = " LIKE '%"
                sEndComparison = "%"
            End If

            If bOR = True Then
                sOR = "OR"
            End If

            'Execute the select statement
            For i = 0 To UBound(aColumns)
                If i = 0 Then
                    sSQL = "SELECT * FROM [" & sTable & "] WHERE " & aColumns(i) & sComparison & aValues(i) & sEndComparison & "'"
                Else
                    sSQL = sSQL & " " & sOR & " " & aColumns(i) & sComparison & aValues(i) & sEndComparison & "'"
                End If
            Next
            sSQL = sSQL & sSort

            'cnSQL = New OleDbConnection("Provider=" & My.Settings.DBProvider & ";Data Source=" & sDBFullName & ";Jet OLEDB:Database Password=" & sPassword & ";Persist Security Info=False")
            'cnSQL.Open()


            cnSQL = New OdbcConnection("Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Replace(UCase(sDBFullName), ".ACCDB", ".MDB") & ";PWD=" & sPassword)
            cnSQL.Open()

            Dim da As New OdbcDataAdapter
            Dim cmd As New OdbcCommand

            With cmd
                .CommandText = sSQL
                .CommandType = CommandType.Text
                .Connection = cnSQL
                .CommandTimeout = 10000
            End With

            da.SelectCommand = cmd
            da.Fill(dsReturn, sTable)

            'da = New OleDb.OleDbDataAdapter(sSQL, cnSQL)
            'da.Fill(dsReturn, sTable)

            cnSQL.Close()
            cnSQL.Dispose()
            da.Dispose()
            da = Nothing
            cnSQL = Nothing
            GoTo ExitFunction
        Catch e As OleDbException
            MsgBox(e.Message, MsgBoxStyle.Critical, "oledb Error")
        Catch e As Exception
            MsgBox(e.Message, MsgBoxStyle.Critical, "General Error")
        End Try
ExitFunction:
        Return dsReturn
        dsReturn.Dispose()
        dsReturn = Nothing
    End Function

    Shared Function GetTables(ByVal sDBFullName As String, Optional ByVal sPassword As String = "") As ArrayList
        'List the Tables in the Database.
        Dim cnSQL As OleDbConnection = Nothing
        Dim cn As OdbcConnection = Nothing
        Dim ds As New DataSet
        Dim alReturn As ArrayList = New ArrayList()
        Dim schemaTable As DataTable
        Dim dr As DataRow
        Dim bOLEDB As Boolean = False

        If sPassword = "" Then
            sPassword = JEGCore.Password
        End If

        If My.Computer.FileSystem.FileExists(sDBFullName) = False Then
            GetTables = Nothing
            Exit Function
        End If

        'Open Database connection
        cnSQL = New OleDbConnection("Provider=" & My.Settings.DBProvider & ";Data Source=" & sDBFullName & ";Jet OLEDB:Database Password=" & sPassword & ";Persist Security Info=False")
        cnSQL.Open()
        schemaTable = cnSQL.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})

        'If My.Computer.FileSystem.FileExists("C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE14\ACEOLEDB.DLL") Then
        '    cnSQL = New OleDbConnection("Provider=" & My.Settings.DBProvider & ";Data Source=" & sDBFullName & ";Jet OLEDB:Database Password=" & sPassword & ";Persist Security Info=False")
        '    cnSQL.Open()
        '    schemaTable = cnSQL.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
        '    bOLEDB = True
        'ElseIf My.Computer.FileSystem.FileExists("C:\Windows\SYSWOW64\ODBCJT32.DLL") Then
        '    cn = New OdbcConnection("Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Replace(UCase(sDBFullName), ".ACCDB", ".MDB") & ";PWD=" & sPassword)
        '    cn.Open()
        '    schemaTable = cn.GetSchema()
        'Else
        '    GoTo ExitFunction
        '    alReturn = Nothing
        'End If

        Try
            For Each dr In schemaTable.Rows
                alReturn.Add(dr("TABLE_NAME"))
            Next
            If bOLEDB = True Then
                cnSQL.Close()
                cnSQL.Dispose()
            Else
                cn.Close()
                cn.Dispose()
            End If

        Catch x As OleDbException
            alReturn = Nothing
        End Try
ExitFunction:
        Return alReturn
    End Function

    Shared Function TableExists(ByVal sDBFullName As String, ByVal sTable As String, Optional ByVal sPassword As String = "") As Boolean
        'Check if specified table exists in database
        Dim aTables As ArrayList
        Dim i As Long
        Dim bReturn As Boolean = False

        If My.Computer.FileSystem.FileExists(sDBFullName) = False Then
            GoTo ExitFunction
        End If

        If sPassword = "" Then
            sPassword = JEGCore.Password
        End If

        aTables = GetTables(sDBFullName, sPassword)
        For i = 0 To (aTables.Count - 1)
            If UCase(aTables(i)) = UCase(sTable) Then
                bReturn = True
                Exit For
            End If
        Next
ExitFunction:
        Return bReturn
    End Function

    Shared Function ColumnExists(ByVal sDBFullName As String, ByVal sTable As String, ByVal sColumn As String, Optional ByVal sPassword As String = "") As Boolean
        'Check if specified column exists in database
        Dim aColumns As Array
        Dim i As Long
        Dim bReturn As Boolean = False

        If My.Computer.FileSystem.FileExists(sDBFullName) = False Then
            GoTo ExitFunction
        End If

        If sPassword = "" Then
            sPassword = JEGCore.Password
        End If

        aColumns = SCEAccess.GetFieldsDB(sDBFullName, sTable, sPassword)
        For i = 0 To UBound(aColumns)
            If UCase(aColumns(i)) = UCase(sColumn) Then
                bReturn = True
                Exit For
            End If
        Next
ExitFunction:
        Return bReturn
    End Function

    Shared Sub ValidateTable(ByVal sDBFullName As String, ByVal sTable As String, Optional ByVal sDefinitionDB As String = "", Optional ByVal sPassword As String = "")
        'validate columns with LUTables, create table if not exist or columns if not exist

        If sDefinitionDB = "" Then
            sDefinitionDB = JEGCore.SCECoreDB
        End If

        If sPassword = "" Then
            sPassword = JEGCore.Password
        End If

        Dim sColumns As String = GetTableColumns(sTable, , sDefinitionDB)

        If sColumns = "" Then
            GoTo ExitSub
        End If

        'Create table if not exist
        If TableExists(sDBFullName, sTable, sPassword) = False Then
            CreateTable(sDBFullName, sTable, sColumns, sPassword)
            GoTo ExitSub
        End If

        Dim aColumns As Array = Split(sColumns, "|")
        Dim sColumn As String
        sColumns = ""

        For Each sColumn In aColumns
            If ColumnExists(sDBFullName, sTable, sColumn, sPassword) = False Then
                If sColumns = "" Then
                    sColumns = sColumn
                Else
                    sColumns = sColumns & "|" & sColumn
                End If
            End If
        Next

        'Create columns if not exist
        If sColumns <> "" Then
            CreateColumns(sDBFullName, sTable, sColumns, sPassword)
        End If
ExitSub:
    End Sub

    Shared Function GetAllDB(ByVal sDBFullName As String, ByVal sTable As String, Optional ByVal sOrderColumn As String = "", Optional ByVal sPassword As String = "") As DataSet
        'Return entire Database Table
        Dim cnSQL As OleDbConnection
        Dim sSQL As String
        Dim dsReturn As New DataSet '= Nothing
        Dim da As OleDb.OleDbDataAdapter
        Dim dac As OdbcDataAdapter
        Dim sOrderBy As String = ""
        Dim sOrigTable As String = stable

        'Dim objRecordSetx
        'Dim TestRec2 As Integer

        'Check Database exists
        If My.Computer.FileSystem.FileExists(sDBFullName) = False Then
            dsReturn = Nothing
            GoTo ExitFunction
        End If

        If sPassword = "" Then
            sPassword = JEGCore.Password
        End If

        sTable = Replace(sTable, " ", "")
        sTable = Replace(sTable, ".", "")
        'Check table exists
        'If TableExists(sDBFullName, sTable, sPassword) = False Then
        '    sOrigTable = Replace(sTable, ".", "")
        '    If TableExists(sDBFullName, sTable, sPassword) = False Then
        '        dsReturn = Nothing
        '        GoTo ExitFunction
        '    End If
        '    dsReturn = Nothing
        '    GoTo ExitFunction
        'End If

        If sOrderColumn <> "" Then
            sOrderBy = " ORDER BY " & sOrderColumn
        End If

        'Open Database connection
        'cnSQL = New OleDbConnection("Provider=" & My.Settings.DBProvider & ";Data Source=" & sDBFullName & ";Jet OLEDB:Database Password=" & sPassword & ";Persist Security Info=False")
        'cnSQL.Open()
        'Execute select statement
        sSQL = "SELECT * FROM [" & sTable & "]" & sOrderBy

        If My.Computer.FileSystem.FileExists("C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE14\ACEOLEDB.DLL") Then
            Try
                cnSQL = New OleDbConnection("Provider=" & My.Settings.DBProvider & ";Data Source=" & sDBFullName & ";Jet OLEDB:Database Password=" & sPassword & ";Persist Security Info=False")
                cnSQL.Open()
                da = New OleDb.OleDbDataAdapter(sSQL, cnSQL)
                da.Fill(dsReturn, sTable)

                cnSQL.Close()
                cnSQL.Dispose()
            Catch ' on error attempt Windows db connection
                Dim cn As OdbcConnection
                cn = New OdbcConnection("Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Replace(UCase(sDBFullName), ".ACCDB", ".MDB") & ";PWD=" & sPassword)
                cn.Open()
                dac = New OdbcDataAdapter(sSQL, cn)
                dac.Fill(dsReturn, sTable)

                cn.Close()
                cn.Dispose()
            End Try
        ElseIf My.Computer.FileSystem.FileExists("C:\Windows\SYSWOW64\ODBCJT32.DLL") Then
            Dim cn As OdbcConnection
            cn = New OdbcConnection("Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Replace(UCase(sDBFullName), ".ACCDB", ".MDB") & ";PWD=" & sPassword)
            cn.Open()
            dac = New OdbcDataAdapter(sSQL, cn)
            dac.Fill(dsReturn, sTable)

            cn.Close()
            cn.Dispose()
        Else
            GoTo ExitFunction
            dsReturn = Nothing
        End If


ExitFunction:
        Return dsReturn
    End Function

    'Sub LMSDBInfoOnly()
    '    Dim objFileSys
    '    Dim objLogFile
    '    Dim strDatabase
    '    Dim objConnection
    '    'Dim objRecordSet
    '    'Dim objRecordSet0
    '    Dim objRecordSetx
    '    Dim strTable
    '    Dim blnFatalError As Boolean
    '    Dim TestRec
    '    Dim TestRec2

    '    Const adOpenStatic = 3
    '    Const adLockOptimistic = 3

    '    objConnection = CreateObject("ADODB.Connection")
    '    'objRecordSet0 = CreateObject("ADODB.Recordset")
    '    'objRecordSet = CreateObject("ADODB.Recordset")
    '    objRecordSetx = CreateObject("ADODB.Recordset")

    '    If objFileSys.FileExists("C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE14\ACEOLEDB.DLL") Then
    '        objLogFile.WriteLine(vbTab & "Microsoft.ACE.OLEDB.12.0 installed")
    '        objLogFile.WriteLine(vbTab & "Opening " & strDatabase & ".accdb")
    '        ' access 2007 accdb connection string
    '        objConnection.Open("Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=" & strDatabase & ".accdb" & ";Persist Security Info=False;")
    '        ' standard error trapping was not working so make a connection to the database and read a record
    '        objRecordSetx.Open("SELECT * FROM " & strTable & " WHERE ADSite = 'JEG'", objConnection, adOpenStatic, adLockOptimistic)
    '        TestRec = objRecordsetx.RecordCount
    '        objRecordSetx = Nothing
    '        If TestRec < 1 Then  ' connection failed to get a record so try windows default
    '            objLogFile.WriteLine(vbTab & "Microsoft.ACE.OLEDB.12.0 record read error - fall back to windows MDB defaults")
    '            objLogFile.WriteLine(vbTab & "Opening " & strDatabase & ".mdb")
    '            objConnection.Open("Driver= {Microsoft Access Driver (*.mdb)};DBQ=" & strDatabase & ".mdb")
    '        End If
    '    ElseIf objFileSys.FileExists("C:\Windows\SYSWOW64\ODBCJT32.DLL") Then   ' office driver not installed so try windows default
    '        objLogFile.WriteLine(vbTab & "Microsoft.ACE.OLEDB.12.0 not installed - fall back to windows MDB defaults")
    '        objLogFile.WriteLine(vbTab & "Opening " & strDatabase & ".mdb")
    '        objConnection.Open("Driver= {Microsoft Access Driver (*.mdb)};DBQ=" & strDatabase & ".mdb")
    '        ' test the database connection
    '        objRecordSetx.Open("SELECT * FROM " & strTable & " WHERE ADSite = 'JEG'", objConnection, adOpenStatic, adLockOptimistic)
    '        TestRec2 = objRecordsetx.RecordCount
    '        objRecordSetx = Nothing
    '        If TestRec2 < 1 Then  ' connection failed to get a record so abort
    '            objLogFile.WriteLine(vbTab & "Microsoft Access Driver (*.mdb) record read error - aborting")
    '            objLogFile.WriteLine(vbTab & "Opening " & strDatabase & ".mdb")
    '            objConnection.Open("Driver= {Microsoft Access Driver (*.mdb)};DBQ=" & strDatabase & ".mdb")
    '            blnFatalError = True
    '        End If
    '    Else
    '        objLogFile.WriteLine(vbTab & "Microsoft database driver NOT installed")
    '        blnFatalError = True
    '    End If
    'End Sub

    Shared Sub SetDB(ByVal sDBFullName As String, ByVal sTable As String, ByVal sFields As String, ByVal sValues As String, Optional ByVal sPassword As String = "")
        'Wtites data to specified Database
        Dim cnSQL As OleDbConnection
        Dim sSQL As String = Nothing
        Dim aFields As Array = Nothing
        Dim aValues As Array = Nothing
        Dim i As Long

        If My.Computer.FileSystem.FileExists(sDBFullName) = False Then
            Exit Sub
        End If

        If sPassword = "" Then
            sPassword = JEGCore.Password
        End If

        If InStr(sFields, "|") > 0 Then
            aFields = Split(sFields, "|")
            sFields = ""
        End If
        If InStr(sValues, "|") > 0 Then
            aValues = Split(sValues, "|")
            sValues = ""
        End If
        'Open Database connection
        cnSQL = New OleDbConnection("Provider=" & My.Settings.DBProvider & ";Data Source=" & sDBFullName & ";Jet OLEDB:Database Password=" & sPassword & ";Persist Security Info=False")

        'Build ADO statement
        sSQL = "INSERT INTO [" & sTable & "] ("
        If sFields = "" Then
            sSQL = sSQL & aFields(i)
            For i = 1 To UBound(aFields)
                sSQL = sSQL & Chr(44) & "[" & aFields(i) & "]"
            Next
        Else
            sSQL = sSQL & sFields
        End If
        sSQL = sSQL & ") VALUES ("
        If sValues = "" Then
            For i = 0 To UBound(aValues)
                If Trim(aValues(i)) = "" Then aValues(i) = " "
                sSQL = sSQL & "'" & aValues(i) & "', "
            Next
            If UBound(aValues) <> UBound(aFields) Then
                For i = i To UBound(aFields)
                    sSQL = sSQL & "' ', "
                Next
            End If
        Else
            sSQL = sSQL & "'" & sValues & "', "
        End If
        sSQL = Left(sSQL, (Len(sSQL) - 2)) & ")"

        Dim myCommand As New OleDbCommand(sSQL)
        myCommand.Connection = cnSQL
        cnSQL.Open()
        myCommand.ExecuteNonQuery()
        myCommand.Connection.Close()
        myCommand.Dispose()
        cnSQL.Close()
        cnSQL.Dispose()
        myCommand = Nothing
        cnSQL = Nothing
    End Sub

    Shared Sub UpdateFromDataSet(ByVal sDBFullName As String, ByVal sTable As String, ByVal dsDataset As DataSet, Optional ByVal sPassword As String = "")

        'pjrxxx - displays error for file not found

        'Dim dataSet As DataSet = New DataSet
        'Dim dt As New DataTable
        'Using connection As New OleDbConnection("Provider=" & My.Settings.DBProvider & ";Data Source=" & sDBFullName & ";Jet OLEDB:Database Password=" & sPassword & ";Persist Security Info=False")
        '    connection.Open()
        '    Dim adapter As New OleDbDataAdapter()

        '    adapter.SelectCommand = New OleDbCommand("SELECT * FROM " & sTable & ";", connection)

        '    Dim builder As OleDbCommandBuilder = New OleDbCommandBuilder(adapter)

        '    adapter.Fill(dt)

        '    'MsgBox(dataSet.Tables(0).Rows.Count)

        '    'Code to modify the data in the DataSet here.
        '    'Dim id As Integer = 300

        '    'For i As Integer = 0 To dataSet.Tables(0).Rows.Count - 1
        '    '    dataSet.Tables(0).Rows(i).Item(0) = id
        '    '    id = id + 1
        '    'Next

        '    ' Without the OleDbCommandBuilder this line would fail.
        '    'builder.GetUpdateCommand()
        '    'adapter.Update(dataSet)

        '    For Each oRow As DataRow In dsDataset.Tables(0).Rows
        '        dt.ImportRow(oRow)
        '        'dataSet.Tables(0).AcceptChanges()
        '    Next

        '    Try

        '        adapter.Update(dt)

        '    Catch x As Exception
        '        ' Error during Update, add code to locate error, reconcile  
        '        ' and try to update again. 
        '    End Try
        '    connection.Close()
        'End Using

    End Sub

    Shared Sub RegisterDataSet(ByVal sDBFullName As String, ByVal sTable As String, ByVal dsDataSet As DataSet, Optional ByVal sPassword As String = "")
        'Wtites data to specified Database
        Dim cnSQL As OleDbConnection = Nothing
        Dim sSQL As String = Nothing
        Dim aFields As Array = Nothing
        Dim aValues As Array = Nothing
        Dim i As Long
        Dim sFields As String
        Dim sValues As String
        Dim oRow As DataRow

        If My.Computer.FileSystem.FileExists(sDBFullName) = False Then
            Exit Sub
        End If

        If sPassword = "" Then
            sPassword = JEGCore.Password
        End If

        sfields = SCEDataSet.GetColumnsAsString(dsDataSet.Tables(0))
        aFields = Split(sFields, "|")

        If SCEAccess.TableExists(sDBFullName, sTable) = False Then
            SCEAccess.CreateTable(sDBFullName, sTable, sFields)
        End If
        sFields = ""

        'Open Database connection
        cnSQL = New OleDbConnection("Provider=" & My.Settings.DBProvider & ";Data Source=" & sDBFullName & ";Jet OLEDB:Database Password=" & sPassword & ";Persist Security Info=False")

        For Each oRow In dsDataSet.Tables(0).Rows
            sValues = SCEVB.DataRowToString(oRow)
            aValues = Split(sValues, "|")
            sValues = ""

            'Build ADO statement
            sSQL = "INSERT INTO [" & sTable & "] ("
            If sFields = "" Then
                sSQL = sSQL & aFields(0)
                For i = 1 To UBound(aFields)
                    sSQL = sSQL & Chr(44) & "[" & aFields(i) & "]"
                Next
            Else
                sSQL = sSQL & sFields
            End If
            sSQL = sSQL & ") VALUES ("
            If sValues = "" Then
                For i = 0 To UBound(aValues)
                    If Trim(aValues(i)) = "" Then aValues(i) = " "
                    sSQL = sSQL & "'" & aValues(i) & "', "
                Next
                If UBound(aValues) <> UBound(aFields) Then
                    For i = i To UBound(aFields)
                        sSQL = sSQL & "' ', "
                    Next
                End If
            Else
                sSQL = sSQL & "'" & sValues & "', "
            End If
            sSQL = Left(sSQL, (Len(sSQL) - 2)) & ")"

            Dim myCommand As New OleDbCommand(sSQL)
            myCommand.Connection = cnSQL
            cnSQL.Open()
            myCommand.ExecuteNonQuery()
            myCommand.Connection.Close()
            myCommand.Dispose()
            myCommand = Nothing
        Next

        cnSQL.Close()
        cnSQL.Dispose()
        cnSQL = Nothing
    End Sub

    Shared Sub DelDB(ByVal sDBFullName As String, ByVal sTable As String, ByVal sFields As String, ByVal sValues As String, Optional ByVal sPassword As String = "")
        'delete data from specified Database
        Dim cnSQL As OleDbConnection
        Dim sSQL As String
        Dim aFields As Array = Nothing
        Dim aValues As Array = Nothing
        Dim i As Long

        If My.Computer.FileSystem.FileExists(sDBFullName) = False Then
            Exit Sub
        End If

        If sPassword = "" Then
            sPassword = JEGCore.Password
        End If

        If InStr(sFields, "|") > 0 Then
            aFields = Split(sFields, "|")
            'sFields = ""
        End If

        If InStr(sValues, "|") > 0 Then
            aValues = Split(sValues, "|")
            sValues = ""
        End If

        'Open Database connection
        cnSQL = New OleDbConnection("Provider=" & My.Settings.DBProvider & ";Data Source=" & sDBFullName & ";Jet OLEDB:Database Password=" & sPassword & ";Persist Security Info=False")

        'Build ADO statement
        sSQL = "DELETE * FROM [" & sTable & "] WHERE "
        If InStr(sFields, "|") > 0 Then
            sSQL = sSQL & aFields(0) & "='" & aValues(0) & "'"
            For i = 1 To UBound(aFields)
                sSQL = sSQL & " And " & aFields(i) & "='" & aValues(i) & "'"
            Next
        Else
            sSQL = sSQL & sFields & "='" & sValues & "'"
        End If

        'Execute select statement
        Dim myCommand As New OleDbCommand(sSQL)
        myCommand.Connection = cnSQL
        cnSQL.Open()
        myCommand.ExecuteNonQuery()
        myCommand.Connection.Close()
        myCommand.Dispose()
        myCommand = Nothing
        cnSQL.Close()
        cnSQL.Dispose()
        cnSQL = Nothing
    End Sub

    Shared Function GetFieldsDB(ByVal sDBFullName As String, ByVal sTableName As String, Optional ByVal sPassword As String = "") As Array
        'List the Field names in the specified Table
        Dim cnSQL As OleDbConnection
        Dim sSQL As String
        Dim aReturn() As String = Nothing
        Dim i As Long
        Dim ds As New DataSet
        Dim da As OleDb.OleDbDataAdapter
        Dim oColumn As DataColumn

        'Check Datadabase exists
        If My.Computer.FileSystem.FileExists(sDBFullName) = False Then
            GoTo ExitFunction
        End If

        If sPassword = "" Then
            sPassword = JEGCore.Password
        End If

        'Open Database connection
        cnSQL = New OleDbConnection("Provider=" & My.Settings.DBProvider & ";Data Source=" & sDBFullName & ";Jet OLEDB:Database Password=" & sPassword & ";Persist Security Info=False")

        'Execute select statement
        sSQL = "SELECT * FROM [" & sTableName & "]"

        'Execute select statement
        Dim myCommand As New OleDbCommand(sSQL)
        myCommand.Connection = cnSQL
        cnSQL.Open()
        da = New OleDb.OleDbDataAdapter(sSQL, cnSQL)
        da.Fill(ds, sTableName)

        For Each oColumn In ds.Tables(0).Columns
            i = i + 1
            ReDim Preserve aReturn(i)
            aReturn(i) = oColumn.Caption.ToString
        Next
        myCommand.ExecuteNonQuery()
        myCommand.Connection.Close()

        cnSQL.Close()
        cnSQL.Dispose()
        cnSQL = Nothing
ExitFunction:
        If i = 0 Then
            ReDim aReturn(0)
        End If

        Return aReturn
    End Function

    Shared Function GetFieldDB(ByVal sDBFullName As String, ByVal sTableName As String, ByVal sField As String, ByVal sValue As String, _
        ByVal sGetField As String, Optional ByVal sPassword As String = "") As String
        'List the Field names in the specified Table
        Dim cnSQL As OleDbConnection
        Dim sSQL As String
        Dim aFields() As String = Nothing
        Dim ds As New DataSet
        Dim da As OleDb.OleDbDataAdapter

        'Check Database exists
        If My.Computer.FileSystem.FileExists(sDBFullName) = False Then
            GoTo ExitFunction
        End If

        If sPassword = "" Then
            sPassword = JEGCore.Password
        End If

        'Open Database connection
        cnSQL = New OleDbConnection("Provider=" & My.Settings.DBProvider & ";Data Source=" & sDBFullName & ";Jet OLEDB:Database Password=" & sPassword & ";Persist Security Info=False")

        'Execute select statement
        sSQL = "SELECT " & sGetField & " FROM [" & sTableName & "] WHERE "
        sSQL = sSQL & sField & "='" & sValue & "'"

        'Execute select statement
        Dim myCommand As New OleDbCommand(sSQL)
        myCommand.Connection = cnSQL
        cnSQL.Open()
        da = New OleDb.OleDbDataAdapter(sSQL, cnSQL)
        da.Fill(ds, sTableName)

        If ds.Tables(0).Rows(0).Item(sGetField) Is System.DBNull.Value Then
            Return ""
        Else
            Return ds.Tables(0).Rows(0).Item(sGetField)
        End If

        myCommand.ExecuteNonQuery()
        myCommand.Connection.Close()

        cnSQL.Close()
        cnSQL.Dispose()
        cnSQL = Nothing
        Exit Function
ExitFunction:
        Return Nothing
    End Function

    Shared Sub CreateTable(ByVal sDBFullName As String, ByVal sTableName As String, ByVal sData As String, Optional ByVal sPassword As String = "")
        'Create a Table in the Database.
        Dim cnSQL As OleDbConnection
        Dim sSQL As String
        Dim aFields() As String = Nothing
        Dim ds As New DataSet
        Dim da As OleDb.OleDbDataAdapter
        Dim i As Integer
        Dim aSplit As Array
        Dim sFields As String = ""

        'Check Database exists
        If My.Computer.FileSystem.FileExists(sDBFullName) = False Then
            GoTo ExitSub
        End If

        If sPassword = "" Then
            sPassword = JEGCore.Password
        End If

        'Open Database connection
        cnSQL = New OleDbConnection("Provider=" & My.Settings.DBProvider & ";Data Source=" & sDBFullName & ";Jet OLEDB:Database Password=" & sPassword & ";Persist Security Info=False")

        'Execute select statement
        aSplit = Split(sData, "|")

        For i = 0 To UBound(aSplit)
            If sFields <> "" Then sFields = sFields & ","
            sFields = sFields & "[" & aSplit(i) & "]" & Space(1) & "VARCHAR(255)"
        Next
        sFields = sFields
        On Error GoTo Err
Retry:
        'Execute Table creation
        sSQL = "CREATE TABLE [" & sTableName & "](" & sFields & ")"

        'Execute select statement
        Dim myCommand As New OleDbCommand(sSQL)
        myCommand.Connection = cnSQL
        cnSQL.Open()
        da = New OleDb.OleDbDataAdapter(sSQL, cnSQL)
        da.Fill(ds, sTableName)
        cnSQL.Close()
        cnSQL.Dispose()
        cnSQL = Nothing
        Exit Sub
Err:
        If Err.Number = -2147217900 Then 'Table already exists
            DeleteTable(sDBFullName, sTableName)
            GoTo Retry
        Else
            MsgBox(Err.Number & ": " & Err.Description, vbCritical, My.Settings.MsgBoxCaption)
        End If

        'Close Database connection
        cnSQL.Close()
        cnSQL.Dispose()
        cnSQL = Nothing
ExitSub:
    End Sub

    Shared Sub DeleteTable(ByVal sDBFullName As String, ByVal sTableName As String, Optional ByVal sPassword As String = "")
        'Delete a Table in the Database.
        Dim cnSQL As OleDbConnection
        Dim sSQL As String
        Dim aFields() As String = Nothing
        Dim ds As New DataSet
        Dim da As OleDb.OleDbDataAdapter

        'Check Database exists
        If My.Computer.FileSystem.FileExists(sDBFullName) = False Then
            GoTo ExitSub
        End If
        If TableExists(sDBFullName, sTableName, sPassword) = False Then
            GoTo ExitSub
        End If
        If sPassword = "" Then
            sPassword = JEGCore.Password
        End If

        'Open Database connection
        cnSQL = New OleDbConnection("Provider=" & My.Settings.DBProvider & ";Data Source=" & sDBFullName & ";Jet OLEDB:Database Password=" & sPassword & ";Persist Security Info=False")

        'Execute Table creation
        sSQL = "DROP TABLE [" & sTableName & "]"

        'Execute select statement
        Dim myCommand As New OleDbCommand(sSQL)
        myCommand.Connection = cnSQL
        cnSQL.Open()
        da = New OleDb.OleDbDataAdapter(sSQL, cnSQL)
        da.Fill(ds, sTableName)
        On Error Resume Next
        myCommand.ExecuteNonQuery()
        myCommand.Connection.Close()
        On Error GoTo 0
        cnSQL.Close()
        cnSQL.Dispose()
        cnSQL = Nothing
ExitSub:
    End Sub

    Shared Sub CreateDB(ByVal sDBFullName As String, Optional ByVal sTemplateDBFullName As String = "")
        'Creates new Area Register in specified path
        Dim sPath As String

        sPath = FileParse(sDBFullName, FileParser.Path)
        'Create folder is specified path doesn't exist
        If My.Computer.FileSystem.DirectoryExists(sPath) = False Then
            My.Computer.FileSystem.CreateDirectory(sPath)
        End If
        If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
        'Build source database path
        If sTemplateDBFullName = "" Then
            sTemplateDBFullName = SCEXML.GetConfigurationValue("MSTNSCERoot") & "\Templates\Template.accdb"
        End If
        'Check source database exists
        If My.Computer.FileSystem.FileExists(sTemplateDBFullName) Then
            FileCopy(sTemplateDBFullName, sDBFullName)
        End If
    End Sub

    ''    Shared Sub CompactDB(ByVal sDBFullName As String)
    ''        'Compact the specified database
    ''        Dim JRO As JRO.JetEngine

    ''        On Error GoTo Err
    ''        JRO = New JRO.JetEngine
    ''        'Compact database
    ''        If My.Computer.FileSystem.FileExists(sDBFullName) = True Then
    ''            JRO.CompactDatabase("Provider=" & my.settings.dbprovider & ";Data Source=" & sDBFullName & ";Jet OLEDB:Database Password=" & SCECore.Password, _
    ''            "Provider=" & my.settings.dbprovider & ";Data Source=" & FileParse(sDBFullName, FileParser.Path) & "\" & FileParse(sDBFullName, FileParser.FileName) & "_Compact.accdb" & ";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Password=" & SCECore.Password)
    ''        End If
    ''        'Delete current database
    ''        Kill(sDBFullName)
    ''        'Rename compacted database to original filename
    ''        Rename(FileParse(sDBFullName, FileParser.Path) & "\" & FileParse(sDBFullName, FileParser.FileName) & "_Compact.accdb", sDBFullName)
    ''Err:    'db opened by someone exclusively
    ''    End Sub

    Shared Sub UpdateTableColumns(ByVal sDBFullName As String, ByVal sTableName As String, ByVal sDatabaseType As String)
        'Update columns in specified table
        Dim oRow As DataRow
        Dim i As Integer
        Dim j As Integer
        Dim aSplit As Array
        Dim aFields() As String
        Dim bFound As Boolean
        Dim sNewFields As String = ""

        oRow = GetRowDB(JEGCore.SCECoreDB, "LUTables", "TableName", "TableName", sTableName, "Database", sDatabaseType, True, True, True)
        If oRow Is Nothing Then
            GoTo ExitSub
        Else
            If TableExists(sDBFullName, sTableName) = False Then
                CreateTable(sDBFullName, sTableName, oRow.Item("Columns"))
            Else
                aFields = GetFieldsDB(sDBFullName, sTableName)
                aSplit = Split(oRow.Item("Columns"), "|")
                For i = 0 To UBound(aSplit)
                    bFound = False
                    For j = 0 To UBound(aFields)
                        If aFields(j) = aSplit(i) Then
                            bFound = True
                        End If
                    Next
                    If bFound = False Then 'create column
                        If sNewFields = "" Then
                            sNewFields = aSplit(i)
                        Else
                            sNewFields = sNewFields & "|" & aSplit(i)
                        End If
                    End If
                Next
                If sNewFields <> "" Then
                    CreateColumns(sDBFullName, sTableName, sNewFields)
                End If
            End If
        End If
ExitSub:
    End Sub

    Shared Sub CreateColumns(ByVal sDBFullName As String, ByVal sTableName As String, ByVal sColumns As String, Optional ByVal sPassword As String = "")
        'Create columns in specified Table in the Database.
        Dim cnSQL As OleDbConnection
        Dim sSQL As String
        Dim aFields() As String = Nothing
        Dim ds As New DataSet
        Dim da As OleDb.OleDbDataAdapter
        Dim i As Integer
        Dim aSplit As Array
        Dim sFields As String = ""

        'Check Database exists
        If My.Computer.FileSystem.FileExists(sDBFullName) = False Then
            GoTo ExitSub
        End If

        If sPassword = "" Then
            sPassword = JEGCore.Password
        End If

        'Open Database connection
        cnSQL = New OleDbConnection("Provider=" & My.Settings.DBProvider & ";Data Source=" & sDBFullName & ";Jet OLEDB:Database Password=" & sPassword & ";Persist Security Info=False")

        'Execute select statement
        aSplit = Split(sColumns, "|")

        For i = 0 To UBound(aSplit)
            If sFields <> "" Then sFields = sFields & ","
            sFields = sFields & "[" & aSplit(i) & "]" & Space(1) & "VARCHAR(255)"
        Next
        sFields = sFields
        On Error GoTo Err
Retry:
        'Execute Table creation
        sSQL = "ALTER TABLE [" & sTableName & "] ADD " & sFields

        'Execute select statement
        Dim myCommand As New OleDbCommand(sSQL)
        myCommand.Connection = cnSQL
        cnSQL.Open()
        da = New OleDb.OleDbDataAdapter(sSQL, cnSQL)
        da.Fill(ds, sTableName)
        cnSQL.Close()
        cnSQL.Dispose()
        cnSQL = Nothing
        Exit Sub
Err:
        If Err.Number = -2147217900 Then 'Table already exists
            DeleteTable(sDBFullName, sTableName)
            GoTo Retry
        Else
            MsgBox(Err.Number & ": " & Err.Description, vbCritical, My.Settings.MsgBoxCaption)
        End If

        'Close Database connection
        cnSQL.Close()
        cnSQL.Dispose()
        cnSQL = Nothing
ExitSub:
    End Sub

    Public Shared Sub UpdateField(ByVal sDBFullName As String, ByVal sTable As String, ByVal sField As String, ByVal sValue As String, ByVal sFilterFields As String, ByVal sFilterValues As String, Optional ByVal sPassword As String = "")
        'Updates data to specified Database
        Dim cnSQL As OleDbConnection
        Dim sSQL As String = ""
        Dim aFields As Array = Nothing
        Dim aValues As Array = Nothing

        'Check if Database exists
        If My.Computer.FileSystem.FileExists(sDBFullName) = False Then
            Exit Sub
        End If

        If sPassword = "" Then
            sPassword = JEGCore.Password
        End If

        'Open Database connection
        cnSQL = New OleDbConnection("Provider=" & My.Settings.DBProvider & ";Data Source=" & sDBFullName & ";Jet OLEDB:Database Password=" & sPassword & ";Persist Security Info=False")

        'Build ADO statement
        If sFilterFields = "ID" Then
            sSQL = "UPDATE [" & sTable & "] SET " & sField & " = '" & sValue & "' WHERE " & sFilterFields & " = " & sFilterValues & ""
        ElseIf InStr(sFilterFields, "|") > 0 Then 'multiple fields
            aFields = Split(sFilterFields, "|")
            aValues = Split(sFilterValues, "|")
            Dim i As Integer

            For i = 0 To UBound(aFields)
                If sSQL = "" Then
                    sSQL = "UPDATE [" & sTable & "] SET " & sField & " = '" & sValue & "' WHERE " & aFields(i) & " = '" & aValues(i) & "'"
                Else
                    sSQL = sSQL & " AND " & aFields(i) & " = '" & aValues(i) & "'"
                End If
            Next
        Else
            sSQL = "UPDATE [" & sTable & "] SET " & sField & " = '" & sValue & "' WHERE " & sFilterFields & " = '" & sFilterValues & "'"
        End If

        'Execute select statement
        Dim myCommand As New OleDbCommand(sSQL)
        myCommand.Connection = cnSQL
        cnSQL.Open()
        myCommand.ExecuteNonQuery()
        myCommand.Connection.Close()
        myCommand.Dispose()
        cnSQL.Close()
        cnSQL.Dispose()
        myCommand = Nothing
        cnSQL = Nothing
    End Sub

    Public Shared Function SearchDB(ByVal sDBFullName As String, ByVal sTable As String, ByVal sValue As String, Optional ByVal bSort As Boolean = False, _
        Optional ByVal bWholeWord As Boolean = False) As DataSet
        'Return specified column from table with query
        Dim cnSQL As OleDbConnection
        Dim sSQL As String
        Dim dsReturn As DataSet = Nothing
        Dim da As OleDb.OleDbDataAdapter
        Dim aColumns As Array
        Dim i As Integer

        aColumns = GetFieldsDB(sDBFullName, sTable)
        sSQL = "SELECT * FROM [" & sTable & "] WHERE " & aColumns(0) & " LIKE '%" & sValue & "'"

        For i = 1 To UBound(aColumns)
            sSQL = sSQL & " OR " & aColumns(i) & " LIKE '%" & sValue & "'"
        Next

        Try
            'Open Database Connection for Access 2007
            cnSQL = New OleDbConnection("Provider=" & My.Settings.DBProvider & ";Data Source=" & sDBFullName & ";Jet OLEDB:Database Password=" & JEGCore.Password & ";Persist Security Info=False")
            'Open Darabase Connection for Access 2003
            'cnSQL = New OleDbConnection("Provider=" & my.settings.dbprovider & ";" & "Data Source=" & sDBFullName & ";" & "Jet OLEDB:Database Password=" & cPassword & ";Persist Security Info=False")
            cnSQL.Open()
            da = New OleDb.OleDbDataAdapter(sSQL, cnSQL)
            da.Fill(dsReturn, sTable)

            cnSQL.Close()
            cnSQL.Dispose()
            da.Dispose()
            da = Nothing
            cnSQL = Nothing
            GoTo ExitFunction
        Catch e As OleDbException
            MsgBox(e.Message, MsgBoxStyle.Critical, "oledb Error")
        Catch e As Exception
            MsgBox(e.Message, MsgBoxStyle.Critical, "General Error")
        End Try
ExitFunction:
        Return dsReturn
        dsReturn = Nothing
        dsReturn.Dispose()
    End Function

    Shared Function GetTableColumns(ByVal sTableName As String, Optional ByVal sDatabaseType As String = "", Optional ByVal sDBFullName As String = "", Optional ByVal sPassword As String = "") As String
        'Return columns from SCE Core db
        Dim oRow As DataRow
        Dim sReturn As String = ""

        If sDBFullName = "" Then
            sDBFullName = JEGCore.SCECoreDB
        End If

        If sPassword = "" Then
            sPassword = JEGCore.Password
        End If

        If sDatabaseType = "" Then
            oRow = GetRowDB(sDBFullName, "LUTables", "TableName", "TableName", sTableName, "", "", True, True, True, sPassword)
        Else
            oRow = GetRowDB(sDBFullName, "LUTables", "TableName", "TableName", sTableName, "Database", sDatabaseType, True, True, True, sPassword)
        End If

        If Not oRow Is Nothing Then
            If Not oRow.Item("Columns") Is System.DBNull.Value Then
                sReturn = oRow.Item("Columns")
            End If
        End If

        Return sReturn
    End Function

    Shared Sub UpdateBulk(ByVal sDBFullName As String, ByVal sTable As String, ByVal dsDataSet As DataSet, Optional ByVal sOrderColumn As String = "", Optional ByVal sPassword As String = "")
        'Return entire Database Table
        Dim sSQL As String
        Dim dsReturn As New DataSet '= Nothing
        Dim da As New OleDb.OleDbDataAdapter
        Dim sOrderBy As String = ""
        Dim sOrigTable As String = sTable

        'Check Database exists
        If My.Computer.FileSystem.FileExists(sDBFullName) = False Then
            dsReturn = Nothing
            GoTo ExitFunction
        End If

        sTable = Replace(sTable, " ", "")
        sTable = Replace(sTable, ".", "")
        'Check table exists
        If TableExists(sDBFullName, sTable, sPassword) = False Then
            sOrigTable = Replace(sTable, ".", "")
            If TableExists(sDBFullName, sTable, sPassword) = False Then
                dsReturn = Nothing
                GoTo ExitFunction
            End If
            dsReturn = Nothing
            GoTo ExitFunction
        End If

        If sOrderColumn <> "" Then
            sOrderBy = "ORDER BY " & sOrderColumn
        End If

        If sPassword = "" Then
            sPassword = JEGCore.Password
        End If

        'Create primary key
        If SCEAccess.ColumnExists(sDBFullName, sTable, "ID") = False Then
            SCEAccess.QueryDB(sDBFullName, sTable, "ALTER TABLE " & sTable & " ADD PRIMARY KEY (ID)")
        End If

        Using con = New OleDbConnection("Provider=" & My.Settings.DBProvider & ";Data Source=" & sDBFullName & ";Jet OLEDB:Database Password=" & sPassword & ";Persist Security Info=False")
            con.Open()
            sSQL = "SELECT * FROM [" & sTable & "]"
            da.SelectCommand = New OleDbCommand(sSQL, con)
            da.Fill(dsReturn, sTable)
            Dim CB As OleDbCommandBuilder = New OleDbCommandBuilder(da)
            da.InsertCommand = CB.GetInsertCommand(True)
            'dsReturn.Merge(dsDataSet)
            dsReturn.Tables(0).Merge(dsDataSet.Tables(0))
            'dsReturn = SCEDataSet.Merge(dsReturn, dsDataSet)
            da.Update(dsReturn, sTable)
        End Using
ExitFunction:
    End Sub

    Shared Function GetRowsDBAdvanced(ByVal sDBFullName As String, ByVal sTable As String, ByVal sSQL As String, Optional ByVal sPassword As String = "") As DataSet
        'Return multiple rows
        Dim cnSQL As OleDbConnection
        Dim dsReturn As New DataSet
        Dim da As OleDb.OleDbDataAdapter
        Dim sComparison As String = ""
        Dim sEndComparison As String = ""
        Dim sSort As String = ""
        Dim sOR As String = "AND"

        If sPassword = "" Then
            sPassword = JEGCore.Password
        End If

        Try
            cnSQL = New OleDbConnection("Provider=" & My.Settings.DBProvider & ";Data Source=" & sDBFullName & ";Jet OLEDB:Database Password=" & sPassword & ";Persist Security Info=False")
            cnSQL.Open()

            da = New OleDb.OleDbDataAdapter(sSQL, cnSQL)
            da.Fill(dsReturn, sTable)

            cnSQL.Close()
            cnSQL.Dispose()
            da.Dispose()
            da = Nothing
            cnSQL = Nothing
            GoTo ExitFunction
        Catch e As OleDbException
            MsgBox(e.Message, MsgBoxStyle.Critical, "oledb Error")
        Catch e As Exception
            MsgBox(e.Message, MsgBoxStyle.Critical, "General Error")
        End Try
ExitFunction:
        Return dsReturn
        dsReturn.Dispose()
        dsReturn = Nothing
    End Function

    Public Shared Function QueryDB(ByVal sDBFullName As String, ByVal sTable As String, ByVal sQuery As String, Optional ByVal sPassword As String = "") As DataSet
        'Return results of specified query, connection is read only to avoid user writing to database
        Dim cnSQL As OleDbConnection
        Dim sSQL As String = sQuery
        Dim dsReturn As New DataSet '= Nothing
        Dim da As OleDb.OleDbDataAdapter
        Dim sOrderBy As String = ""

        'Check Database exists
        If My.Computer.FileSystem.FileExists(sDBFullName) = False Then
            dsReturn = Nothing
            GoTo ExitFunction
        End If

        If sPassword = "" Then
            sPassword = JEGCore.Password
        End If

        cnSQL = New OleDbConnection("Provider=" & My.Settings.DBProvider & ";Data Source=" & sDBFullName & ";Jet OLEDB:Database Password=" & sPassword & ";Persist Security Info=False;Mode=Read")
        cnSQL.Open()
        da = New OleDb.OleDbDataAdapter(sSQL, cnSQL)
        Try
            da.Fill(dsReturn, sTable)
        Catch
        End Try

        cnSQL.Close()
        cnSQL.Dispose()
ExitFunction:
        Return dsReturn
    End Function

End Class

'Public Class SCEAccessUpdate

'    Public Shared Sub UpdateDB(ByVal sDBFullName As String, ByVal sTable As String, ByVal dsDataSet As DataSet, Optional ByVal sPassword As String = "")
'        Dim cn As New Odbc.OdbcConnection("Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=shreyas; User=root;Password=;")
'        Dim cmd As Odbc.OdbcCommand
'        Dim adp As Odbc.OdbcDataAdapter

'        If sPassword = "" Then
'            sPassword = SCECore.Password
'        End If

'        cn.Open()
'        cmd = New Odbc.OdbcCommand("Select * from " & sTable, cn)
'        adp = New Odbc.OdbcDataAdapter(cmd)
'        adp.Fill(dsDataSet, sTable)
'        Me.DataGridView1.DataSource = dsDataSet
'        Me.DataGridView1.DataMember = sTable

'        Dim cmdbuilder As New Odbc.OdbcCommandBuilder(adp)
'        Dim i As Integer

'        i = adp.Update(dsDataSet, "trial1")
'    End Sub


'End Class