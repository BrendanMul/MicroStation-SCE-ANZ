Imports JMECore.SCEFile

Public Class SCEClients

    Public Shared dsClients As DataSet
    Public Shared Loaded As Boolean = False
    Public Shared Count As Integer
    Public Shared Countries As Array
    Public Shared Regions As Array

    Public Enum tSettings
        Null
        Client
        Software
        Version
        Owner
        DefaultTemplate
        FullName
        STDCompliance
        STDDoc
        RunFirst
        DisallowAcadSettings
        AcadRegion
        AcadClient
        AcadVersion
        AcadForce
    End Enum

    Public Enum tPaths
        Country
        Client
        Region
        ClientFolderName
        Discipline
        ClientPath
    End Enum

    Shared Sub Initialise()
        'Load all clients into memory
        Dim oTable As DataTable = Nothing
        Dim oPathsTable As New DataTable
        Dim oPathsRow As DataRow
        Dim sClientFolder As String
        Dim sTempPath
        Dim aSplit As Array
        Dim i As Integer

        If JEGCore.bDebug = True Then JEGCore.RecordSequence(System.Reflection.MethodBase.GetCurrentMethod.Module.Name, System.Reflection.MethodBase.GetCurrentMethod.Name())

        If Loaded = True Then
            GoTo ExitSub
        End If

        dsClients = New DataSet
        oPathsTable.TableName = "Paths"
        oPathsTable.Columns.Add("Country") 'Australia
        oPathsTable.Columns.Add("Region") 'WA NT
        oPathsTable.Columns.Add("Client") 'RioTinto 9.2
        oPathsTable.Columns.Add("Discipline") 'Plant Design v8.0

        oPathsTable.Columns.Add("PCFName") 'Australia-WA NT-RioTinto
        oPathsTable.Columns.Add("ClientFolderName") 'Australia-WA NT-RioTinto
        oPathsTable.Columns.Add("ClientPath") 'C:\ProgramData\Jacobs\MicroStation Environment\Clients\Australia-WA NT-RioTinto 9.2
        oPathsTable.Columns.Add("ClientLocation") 'Local\LocalPublic\Network
        oPathsTable.Columns.Add("Current") 'True\False (whether local\localpublic version is same as network)
        oPathsTable.Columns.Add("Status")
        'oPathsTable.Columns.Add("ClientVersion") '9.2
        'oPathsTable.Columns.Add("JacobsVersion") '1.00
        'oPathsTable.Columns.Add("Documentation") 'C:\ProgramData\Jacobs\MicroStation Environment\Clients\Australia-WA NT-RioTinto 9.2\Documentation
        'oPathsTable.Columns.Add("ClientContact") 'John Smith
        'oPathsTable.Columns.Add("JacobsContact") 'Paul Ripp
        'oPathsTable.Columns.Add("Downloadable") 'True\False
        'oPathsTable.Columns.Add("IgnoreJMEConfig") 'True\False

        'pjrxxx - add identifier if client build has been downloaded and/or is not current

        dsClients = SCEDataSet.CreateDataset("Clients", "", True, "LUSettings") 'Create DataSet
        dsClients.Tables.Add(oPathsTable)

        'initialise local client builds
        If My.Computer.FileSystem.DirectoryExists(JEGCore.ClientPath) = True Then
            For Each sTempPath In My.Computer.FileSystem.GetFiles(JEGCore.ClientPath, FileIO.SearchOption.SearchTopLevelOnly, "*.pcf")
                'check for HIDE file to hide client build
                sClientFolder = SCEFile.FileParse(sTempPath, FileParser.FileName)
                If My.Computer.FileSystem.FileExists(JEGCore.ClientPath & "\" & sClientFolder & "\HIDE.txt") = True Then
                    GoTo NextRegion
                End If

                aSplit = Split(sClientFolder, "-")
                oPathsRow = oPathsTable.Rows.Add
                oPathsRow.Item("PCFName") = sClientFolder
                oPathsRow.Item("ClientFolderName") = aSplit(0) & "-" & aSplit(1) & "-" & aSplit(2) 'sClientFolder
                oPathsRow.Item("ClientPath") = JEGCore.ClientPath & "\" & aSplit(0) & "-" & aSplit(1) & "-" & aSplit(2) 'sClientFolder
                oPathsRow.Item("ClientLocation") = "Local"
                oPathsRow.Item("Country") = aSplit(0)
                If UBound(aSplit) >= 1 Then
                    oPathsRow.Item("Region") = aSplit(1)
                End If
                If UBound(aSplit) >= 2 Then
                    oPathsRow.Item("Client") = aSplit(2)
                End If
                If UBound(aSplit) >= 3 Then
                    oPathsRow.Item("Discipline") = aSplit(3)
                    If UBound(aSplit) > 3 Then
                        For i = 4 To UBound(aSplit)
                            oPathsRow.Item("Discipline") = oPathsRow.Item("Discipline") & "-" & aSplit(i)
                        Next
                    End If
                End If
                oPathsRow.AcceptChanges()
NextRegion:
            Next
            oPathsTable.AcceptChanges()
        End If

        'initialise local public client builds
        If My.Computer.FileSystem.DirectoryExists(JEGCore.LocalPublicPath & "\Clients") = True Then
            For Each sTempPath In My.Computer.FileSystem.GetFiles(JEGCore.LocalPublicPath & "\Clients", FileIO.SearchOption.SearchTopLevelOnly, "*.pcf")
                'check for HIDE file to hide client build
                sClientFolder = SCEFile.FileParse(sTempPath, FileParser.FileName)
                If My.Computer.FileSystem.FileExists(JEGCore.LocalPublicPath & "\Clients" & "\" & sClientFolder & "\HIDE.txt") = True Then
                    GoTo NextPublicRegion
                End If

                aSplit = Split(sClientFolder, "-")
                oPathsRow = oPathsTable.Rows.Add
                oPathsRow.Item("PCFName") = sClientFolder
                oPathsRow.Item("ClientFolderName") = aSplit(0) & "-" & aSplit(1) & "-" & aSplit(2) 'sClientFolder
                oPathsRow.Item("ClientPath") = JEGCore.LocalPublicPath & "\Clients" & "\" & aSplit(0) & "-" & aSplit(1) & "-" & aSplit(2) 'sClientFolder
                'oPathsRow.Item("ClientFolderName") = sClientFolder
                'oPathsRow.Item("ClientPath") = JEGCore.LocalPublicPath & "\Clients" & "\" & sClientFolder
                oPathsRow.Item("ClientLocation") = "LocalPublic"
                oPathsRow.Item("Country") = aSplit(0)
                If UBound(aSplit) >= 1 Then
                    oPathsRow.Item("Region") = aSplit(1)
                End If
                If UBound(aSplit) >= 2 Then
                    oPathsRow.Item("Client") = aSplit(2)
                End If
                If UBound(aSplit) >= 3 Then
                    oPathsRow.Item("Discipline") = aSplit(3)
                    If UBound(aSplit) > 3 Then
                        For i = 4 To UBound(aSplit)
                            oPathsRow.Item("Discipline") = oPathsRow.Item("Discipline") & "-" & aSplit(i)
                        Next
                    End If
                End If
                oPathsRow.AcceptChanges()
NextPublicRegion:
            Next
            oPathsTable.AcceptChanges()
        End If

        'initialise network client builds
        If My.Computer.FileSystem.DirectoryExists(JEGCore.NetworkClientPath) = True Then
            For Each sTempPath In My.Computer.FileSystem.GetFiles(JEGCore.NetworkClientPath, FileIO.SearchOption.SearchTopLevelOnly, "*.pcf")
                sClientFolder = SCEFile.FileParse(sTempPath, FileParser.FileName)
                'check for HIDE file to hide client build
                If My.Computer.FileSystem.FileExists(JEGCore.NetworkClientPath & "\" & sClientFolder & "\HIDE.txt") = True Then
                    GoTo NextNetworkRegion
                End If

                aSplit = Split(sClientFolder, "-")
                oPathsRow = oPathsTable.Rows.Add
                oPathsRow.Item("PCFName") = sClientFolder
                oPathsRow.Item("ClientFolderName") = aSplit(0) & "-" & aSplit(1) & "-" & aSplit(2) 'sClientFolder
                oPathsRow.Item("ClientPath") = JEGCore.NetworkClientPath & "\" & aSplit(0) & "-" & aSplit(1) & "-" & aSplit(2) 'sClientFolder
                'oPathsRow.Item("ClientFolderName") = sClientFolder
                'oPathsRow.Item("ClientPath") = JEGCore.NetworkClientPath & "\" & sClientFolder
                oPathsRow.Item("ClientLocation") = "Network"
                oPathsRow.Item("Country") = aSplit(0)
                If UBound(aSplit) >= 1 Then
                    oPathsRow.Item("Region") = aSplit(1)
                End If
                If UBound(aSplit) >= 2 Then
                    oPathsRow.Item("Client") = aSplit(2)
                End If
                If UBound(aSplit) >= 3 Then
                    oPathsRow.Item("Discipline") = aSplit(3)
                    If UBound(aSplit) > 3 Then
                        For i = 4 To UBound(aSplit)
                            oPathsRow.Item("Discipline") = oPathsRow.Item("Discipline") & "-" & aSplit(i)
                        Next
                    End If
                End If
                oPathsRow.AcceptChanges()
NextNetworkRegion:
            Next
            oPathsTable.AcceptChanges()
        End If

        Count = dsClients.Tables("Paths").Rows.Count
        Countries = SCEDataSet.GetColumn(oPathsTable, "Country", True)
        Regions = SCEDataSet.GetColumn(oPathsTable, "Region", True)

        'load client build settings from Settings.ini
        Dim sSettingsFile As String = ""
        Dim sLine As String
        Dim iRow As Integer

        For iRow = 0 To (dsClients.Tables("Paths").Rows.Count - 1)
            sSettingsFile = dsClients.Tables("Paths").Rows(iRow).Item("ClientPath") & "\Settings.ini"
            If My.Computer.FileSystem.FileExists(sSettingsFile) = True Then
                FileOpen(1, sSettingsFile, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
                Do While Not EOF(1)
                    sLine = Trim(LineInput(1))
                    If SCEVB.IsCommented(sLine) = False Then
                        aSplit = Split(sLine, "=")
                        aSplit(0) = Trim(aSplit(0))
                        aSplit(1) = Trim(aSplit(1))
                        If aSplit(1) <> "" Then
                            i = SCEDataSet.GetColumnIndex(dsClients.Tables("Paths"), aSplit(0))
                            If i = -1 Then
                                dsClients.Tables("Paths").Columns.Add(aSplit(0)) 'Add column, allows scalability
                                i = SCEDataSet.GetColumnIndex(dsClients.Tables("Paths"), aSplit(0))
                                dsClients.Tables("Paths").Rows(iRow).Item(i) = aSplit(1)
                            Else
                                dsClients.Tables("Paths").Rows(iRow).Item(i) = aSplit(1)
                            End If
                        End If
                    End If
                Loop
                dsClients.Tables("Paths").Rows(iRow).AcceptChanges()
                FileClose(1)
            End If
        Next
ExitSub:
    End Sub

    Shared Function GetIndex(Optional ByVal sClientFullName As String = "", Optional ByVal sClientName As String = "", Optional ByVal sClientBuildName As String = "") As Integer
        'get index of specified software
        Dim iReturn As Integer = -1
        Dim i As Integer

        For i = 0 To Count - 1
            If sClientFullName <> "" Then
                If UCase(sClientFullName) = UCase(dsClients.Tables(0).Rows(i).Item("Client")) Then
                    iReturn = i
                    Exit For
                End If
            ElseIf sClientName <> "" Then
                If UCase(sClientName) = UCase(dsClients.Tables(1).Rows(i).Item("Client")) Then
                    iReturn = i
                    Exit For
                ElseIf InStr(UCase(dsClients.Tables(1).Rows(i).Item("ClientPath")), UCase(sClientName)) > 0 Then
                    iReturn = i
                    Exit For
                End If
            ElseIf sClientBuildName <> "" Then
                If InStr(UCase(dsClients.Tables(1).Rows(i).Item("PCFName")), UCase(sClientBuildName)) > 0 Then
                    iReturn = i
                    Exit For
                End If
            End If
        Next

        Return iReturn
    End Function

    Public Shared Function GetValue(ByVal iClientIndex As Integer, Optional ByVal oSetting As tSettings = tSettings.Null, Optional ByVal oPath As tPaths = tPaths.Client) As String
        'return value of client column
        Dim sReturn As String = ""
        Dim iIndex As Integer = -1
        Dim i As Integer

        If oSetting <> tSettings.Null Then
            For i = 0 To dsClients.Tables(1).Columns.Count - 1
                If UCase(dsClients.Tables(1).Columns.Item(i).ColumnName) = UCase(oSetting.ToString()) Then
                    iIndex = i
                    Exit For
                End If
            Next
            If iIndex > -1 Then
                If Not dsClients.Tables(1).Rows(iClientIndex).Item(iIndex) Is System.DBNull.Value Then
                    sReturn = dsClients.Tables(1).Rows(iClientIndex).Item(iIndex)
                End If
            End If
            'For i = 0 To dsClients.Tables(0).Columns.Count - 1
            '    If UCase(dsClients.Tables(0).Columns.Item(i).ColumnName) = UCase(oSetting.ToString()) Then
            '        iIndex = i
            '        Exit For
            '    End If
            'Next
            'If iIndex > -1 Then
            '    If Not dsClients.Tables(0).Rows(iClientIndex).Item(iIndex) Is System.DBNull.Value Then
            '        sReturn = dsClients.Tables(0).Rows(iClientIndex).Item(iIndex)
            '    End If
            'End If
        Else
            For i = 0 To dsClients.Tables(1).Columns.Count - 1
                If UCase(dsClients.Tables(1).Columns.Item(i).ColumnName) = UCase(oPath.ToString) Then
                    iIndex = i
                    Exit For
                End If
            Next
            If iIndex > -1 Then
                If Not dsClients.Tables(1).Rows(iClientIndex).Item(iIndex) Is System.DBNull.Value Then
                    sReturn = dsClients.Tables(1).Rows(iClientIndex).Item(iIndex)
                End If
            End If
        End If

        Return sReturn
    End Function

End Class
