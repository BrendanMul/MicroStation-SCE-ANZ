Imports MicroStationSCE.SCEAccess
Imports MicroStationSCE.SCEFavourite
Imports MicroStationSCE.SCEFile
Imports MicroStationSCE.SCEXML
Imports MicroStationSCE.SCEZip

Public Class SCEMicroStationSession
    Private Shared _myInstance As MicroStationDGN.Application
    Private Shared oProcess As Process

    Shared oMSTN As MicroStationDGN.Application = GetInstance()

    Shared Function GetInstance() As MicroStationDGN.Application
        'Get MicroStation instance, only the first opened session is activated
        If _myInstance Is Nothing Then
            _myInstance = New MicroStationDGN.Application 'TODO: Add check if App is already(running)
            If _myInstance.HasActiveDesignFile = False Then
                _myInstance.Quit()
                MsgBox("MicroStation is not currently operating, initiate a MicroStation session first.", MsgBoxStyle.Critical, My.Settings.MsgBoxCaption)
                _myInstance = Nothing
            End If
        End If
        Return _myInstance
ExitFunction:
    End Function

    Shared Function GetMicroStationPath(ByVal sSoftware As String, ByVal sVersion As String) As String
        'Return path to installed software
        Dim dsVersion As DataSet
        Dim dsCoreSoftware As DataSet
        Dim i As Long
        Dim sReturn As String = ""

        dsVersion = SCEAccess.GetRowDB(SCECore.SCECoreDB, "LUSoftware", "RegistryKey", "Software", sSoftware, "", "", False, True)
        If dsVersion Is Nothing Then
            GoTo SkipVersion
        End If
        If dsVersion.Tables(0).Rows.Count <> 0 Then
            If UCase(dsVersion.Tables(0).Rows(i).Item("LoadType")) = "PROGRAM" Then
                'Get path from registry
                For i = 0 To (dsVersion.Tables(0).Rows.Count - 1)
                    sReturn = My.Computer.Registry.GetValue(dsVersion.Tables(0).Rows(i).Item("RegistryKey"), dsVersion.Tables(0).Rows(i).Item("RegistryString"), Nothing)
                    If sReturn <> Nothing Then
                        Exit For
                    End If
                Next
            ElseIf UCase(dsVersion.Tables(0).Rows(i).Item("LoadType")) = "ADDON" Then
                'Get core software
                dsCoreSoftware = SCEAccess.GetRowDB(SCECore.SCECoreDB, "LUSoftware", "RegistryKey", "Software", sSoftware, "", "", False, True)
            End If
        End If
SkipVersion:
        'Check file exists
        'If My.Computer.FileSystem.FileExists(sReturn) = False Then
        '    MsgBox("The following software is required for this client:" & Chr(10) & Chr(10) & _
        '        "Software: " & sSoftware & Chr(10) & "Version: " & sVersion, MsgBoxStyle.Exclamation, My.Settings.MsgBoxCaption)
        'End If

        dsVersion = Nothing
        dsCoreSoftware = Nothing

        Return sReturn
    End Function

    Shared Sub LoadSoftware(ByVal sRegion As String, ByVal sClient As String, _
        Optional ByVal sSoftware As String = "", Optional ByVal sSoftwareVersion As String = "", _
        Optional ByVal sSoftwareAddon As Integer = -1, Optional ByVal sTools As String = "", _
        Optional ByVal iFavouriteClientIndex As String = "", Optional ByVal sDrawing As String = "", _
        Optional ByVal sAutoCADRegion As String = "", Optional ByVal sAutoCADClient As String = "")
        'Execute MicroStation with specified details
        Dim dsVersion As DataSet
        Dim dsSoftware As DataSet = Nothing
        Dim dsClient As DataSet
        Dim sMSFullName As String = Nothing
        Dim sMSPath As String = Nothing
        Dim sLine As String = Nothing
        Dim sConfigFullName As String = Nothing
        Dim iClientIndex As Integer
        'Dim aFolders As Array
        Dim sFolder As String
        Dim vSplit() As String
        Dim bProjectLoaded As Boolean = False
        Dim i As Integer
        Dim j As Integer
        Dim sConfigName As String = ""
        Dim sMSCFGFullName As String
        Dim sMSProject As String

        If SCEClients.Loaded = False Then
            SCEClients.Load()
        End If
        iClientIndex = SCEClients.GetIndex(sClient)
        sMSProject = SCEClients.dsClients.Tables(0).Rows(iClientIndex).Item("Client")

        If iFavouriteClientIndex = -1 Then
            If sDrawing = "" Then 'Client selector used
                sSoftware = SCEClients.dsClients.Tables(0).Rows(iClientIndex).Item("Software")
                sSoftwareVersion = SCEClients.dsClients.Tables(0).Rows(iClientIndex).Item("Version")
                sSoftwareAddon = sSoftwareAddon
            End If 'sdrawing
        Else 'Favourite used
            If SCEFavourite.ClientsLoaded = False Then
                SCEFavourite.LoadClients()
            End If
            iClientIndex = SCEClients.GetIndex(, SCEFavourite.dsFavouriteClients.Tables(0).Rows(iFavouriteClientIndex).Item("Client"))
            sSoftware = SCEClients.dsClients.Tables(0).Rows(iClientIndex).Item("Software")
            sSoftwareAddon = SCEClients.dsClients.Tables(0).Rows(iClientIndex).Item("SoftwareAddon")
            sSoftwareVersion = SCEClients.dsClients.Tables(0).Rows(iClientIndex).Item("SoftwareVersion")
            sRegion = SCEClients.dsClients.Tables(0).Rows(iClientIndex).Item("Region")
            sClient = SCEClients.dsClients.Tables(0).Rows(iClientIndex).Item("Client")
        End If 'favourite
        iClientIndex = SCEClients.GetIndex("", sClient)

        'Set last used software
        My.Computer.Registry.SetValue(SCECore.MSTNSCEKey, "NewLastMSTNAddonSoftware", sSoftwareAddon)

        If sSoftwareAddon = "{None}" Then
            sSoftwareAddon = Nothing
        Else
            dsSoftware = GetRowDB(SCECore.SCECoreDB, "LUSoftware", "Software", "Software", sSoftwareAddon, "", "", False, True)
        End If

        sMSFullName = GetMicroStationPath(sSoftware, sSoftwareVersion)
        If sMSFullName = Nothing Then GoTo ExitSub
        sMSPath = FileParse(sMSFullName, FileParser.Path)
        sMSCFGFullName = sMSPath & "\Config\Appl\AutoSCE.cfg"

        'Delete AutoSCE.cfg file
        dsVersion = GetRowDB(SCECore.SCECoreDB, "LUSoftware", "Software", "Software", sSoftware, "", "", False, False)
        If My.Computer.FileSystem.FileExists(sMSCFGFullName) = True Then
            My.Computer.FileSystem.DeleteFile(sMSCFGFullName)
        End If

        'Create AutoSCE.cfg file
        FileOpen(2, sMSCFGFullName, OpenMode.Output, OpenAccess.Write, OpenShare.Default)

        PrintLine(2, "SCE_CLIENT_REGION = " & sRegion)
        PrintLine(2, "SCE_CLIENT_NAME = " & sClient)
        PrintLine(2, "SCE_SOFTWARE = " & sSoftware)
        PrintLine(2, "SCE_SOFTWAREVERSION = " & sSoftwareVersion)
        PrintLine(2, "SCE_SOFTWAREADDON = " & sSoftwareAddon)
        PrintLine(2, "SCE_TOOLS = " & sTools)
        PrintLine(2, "SCE_ROOT = " & SCECore.SCEPath() & "/")
        PrintLine(2, "SCE_CONFIGROOT = " & SCECore.SCEConfigPath() & "/")
        PrintLine(2, "SCE_WORKINGROOT = " & SCECore.WorkingPath() & "/")
        PrintLine(2, "_SCE_DB = " & Replace(SCECore.SCECoreDB, "\", "/"))
        PrintLine(2, "SCE_NET_PATH = " & Replace(GetConfigurationValue("MSTNNetworkPath"), "\", "/") & "/")
        PrintLine(2, "SCE_LOCAL_TEMP = " & Replace(GetConfigurationValue("MSTNLocalTemp"), "\", "/") & "/")
        PrintLine(2, "_SCE_XMLCONFIGFILE = " & Replace(SCECore.XMLFileName, "\", "/"))

        'Load AutoCAD client config if necessary
        If sAutoCADRegion <> "" Then
            PrintLine(2, "_SCE_LOADACADCONFIG   = " & SCECore.WorkingPath & "\AutoCAD\" & sAutoCADRegion & "\" & sAutoCADClient & "/AutoCAD Config.cfg")
        End If

        'Load verticals
        dsSoftware = SCEAccess.GetAllDB(SCECore.SCECoreDB, "LUSoftware", "Software")
        For j = 0 To dsSoftware.Tables(0).Rows.Count - 1
            If Not dsSoftware.Tables(0).Rows(j).Item("LoadType") Is System.DBNull.Value Then
                If UCase(dsSoftware.Tables(0).Rows(j).Item("LoadType")) = "VERTICAL" Then
                    If UCase(dsSoftware.Tables(0).Rows(j).Item("Software")) = UCase(sSoftware) Then
                        If My.Computer.FileSystem.FileExists(dsSoftware.Tables(0).Rows(j).Item("Location")) = True Then
                            PrintLine(2, "_VERTICAL = " & dsSoftware.Tables(0).Rows(j).Item("Location"))
                            PrintLine(2, "%include $(_VERTICAL) level 3")
                        End If
                        'Include extra configuration
                        If My.Computer.FileSystem.FileExists(SCECore.NetworkPath & "\Software\" & sSoftware & "\Vertical.cfg") = True Then
                            PrintLine(2, "%include " & SCECore.NetworkPath & "\Software\" & sSoftware & "\Vertical.cfg level 3")
                        End If
                    End If
                End If
            End If
        Next

        'Load addons
        Dim aTools As Array
        If sTools <> "" Then
            aTools = Split(sTools, ";")
            dsSoftware = SCEAccess.GetAllDB(SCECore.SCECoreDB, "LUSoftware", "Software")
            For i = 0 To UBound(aTools)
                For j = 0 To dsSoftware.Tables(0).Rows.Count - 1
                    If UCase(dsSoftware.Tables(0).Rows(j).Item("LoadType")) = "ADDON" Then
                        If UCase(dsSoftware.Tables(0).Rows(j).Item("Software")) = UCase(aTools(i)) Then
                            If UCase(dsSoftware.Tables(0).Rows(j).Item("Prerequisite")) = UCase(sSoftware & "|" & sSoftwareVersion) Then
                                If My.Computer.FileSystem.FileExists(dsSoftware.Tables(0).Rows(j).Item("Location")) = True Then
                                    PrintLine(2, "%include " & dsSoftware.Tables(0).Rows(j).Item("Location") & " level 3")
                                End If
                            End If
                        End If
                    End If
                Next
            Next
        End If

        Print(2, My.Resources.SCEConfigFile)
        FileClose(2)

        'Software Addon
        If sSoftwareAddon <> "" Then
            If Not dsSoftware.Tables(0).Rows(0).Item("Location") Is System.DBNull.Value Then
                If Microsoft.VisualBasic.Left(dsSoftware.Tables(0).Rows(0).Item("Location"), 1) = "!" Then
                    vSplit = Split(dsSoftware.Tables(0).Rows(0).Item("Location"), "|")
                    For i = 0 To UBound(vSplit)
                        sConfigFullName = vSplit(i)
                        sConfigFullName = Microsoft.VisualBasic.Right(sConfigFullName, (Len(sConfigFullName) - 1))
                        sConfigName = FileParse(sConfigFullName, FileParser.FullFileName)
                        FileCopy(sConfigFullName, sMSPath & "\Config\Appl\" & sConfigName)
                        If i = 0 Then
                            sSoftwareAddon = Space(1) & Chr(34) & sMSPath & "\Config\Appl\" & sConfigName & Chr(34)
                            sMSCFGFullName = sMSPath & "\Config\Appl\" & sConfigName
                        Else
                            sSoftwareAddon = sSoftwareAddon & Space(1) & Chr(34) & sMSPath & "\Config\Appl\" & sConfigName & Chr(34)
                            sMSCFGFullName = sMSCFGFullName & Chr(44) & sMSPath & "\Config\Appl\" & sConfigName
                        End If
                    Next
                End If
            End If
        End If

        'Tools
        sTools = Replace(sTools, ";", "|") 'Coming from notification icon
        aTools = Split(sTools, "|")
        Dim dsTools As DataSet
        Dim sTool As String
        dsTools = SCEAccess.GetRowDB(SCECore.SCECoreDB, "LUSoftware", "Software", "LoadType", "Addon", "", "", True, True)
        For i = 0 To UBound(aTools)
            For j = 0 To (dsTools.Tables(0).Rows.Count - 1)
                If UCase(aTools(i)) = UCase(dsTools.Tables(0).Rows(j).Item("Software")) Then
                    sTool = Replace(dsTools.Tables(0).Rows(j).Item("Location"), "!", "")
                    If My.Computer.FileSystem.FileExists(sTool) Then
                        FileCopy(sTool, sMSPath & "\Config\Appl\" & FileParse(sTool, FileParser.FullFileName))
                    End If
                End If
            Next
        Next

        'Copy client provided user data to user location
        If My.Computer.FileSystem.DirectoryExists(SCECore.SCEConfigPath() & "\Working\" & sRegion & "\" & sClient & "\User") = True Then
            dsClient = GetRowDB(SCECore.SCEConfigPath() & "\Working\" & sRegion & "\" & sClient & "\Database\SCE Client - " & sClient & ".accdb", "LUSettings", "Client", "Client", sClient, "", "", False, True)
            ''aFolders = GetFolders(SCECore.SCEConfigPath() & "\Working\" & sRegion & "\" & sClient & "\User")
            For Each sFolder In My.Computer.FileSystem.GetDirectories(SCECore.SCEConfigPath() & "\Working\" & sRegion & "\" & sClient & "\User")
                Try
                    My.Computer.FileSystem.CopyDirectory(SCECore.SCEConfigPath() & "\Working\" & sRegion & "\" & sClient & "\User\" & sFolder, SCECore.UserPath() & "\" & dsClient.Tables(0).Rows(0).Item("Software") & "\" & sFolder, False)
                Catch
                End Try
            Next
        End If

        'Start MicroStation Session
        With dsVersion.Tables(0).Rows(0)
            If .Item("LoadType") = "Program" Then
                sDrawing = Space(1) & Chr(34) & sDrawing & Chr(34)
                Shell(sMSFullName & Space(1) & Chr(34) & "-wp" & sClient & Chr(34) & sSoftwareAddon & sDrawing, AppWinStyle.NormalFocus)
            End If
        End With
ExitSub:
        dsVersion = Nothing
        dsSoftware = Nothing
    End Sub

    Shared Sub SendCommand(ByVal sCommand As String)
        'Send specified command to MicroStation
        For Each oProcess In Process.GetProcessesByName("ustation")
            Exit For
        Next

        If oProcess.HasExited = True Then
            _myInstance = New MicroStationDGN.Application
            oMSTN = GetInstance()
        End If
        oMSTN.CommandState.SetDefaultCursor()
        oMSTN.CadInputQueue.SendCommand(sCommand)
        AppActivate(oMSTN.ProcessID)
    End Sub

    Shared Sub CloseMicroStation()
        'End the MicroStation session
        For Each oProcess In Process.GetProcessesByName("ustation")
            Exit For
        Next

        oMSTN.Quit()
    End Sub

End Class
