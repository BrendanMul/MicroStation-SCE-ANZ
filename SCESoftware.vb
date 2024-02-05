'Imports MicroStationDGN

Public Class SCESoftware

    Public Shared dsSoftware As DataSet
    Public Shared Loaded As Boolean
    Public Shared Count As Integer

    Public Shared iMicroStationIndex As Integer = -1
    Public Shared iViewIndex As Integer = -1
    Public Shared iMarkupIndex As Integer = -1

    Public Shared sMicroStationPath As String
    Public Shared sMicroStationWorkspacePath As String
    Public Shared sNavigatorPath As String
    Public Shared sNavigatorWorkspacePath As String
    Public Shared sOpenPlantModelerPath As String
    Public Shared sOpenPlantModelerWorkspacePath As String
    Public Shared sOpenPlantIsometricsManagerPath As String
    Public Shared sOpenPlantIsometricsManagerWorkspacePath As String
    Public Shared sOpenRoadsDesignerPath As String
    Public Shared sOpenRoadsDesignerWorkspacePath As String
    Public Shared sOpenPlantPIDPath As String
    Public Shared sOpenPlantPIDWorkspacePath As String
    Public Shared sViewPath As String
    Public Shared sViewWorkspacePath As String
    Public Shared sProStructuresPath As String
    Public Shared sProStructuresWorkspacePath As String
    

    Public Structure DrawingOptions
        Dim Debug As Boolean
        Dim DisableAutoSave As Boolean
        Dim DisplayHiddenModels As Boolean
        Dim DisableReferences As Boolean
        Dim OpenReadOnly As Boolean
    End Structure

    Public Shared Sub Initialise()
        'Load software list into memory

        If JEGCore.bDebug = True Then JEGCore.RecordSequence(System.Reflection.MethodBase.GetCurrentMethod.Module.Name, System.Reflection.MethodBase.GetCurrentMethod.Name())

        If Loaded = True Then
            GoTo ExitSub
        End If

        'dsSoftware = SCEAccess.GetAllDB(SCECore.SCECoreDB, "LUSoftware", "Software")
        GetRegisteredSoftware()

        If SCEVB.IsNullDataSet(dsSoftware) = True Then
            Loaded = False
            Count = 0
            GoTo ExitSub
        End If

        Loaded = True
        Count = dsSoftware.Tables(0).Rows.Count
ExitSub:
    End Sub

    Shared Function GetRegisteredSoftware() As Boolean
        'load registered software from csv file
        Dim bReturn As Boolean = False
        Dim sPath As String = ""
        Dim sLine As String
        Dim aSplit As Array
        Dim oRow As DataRow
        Dim i As Integer
        Dim bFirstLine As Boolean = True

        If My.Computer.FileSystem.FileExists(My.Application.Info.DirectoryPath & "\Software.csv") = True Then
            sPath = My.Application.Info.DirectoryPath & "\Software.csv"
        ElseIf My.Computer.FileSystem.FileExists(JEGCore.LocalJMEPath() & "\software.csv") = True Then
            sPath = JEGCore.LocalJMEPath() & "\software.csv"
        End If

        If sPath <> "" Then
            FileOpen(1, sPath, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
            Do While Not EOF(1)
                If bFirstLine = True Then 'Skip column headers
                    bFirstLine = False
                    sLine = Trim(LineInput(1))
                    dsSoftware = SCEDataSet.CreateDataset("dsSoftware", Replace(sLine, ",", "|"), , "RGSoftware")
                    GoTo NextLine
                End If

                sLine = Trim(LineInput(1))

                'check whether to lines
                If sLine = "" Then
                    GoTo NextLine
                End If
                If SCEVB.IsCommented(sLine) = True Then
                    GoTo NextLine
                End If

                Dim sReturn As String

                'Get software 
                aSplit = Split(sLine, ",")
                'If UBound(aSplit) = 8 Then 'correct number of columns
                'only add software to table if it exists on the PC
                If UCase(aSplit(6)) = "PROGRAM" Then
                    Try
                        sReturn = My.Computer.Registry.GetValue(aSplit(2), aSplit(3), "")
                    Catch
                        sReturn = ""
                    End Try

                    If My.Computer.FileSystem.FileExists(sReturn) = True Then
                        oRow = dsSoftware.Tables(0).Rows.Add
                        For i = 0 To UBound(aSplit)
                            oRow.Item(i) = aSplit(i)
                        Next
                        oRow.AcceptChanges()
                    ElseIf aSplit(3) = "" Then
                        If My.Computer.FileSystem.FileExists(aSplit(4)) = True And InStr(aSplit(0), "OpenRoad") > 0 Then
                            oRow = dsSoftware.Tables(0).Rows.Add
                            For i = 0 To UBound(aSplit)
                                oRow.Item(i) = aSplit(i)
                            Next
                            oRow.AcceptChanges()
                        End If
                    End If
                Else
                    oRow = dsSoftware.Tables(0).Rows.Add
                    For i = 0 To UBound(aSplit)
                        oRow.Item(i) = aSplit(i)
                    Next
                    oRow.AcceptChanges()
                End If
                'End If
NextLine:
            Loop
            FileClose(1)

            'check if software has been loaded
            If Not dsSoftware Is Nothing Then
                If dsSoftware.Tables.Count > 0 Then
                    If dsSoftware.Tables(0).Rows.Count > 0 Then
                        bReturn = True
                    End If
                End If
            End If
        End If

        Return bReturn
    End Function

    Shared Function GetIndex(ByVal sSoftware As String, Optional ByVal bLike As Boolean = False) As Integer
        'get index of specified software
        Dim iReturn As Integer = -1
        Dim i As Integer

        For i = 0 To Count - 1
            If bLike = True Then
                If InStr(UCase(dsSoftware.Tables(0).Rows(i).Item("Software")), UCase(sSoftware)) > 0 Then
                    iReturn = i
                    Exit For
                End If
            Else
                If UCase(sSoftware) = UCase(dsSoftware.Tables(0).Rows(i).Item("Software")) Then
                    iReturn = i
                    Exit For
                End If
            End If
        Next

        Return iReturn
    End Function

    Shared Function GetIndexByKeyword(ByVal sKeyword As String) As Integer
        'get index of specified keyword
        Dim iReturn As Integer = -1
        Dim i As Integer

        For i = 0 To Count - 1
            If Not dsSoftware.Tables(0).Rows(i).Item("Keyword") Is System.DBNull.Value Then
                If InStr(UCase(dsSoftware.Tables(0).Rows(i).Item("Keyword")), UCase(sKeyword)) > 0 Then
                    iReturn = i
                    Exit For
                End If
            End If
        Next

        Select Case UCase(sKeyword)
            Case "MICROSTATION"
                iMicroStationIndex = iReturn
            Case "VIEW"
                iViewIndex = iReturn
            Case "MARKUP"
                iMarkupIndex = iReturn
        End Select

        Return iReturn
    End Function

    Shared Function GetIndexes(ByVal sPrerequisite As String) As Array
        'get index of specified software
        Dim i As Integer
        Dim j As Integer = -1
        Dim aReturn() As String = Nothing

        For i = 0 To Count - 1
            If dsSoftware.Tables(0).Rows(i).Item("Prerequisite") Is System.DBNull.Value Then
                GoTo NextSoftware
            End If
            If InStr(UCase(dsSoftware.Tables(0).Rows(i).Item("Prerequisite")), UCase(sPrerequisite)) > 0 Then
                j = j + 1
                ReDim Preserve aReturn(j)
                aReturn(j) = i
            End If
NextSoftware:
        Next

        If j = -1 Then
            ReDim aReturn(0)
        End If

        Return aReturn
    End Function

    Public Shared Sub LoadUsingDefaultProgram(ByVal sFileName As String, Optional ByVal oWindowStyle As ProcessWindowStyle = ProcessWindowStyle.Normal)
        'load specified file using default program (file association)
        Dim p As New System.Diagnostics.Process
        Dim s As New System.Diagnostics.ProcessStartInfo(sFileName)

        If My.Computer.FileSystem.FileExists(sFileName) = False Then
            GoTo ExitSub
        End If

        s.UseShellExecute = True
        s.WindowStyle = oWindowStyle
        p.StartInfo = s
        p.Start()
ExitSub:
    End Sub

    Shared Function IsInstalled(ByVal iSoftwareIndex As String) As Boolean
        'Return boolean is specified software is installed
        Dim bReturn As Boolean = False
        Dim sPath As String

        If iSoftwareIndex > -1 Then
            With dsSoftware.Tables(0).Rows(iSoftwareIndex)
                If .Item("RegistryKey") Is System.DBNull.Value Then GoTo CheckLocation
                sPath = My.Computer.Registry.GetValue(.Item("RegistryKey"), .Item("RegistryString"), Nothing)
                If sPath Is Nothing Then
CheckLocation:
                    bReturn = My.Computer.FileSystem.FileExists(.Item("Location"))
                Else
                    If My.Computer.FileSystem.FileExists(My.Computer.Registry.GetValue(.Item("RegistryKey"), .Item("RegistryString"), Nothing)) = True Then
                        bReturn = True
                    ElseIf My.Computer.FileSystem.DirectoryExists(My.Computer.Registry.GetValue(.Item("RegistryKey"), .Item("RegistryString"), Nothing)) = True Then
                        bReturn = True
                    End If
                End If
            End With
        End If
ExitFunction:
        Return bReturn
    End Function

    Shared Function GetFileName(ByVal sSoftware As String) As String
        'Return full filename to specified software
        Dim sReturn As String = ""
        Dim iIndex As Integer

        iIndex = GetIndex(sSoftware)
        If iIndex = -1 Then
            GoTo ExitFunction
        End If

        'Check if path determined by registry
        If Not dsSoftware.Tables(0).Rows(iIndex).Item("RegistryKey") Is System.DBNull.Value And Not dsSoftware.Tables(0).Rows(iIndex).Item("RegistryString") Is System.DBNull.Value Then
            sReturn = My.Computer.Registry.GetValue(dsSoftware.Tables(0).Rows(iIndex).Item("RegistryKey"), dsSoftware.Tables(0).Rows(iIndex).Item("RegistryString"), "")
        End If
        If sReturn = "" Then
            'no path in registry, get from location column instead
            If Not dsSoftware.Tables(0).Rows(iIndex).Item("Location") Is System.DBNull.Value Then
                sReturn = dsSoftware.Tables(0).Rows(iIndex).Item("Location")
            End If
        End If
ExitFunction:
        Return sReturn
    End Function

    Shared Function GetSoftwarePath(ByVal sSoftware As String) As String
        'Return path to installed software
        Dim sReturn As String = ""
        Dim iIndex As Integer = SCESoftware.GetIndex(sSoftware)

        If iIndex = -1 Then
            GoTo ExitFunction
        End If

        sMicroStationPath = ""
        sMicroStationWorkspacePath = ""
        sNavigatorPath = ""
        sNavigatorWorkspacePath = ""
        sOpenPlantModelerPath = ""
        sOpenPlantModelerWorkspacePath = ""
        sOpenPlantIsometricsManagerPath = ""
        sOpenPlantIsometricsManagerWorkspacePath = ""
        sOpenRoadsDesignerPath = ""
        sOpenRoadsDesignerWorkspacePath = ""
        sOpenPlantPIDPath = ""
        sOpenPlantPIDWorkspacePath = ""
        sViewPath = ""
        sViewWorkspacePath = ""
        sProStructuresPath = ""
        sProStructuresWorkspacePath = ""

        If UCase(SCESoftware.dsSoftware.Tables(0).Rows(iIndex).Item("LoadType")) = "PROGRAM" Then
            'Get path from registry 
            sReturn = My.Computer.Registry.GetValue(SCESoftware.dsSoftware.Tables(0).Rows(iIndex).Item("RegistryKey"), SCESoftware.dsSoftware.Tables(0).Rows(iIndex).Item("RegistryString"), "")
            If InStr(UCase(sSoftware), "MICROSTATION") > 0 Then
                sMicroStationPath = sReturn
                sMicroStationWorkspacePath = My.Computer.Registry.GetValue(SCESoftware.dsSoftware.Tables(0).Rows(iIndex).Item("RegistryKey"), "WorkspacePath", "")
            ElseIf InStr(UCase(sSoftware), "NAVIGATOR") > 0 Then
                sNavigatorPath = sReturn
                sNavigatorWorkspacePath = My.Computer.Registry.GetValue(SCESoftware.dsSoftware.Tables(0).Rows(iIndex).Item("RegistryKey"), "WorkspacePath", "")
            ElseIf InStr(UCase(sSoftware), "OPENPLANT MODELER") > 0 Then
                sOpenPlantModelerPath = sReturn
                sOpenPlantModelerWorkspacePath = My.Computer.Registry.GetValue(SCESoftware.dsSoftware.Tables(0).Rows(iIndex).Item("RegistryKey"), "WorkspacePath", "")
            ElseIf InStr(UCase(sSoftware), "OPENPLANT ISOMETRICS MANAGER") > 0 Then
                sReturn = SCESoftware.dsSoftware.Tables(0).Rows(iIndex).Item("Location")
                sOpenPlantIsometricsManagerPath = sReturn
                sOpenPlantIsometricsManagerWorkspacePath = My.Computer.Registry.GetValue(SCESoftware.dsSoftware.Tables(0).Rows(iIndex).Item("RegistryKey"), "WorkspacePath", "")
            ElseIf InStr(UCase(sSoftware), "OPENROADS DESIGNER CONNECT EDITION") > 0 Then
                sReturn = SCESoftware.dsSoftware.Tables(0).Rows(iIndex).Item("Location")
                sOpenRoadsDesignerPath = sReturn
                sOpenRoadsDesignerWorkspacePath = My.Computer.Registry.GetValue(SCESoftware.dsSoftware.Tables(0).Rows(iIndex).Item("RegistryKey"), "WorkspacePath", "")
            ElseIf InStr(UCase(sSoftware), "OPENPLANT PID") > 0 Then
                sOpenPlantPIDPath = sReturn
                sOpenPlantPIDWorkspacePath = My.Computer.Registry.GetValue(SCESoftware.dsSoftware.Tables(0).Rows(iIndex).Item("RegistryKey"), "WorkspacePath", "")
            ElseIf InStr(UCase(sSoftware), "PROSTRUCTURES") > 0 Then
                sProStructuresPath = sReturn
                sProStructuresWorkspacePath = My.Computer.Registry.GetValue(SCESoftware.dsSoftware.Tables(0).Rows(iIndex).Item("RegistryKey"), "WorkspacePath", "")
            ElseIf InStr(UCase(sSoftware), "VIEW") > 0 Then
                sViewPath = sReturn
                sViewWorkspacePath = My.Computer.Registry.GetValue(SCESoftware.dsSoftware.Tables(0).Rows(iIndex).Item("RegistryKey"), "WorkspacePath", "")
            End If
        End If
ExitFunction:
        Return sReturn
    End Function

    Shared Sub LoadSoftware(ByVal sRegion As String, ByVal sClient As String, _
        Optional ByVal sSoftware As String = "", Optional ByVal sSoftwareVersion As String = "", _
        Optional ByVal sSoftwareAddon As String = "", Optional ByVal sTools As String = "", _
        Optional ByVal iClientIndex As Integer = -1, Optional ByVal sDrawing As String = "", _
        Optional ByVal sProjectName As String = "", Optional ByVal sProjectFileName As String = "", _
        Optional ByVal sAutoCADRegion As String = "", Optional ByVal sAutoCADClient As String = "", _
        Optional ByVal bReadOnly As Boolean = False, Optional ByVal bDisableReferences As Boolean = False, _
        Optional ByVal bDebug As Boolean = False, Optional ByVal sConfigVars As String = "", _
        Optional ByVal bLoadMinimal As Boolean = False, Optional ByVal sModelName As String = "", _
        Optional ByVal bWait As Boolean = False, Optional ByVal bProjectWise As Boolean = False, _
        Optional ByVal sSubstitution As String = "", Optional ByVal bIncludeHiddenModels As Boolean = False, _
        Optional ByVal bDisableAutoSave As Boolean = False, Optional ByVal WindowsStyle As AppWinStyle = AppWinStyle.NormalFocus, _
        Optional ByVal sDepartment As String = "", Optional ByVal sSelectedClient As String = "", Optional bIgnoreJMEConfig As Boolean = False, Optional sClientShortName As String = "")
        'Execute MicroStation with specified details
        Dim dsSoftware As DataRow = Nothing
        Dim sMSFullName As String = Nothing
        Dim sMSPath As String = Nothing
        Dim sLine As String = Nothing
        Dim sConfigFullName As String = Nothing
        Dim bProjectLoaded As Boolean = False
        Dim i As Integer
        Dim sConfigName As String = ""
        Dim sMSCFGFullName As String = ""
        Dim sReadOnly As String
        Dim sDisableReferences As String = ""
        Dim sDebug As String = ""
        Dim frmSplash As New frmSplash
        Dim sModel As String = ""
        Dim aSplit As Array
        Dim sJMELauncher As String = ""

        If sRegion = "" And sClient = "" And InStr(UCase(sConfigVars), "AUTOSCE") > 0 Then
            sMSCFGFullName = Space(1) & Chr(34) & "-wc" & sConfigVars & Chr(34)
            sMSFullName = Chr(34) & GetSoftwarePath(sSoftware) & Chr(34)
            If sMSFullName = Nothing Then GoTo ExitSub
            GoTo LoadSoftware
        End If

        If iClientIndex = -1 Then
            iClientIndex = SCEClients.GetIndex(, sClient)
        End If

        'If iClientIndex = -1 Then
        '    If sDrawing = "" Then 'Client selector used

        '    End If 'sdrawing
        'Else 'Favourite used
        '    If JEGCore.bDebug = True Then JEGCore.RecordSequence(System.Reflection.MethodBase.GetCurrentMethod.Module.Name, System.Reflection.MethodBase.GetCurrentMethod.Name(), "Favourite used, loading")
        '    If SCEFavourite.ClientsLoaded = False Then
        '        SCEFavourite.LoadClients()
        '    End If
        '    iClientIndex = SCEClients.GetIndex(, SCEFavourite.dsFavouriteClients.Tables(0).Rows(iFavouriteClientIndex).Item("Client"))
        '    If JEGCore.bDebug = True Then JEGCore.RecordSequence(System.Reflection.MethodBase.GetCurrentMethod.Module.Name, System.Reflection.MethodBase.GetCurrentMethod.Name(), "Favourite loaded")
        'End If 'favourite

        'Set last used software
        'My.Computer.Registry.SetValue(JEGCore.MSTNSCEKey, "NewLastMSTNAddonSoftware", sSoftwareAddon)

        If JEGCore.bDebug = True Then JEGCore.RecordSequence(System.Reflection.MethodBase.GetCurrentMethod.Module.Name, System.Reflection.MethodBase.GetCurrentMethod.Name(), "Get software")
        sMSFullName = GetSoftwarePath(sSoftware)
        If sMSFullName = Nothing Then GoTo ExitSub

        sMSPath = SCEFile.FileParse(sMSFullName, SCEFile.FileParser.Path)

        If sClient = "Standard MicroStation" Then
            GoTo LoadSoftware
        End If

        sMSCFGFullName = JEGCore.UserTempPath & "\" & sSoftware & "\Config\AutoJME.cfg"

        If JEGCore.bDebug = True Then JEGCore.RecordSequence(System.Reflection.MethodBase.GetCurrentMethod.Module.Name, System.Reflection.MethodBase.GetCurrentMethod.Name(), "Reset AutoSCE.cfg")

        If My.Computer.FileSystem.DirectoryExists(SCEFile.FileParse(sMSCFGFullName, SCEFile.FileParser.Path)) = False Then
            My.Computer.FileSystem.CreateDirectory(SCEFile.FileParse(sMSCFGFullName, SCEFile.FileParser.Path))
        End If

        'Create local temp path
        If JEGCore.bDebug = True Then JEGCore.RecordSequence(System.Reflection.MethodBase.GetCurrentMethod.Module.Name, System.Reflection.MethodBase.GetCurrentMethod.Name(), "Create local temp path")
        If My.Computer.FileSystem.DirectoryExists(Replace(JEGCore.UserTempPath & "/" & sSoftware & "\Temp", "/", "\")) = False Then
            My.Computer.FileSystem.CreateDirectory(Replace(JEGCore.UserTempPath & "/" & sSoftware & "\Temp", "/", "\"))
        End If

        FileOpen(2, sMSCFGFullName, OpenMode.Output, OpenAccess.Write, OpenShare.LockWrite)
        PrintLine(2, "# Load session based JME settings")
        PrintLine(2, "# Last Loaded: " & Now)
        PrintLine(2, "")

        'Software
        PrintLine(2, "JME_SOFTWARE_ROOT = " & JEGCore.LocalJMEPath & "\Software/")
        If My.Computer.FileSystem.DirectoryExists(JEGCore.LocalJMEPath & "\Software\" & sSoftware) = True Then
            PrintLine(2, "JME_SOFTWARE_PATH = " & JEGCore.LocalJMEPath & "\Software\" & sSoftware)
            If My.Computer.FileSystem.FileExists(JEGCore.LocalJMEPath & "\Software\" & sSoftware & "\" & sSoftware & ".cfg") = True Then
                PrintLine(2, "JME_SOFTWARE_CFG = " & JEGCore.LocalJMEPath & "\Software\" & sSoftware & "\" & sSoftware & ".cfg")
                PrintLine(2, "%if exists ($(JME_SOFTWARE_CFG))")
                PrintLine(2, "  %include $(JME_SOFTWARE_CFG)")
                PrintLine(2, "%endif")
            End If
        End If

        'Client specific settings
        If InStr(SCEClients.dsClients.Tables(1).Rows(iClientIndex).Item("ClientPath"), "C:\Users\Public", CompareMethod.Binary) > 0 Then
            PrintLine(2, "_USTN_PROJECTSROOT = " & JEGCore.LocalPublicPath & "\Clients/")
        ElseIf Microsoft.VisualBasic.Left(SCEClients.dsClients.Tables(1).Rows(iClientIndex).Item("ClientPath"), 3) <> "C:\" Then 'network
            PrintLine(2, "_USTN_PROJECTSROOT = " & Replace(Microsoft.VisualBasic.Left(SCEClients.dsClients.Tables(1).Rows(iClientIndex).Item("ClientPath"), InStr(UCase(SCEClients.dsClients.Tables(1).Rows(iClientIndex).Item("ClientPath")), "CLIENTS") + 7), "\", "/"))
        End If

        'ProjectWise - pjrxxx: call cfg file instead
        If bProjectWise = True Then
            PrintLine(2, "# Enable ProjectWise integration")
            PrintLine(2, "%undef PW_DISABLE_INTEGRATION")
        Else
            PrintLine(2, "# Disable ProjectWise integration")
            PrintLine(2, "PW_DISABLE_INTEGRATION=1")
        End If

        'Tools
        If sTools <> "" Then
            aSplit = Split(sTools, ";")
            For i = 0 To UBound(aSplit)
                sTools = SCESoftware.dsSoftware.Tables(0).Rows(SCESoftware.GetIndex(aSplit(i))).Item("Location")
                PrintLine(2, "# Load tool " & aSplit(i))
                PrintLine(2, "JME_TOOL_" & i + 1 & " = " & sTools)
                PrintLine(2, "%if exists ($(JME_TOOL_" & i + 1 & "))")
                PrintLine(2, "  %include $(JME_TOOL_" & i + 1 & ")")
                PrintLine(2, "%endif")
            Next
        End If

        ''Project
        If My.Computer.FileSystem.FileExists(sProjectFileName) = True Then
            PrintLine(2, "# Load project settings")
            PrintLine(2, "")
            PrintLine(2, "JME_PROJECT_CFG = " & sProjectFileName)
            PrintLine(2, "%if exists ($(JME_PROJECT_CFG))")
            PrintLine(2, "  %include $(JME_PROJECT_CFG)")
            PrintLine(2, "%endif")
            PrintLine(2, "")
        End If

        If InStr(UCase(sSoftware), "PROSTRUCTURES V8I SELECTSERIES 8") > 0 Then
            PrintLine(2, "JME_PROSTRUCTURES_CFG = $(JME_PROJECTPATH)Software\ProStructures\ProStructures.cfg")
            PrintLine(2, "%if exists ($(JME_PROSTRUCTURES_CFG))")
            PrintLine(2, "  %include $(JME_PROSTRUCTURES_CFG)")
            PrintLine(2, "%endif")
        End If

        FileClose(2)

        If bIncludeHiddenModels = True Then
            PrintLine(2, vbTab & "MS_INCLUDEHIDDENMODELS = 1")
        End If
        If bDisableAutoSave = True Then
            PrintLine(2, vbTab & "MS_DGNAUTOSAVE = 0")
        End If

        sMSFullName = Chr(34) & sMSFullName & Chr(34)

        If JEGCore.bDebug = True Then JEGCore.RecordSequence(System.Reflection.MethodBase.GetCurrentMethod.Module.Name, System.Reflection.MethodBase.GetCurrentMethod.Name(), "Set miscellaneous configuration variables")

        If bDisableReferences = True Then
            sDisableReferences = Space(1) & "-O"
        Else
            sDisableReferences = ""
        End If

        If bDebug = True Then
            sDebug = Space(1) & "-DEBUG"
        Else
            sDebug = ""
        End If

        If My.Computer.FileSystem.DirectoryExists(JEGCore.UserPath() & "\" & sSoftware & "\Prefs") = False Then
            If JEGCore.bDebug = True Then JEGCore.RecordSequence(System.Reflection.MethodBase.GetCurrentMethod.Module.Name, System.Reflection.MethodBase.GetCurrentMethod.Name(), "Create user Preferences location")
            My.Computer.FileSystem.CreateDirectory(JEGCore.UserPath() & "\" & sSoftware & "\Prefs")
        End If
LoadSoftware:

        'Readonly
        If bReadOnly = True Then
            sReadOnly = Space(1) & "-R"
        Else
            sReadOnly = ""
        End If

        'If WindowsStyle = AppWinStyle.Hide Then
        '    'Dim oProcess As New Process
        '    Dim StartInfo As New ProcessStartInfo
        '    StartInfo.FileName = sMSFullName
        '    StartInfo.Arguments = sMSCFGFullName & sDrawing & sReadOnly & sDisableReferences & sDebug & sModel & sProjectWise
        '    StartInfo.UseShellExecute = False
        '    StartInfo.RedirectStandardOutput = True
        '    StartInfo.UseShellExecute = True
        '    StartInfo.RedirectStandardOutput = False
        '    StartInfo.WindowStyle = ProcessWindowStyle.Hidden

        '    Process.Start(StartInfo)
        'Else

        'Download JME Standards.cfg
        If bIgnoreJMEConfig = False Then
            If SCESoftware.sMicroStationWorkspacePath <> "" Then
                My.Computer.FileSystem.CopyFile(My.Application.Info.DirectoryPath & "\Workspace\JME Standards.cfg", SCESoftware.sMicroStationWorkspacePath & "\Standards\JME Standards.cfg", True)
                FileOpen(2, SCESoftware.sMicroStationWorkspacePath & "\Standards\JME Standards.cfg", OpenMode.Append, OpenAccess.Write, OpenShare.LockWrite)
                PrintLine(2, vbTab & "SCE_SOFTWARE = " & sSoftware)
                PrintLine(2, "%if exists ($(SCE_LOCAL_TEMP)Config/AutoJME.cfg)")
                PrintLine(2, "  %include $(SCE_LOCAL_TEMP)Config/AutoJME.cfg")
                PrintLine(2, "%endif")
                FileClose(2)
            End If
            If SCESoftware.sNavigatorWorkspacePath <> "" Then
                My.Computer.FileSystem.CopyFile(My.Application.Info.DirectoryPath & "\Workspace\JME Standards.cfg", SCESoftware.sNavigatorWorkspacePath & "\Standards\JME Standards.cfg", True)
                FileOpen(2, SCESoftware.sNavigatorWorkspacePath & "\Standards\JME Standards.cfg", OpenMode.Append, OpenAccess.Write, OpenShare.LockWrite)
                PrintLine(2, vbTab & "SCE_SOFTWARE = " & sSoftware)
                PrintLine(2, "%if exists ($(SCE_LOCAL_TEMP)Config/AutoJME.cfg)")
                PrintLine(2, "  %include $(SCE_LOCAL_TEMP)Config/AutoJME.cfg")
                PrintLine(2, "%endif")
                FileClose(2)
            End If
            If SCESoftware.sOpenPlantModelerWorkspacePath <> "" Then
                My.Computer.FileSystem.CopyFile(My.Application.Info.DirectoryPath & "\Workspace\JME Standards.cfg", SCESoftware.sOpenPlantModelerWorkspacePath & "\Standards\JME Standards.cfg", True)
                FileOpen(2, SCESoftware.sOpenPlantModelerWorkspacePath & "\Standards\JME Standards.cfg", OpenMode.Append, OpenAccess.Write, OpenShare.LockWrite)
                PrintLine(2, vbTab & "SCE_SOFTWARE = " & sSoftware)
                PrintLine(2, "%if exists ($(SCE_LOCAL_TEMP)Config/AutoJME.cfg)")
                PrintLine(2, "  %include $(SCE_LOCAL_TEMP)Config/AutoJME.cfg")
                PrintLine(2, "%endif")
                FileClose(2)
            End If
            If SCESoftware.sOpenPlantIsometricsManagerWorkspacePath <> "" Then
                My.Computer.FileSystem.CopyFile(My.Application.Info.DirectoryPath & "\Workspace\JME Standards.cfg", SCESoftware.sOpenPlantIsometricsManagerWorkspacePath & "\Standards\JME Standards.cfg", True)
                FileOpen(2, SCESoftware.sOpenPlantIsometricsManagerWorkspacePath & "\Standards\JME Standards.cfg", OpenMode.Append, OpenAccess.Write, OpenShare.LockWrite)
                PrintLine(2, vbTab & "SCE_SOFTWARE = " & sSoftware)
                PrintLine(2, "%if exists ($(SCE_LOCAL_TEMP)Config/AutoJME.cfg)")
                PrintLine(2, "  %include $(SCE_LOCAL_TEMP)Config/AutoJME.cfg")
                PrintLine(2, "%endif")
                FileClose(2)
            End If
            If SCESoftware.sOpenRoadsDesignerWorkspacePath <> "" Then
                My.Computer.FileSystem.CopyFile(My.Application.Info.DirectoryPath & "\Workspace\JME Standards.cfg", SCESoftware.sOpenRoadsDesignerWorkspacePath & "\Standards\JME Standards.cfg", True)
                FileOpen(2, SCESoftware.sOpenRoadsDesignerWorkspacePath & "\Standards\JME Standards.cfg", OpenMode.Append, OpenAccess.Write, OpenShare.LockWrite)
                PrintLine(2, vbTab & "SCE_SOFTWARE = " & sSoftware)
                PrintLine(2, "%if exists ($(SCE_LOCAL_TEMP)Config/AutoJME.cfg)")
                PrintLine(2, "  %include $(SCE_LOCAL_TEMP)Config/AutoJME.cfg")
                PrintLine(2, "%endif")
                FileClose(2)
            End If
            If SCESoftware.sOpenPlantPIDWorkspacePath <> "" Then
                My.Computer.FileSystem.CopyFile(My.Application.Info.DirectoryPath & "\Workspace\JME Standards.cfg", SCESoftware.sOpenPlantPIDWorkspacePath & "\Standards\JME Standards.cfg", True)
                FileOpen(2, SCESoftware.sOpenPlantPIDWorkspacePath & "\Standards\JME Standards.cfg", OpenMode.Append, OpenAccess.Write, OpenShare.LockWrite)
                PrintLine(2, vbTab & "SCE_SOFTWARE = " & sSoftware)
                PrintLine(2, "%if exists ($(SCE_LOCAL_TEMP)Config/AutoJME.cfg)")
                PrintLine(2, "  %include $(SCE_LOCAL_TEMP)Config/AutoJME.cfg")
                PrintLine(2, "%endif")
                FileClose(2)
            End If
            If SCESoftware.sProStructuresWorkspacePath <> "" Then
                My.Computer.FileSystem.CopyFile(My.Application.Info.DirectoryPath & "\Workspace\JME Standards.cfg", SCESoftware.sProStructuresWorkspacePath & "\Standards\JME Standards.cfg", True)
                FileOpen(2, SCESoftware.sProStructuresWorkspacePath & "\Standards\JME Standards.cfg", OpenMode.Append, OpenAccess.Write, OpenShare.LockWrite)
                PrintLine(2, vbTab & "SCE_SOFTWARE = " & sSoftware)
                PrintLine(2, "%if exists ($(SCE_LOCAL_TEMP)Config/AutoJME.cfg)")
                PrintLine(2, "  %include $(SCE_LOCAL_TEMP)Config/AutoJME.cfg")
                PrintLine(2, "%endif")
                FileClose(2)
            End If
            If SCESoftware.sViewWorkspacePath <> "" Then
                My.Computer.FileSystem.CopyFile(My.Application.Info.DirectoryPath & "\Workspace\JME Standards.cfg", SCESoftware.sViewWorkspacePath & "\Standards\JME Standards.cfg", True)
                FileOpen(2, SCESoftware.sViewWorkspacePath & "\Standards\JME Standards.cfg", OpenMode.Append, OpenAccess.Write, OpenShare.LockWrite)
                PrintLine(2, vbTab & "SCE_SOFTWARE = " & sSoftware)
                PrintLine(2, "%if exists ($(SCE_LOCAL_TEMP)Config/AutoJME.cfg)")
                PrintLine(2, "  %include $(SCE_LOCAL_TEMP)Config/AutoJME.cfg")
                PrintLine(2, "%endif")
                FileClose(2)
            End If
        End If

        Dim sProject As String = ""
        Dim sUser As String = ""

        'configure software command line switches otherwise launch default software
        If sClient <> "" Then
            'sProject = Space(1) & Chr(34) & "-wp" & sClient & Chr(34)
            sUser = " -wu" & Environment.UserName
            sJMELauncher = " -wsJME_LAUNCHER=1"
        End If

        If InStr(UCase(sSoftware), "ISOMETRICS MANAGER") > 0 Then
            'Dim sTempPath As String = SCEFile.FileParse(sMSFullName, SCEFile.FileParser.Path)
            'If Trim(sTempPath) <> "" Then
            '    sTempPath = Replace(sTempPath, "C:\", "")
            '    sTempPath = Replace(sTempPath, Chr(34), "")
            'End If
            'Dim sTempFilename As String = SCEFile.FileParse(sMSFullName, SCEFile.FileParser.FullFileName)
            'sTempFilename = Replace(sTempFilename, Chr(34), "")

            'MsgBox("cd\" & sTempPath & " & " & sTempFilename & sProject & sUser & sJMELauncher)
            'Shell("cd\" & sTempPath & " & " & sTempFilename & sProject & sUser & sJMELauncher, WindowsStyle, False)
            'MsgBox(sMSFullName)

            'Reset OpenPlant Modeller
            If My.Computer.FileSystem.FileExists(SCESoftware.sOpenPlantModelerWorkspacePath & "\Standards\JME Standards.cfg") = True Then
                My.Computer.FileSystem.DeleteFile(SCESoftware.sOpenPlantModelerWorkspacePath & "\Standards\JME Standards.cfg")
            End If

            LoadUsingDefaultProgram(Replace(sMSFullName, Chr(34), ""))
            'Shell(sMSFullName, WindowsStyle)
        ElseIf InStr(UCase(sSoftware), "OPENROADS") > 0 Then
            If bIgnoreJMEConfig = False Then
                'Reset OpenRoads
                ' PJR laptop = C:\ProgramData\Bentley\OpenRoads Designer CE\Configuration\
                If My.Computer.FileSystem.FileExists("C:\ProgramData\Bentley\OpenRoads Designer CE\Configuration\WorkSpaceSetup.orig") = False Then
                    My.Computer.FileSystem.RenameFile("C:\ProgramData\Bentley\OpenRoads Designer CE\Configuration\WorkSpaceSetup.cfg", "WorkSpaceSetup.orig")
                    'My.Computer.FileSystem.DeleteFile("C:\ProgramData\Bentley\OpenRoads Designer CE\Configuration\WorkSpaceSetup.cfg")
                    My.Computer.FileSystem.CopyFile("Z:\Jacobs\Bentley\MicroStation\Software\OpenRoads Designer CONNECT\Configuration\WorkSpaceSetup.cfg", "C:\ProgramData\Bentley\OpenRoads Designer CE\Configuration\WorkSpaceSetup.cfg")
                Else
                    My.Computer.FileSystem.CopyFile("Z:\Jacobs\Bentley\MicroStation\Software\OpenRoads Designer CONNECT\Configuration\WorkSpaceSetup.cfg", "C:\ProgramData\Bentley\OpenRoads Designer CE\Configuration\WorkSpaceSetup.cfg", True)
                End If

                Shell("Z:\Jacobs\Bentley\MicroStation\Software\OpenRoads Designer CONNECT\Prestart.bat", , True)
            End If

            LoadUsingDefaultProgram(Replace(sMSFullName, Chr(34), ""))
            'ElseIf InStr(UCase(sSoftware), "PROSTRUCTURES V8I SELECTSERIES 8") > 0 Then
            '    If sProjectName <> "" Then
            '        FileOpen(1, JEGCore.NetworkPath & "\Projects\" & sClientShortName & "\" & sProjectName & ".cfg", OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
            '        Do While Not EOF(1)
            '            sLine = Trim(LineInput(1))
            '            If InStr(sLine, "SCE_PROJECTDRAWINGPATH") > 0 Then
            '                sLine = Replace(UCase(sLine), "SCE_PROJECTDRAWINGPATH", "")
            '                sLine = Replace(sLine, "=", "")
            '                sLine = Trim(Replace(sLine, "/", ""))
            '                sLine = Trim(sLine)
            '                sLine = Replace(sLine, vbTab, "")
            '                Exit Do
            '            End If
            '        Loop
            '        FileClose(1)
            '    End If

            '    UpdateProStructuresS8(sLine)
        Else
            Shell(sMSFullName & sProject & sUser & sJMELauncher, WindowsStyle, False) 'sReadOnly & sDisableReferences & sDebug & sModel &
        End If

ExitSub:
        dsSoftware = Nothing
        If JEGCore.bDebug = True Then JEGCore.RecordSequence(System.Reflection.MethodBase.GetCurrentMethod.Module.Name, System.Reflection.MethodBase.GetCurrentMethod.Name(), "Load software: " & sSoftware & " complete")
    End Sub

    Private Shared Sub UpdateProStructuresS8(ByVal sProjectDrawingPath As String)
        If My.Computer.FileSystem.FileExists(sProjectDrawingPath & "\Admin\SCE\Software\ProStructures\product.cfg") = True Then
            'Backup original file
            If My.Computer.FileSystem.FileExists("C:\Program Files (x86)\Bentley\ProStructures V8i S8\ProStructures\config\appl\product.orig") = False Then
                Rename("C:\Program Files (x86)\Bentley\ProStructures V8i S8\ProStructures\config\appl\product.cfg", "C:\Program Files (x86)\Bentley\ProStructures V8i S8\ProStructures\config\appl\product.orig")
            End If
            'Copy Project CFG file to local
            FileCopy(sProjectDrawingPath & "\Admin\SCE\Software\ProStructures\product.cfg", "C:\Program Files (x86)\Bentley\ProStructures V8i S8\ProStructures\config\appl\product.cfg")
        Else
            'Restore original CFG
            If My.Computer.FileSystem.FileExists("C:\Program Files (x86)\Bentley\ProStructures V8i S8\ProStructures\config\appl\product.orig") = True Then
                My.Computer.FileSystem.DeleteFile("C:\Program Files (x86)\Bentley\ProStructures V8i S8\ProStructures\config\appl\product.cfg", FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
                Rename("C:\Program Files (x86)\Bentley\ProStructures V8i S8\ProStructures\config\appl\product.orig", "C:\Program Files (x86)\Bentley\ProStructures V8i S8\ProStructures\config\appl\product.cfg")
            End If
        End If
    End Sub

    Public Shared Function GetMicroStationPath(ByVal sSoftware As String, ByVal sVersion As String) As String
        'Return path to installed software
        Dim dsVersion As DataSet
        Dim dsCoreSoftware As DataSet
        Dim i As Long
        Dim j As Long
        Dim sValue As String = Nothing
        Dim aSoftware As Array
        Dim sMessage As String = ""

        aSoftware = Split(sSoftware, "|")

        For j = 0 To UBound(aSoftware)
            dsVersion = SCEAccess.GetRowsDB(JEGCore.SCECoreDB, "LUSoftware", "RegistryKey", "Software", aSoftware(j), "", "", False, True)
            If dsVersion Is Nothing Then GoTo SkipVersion
            If dsVersion.Tables(0).Rows.Count = 0 Then
            Else
                If UCase(dsVersion.Tables(0).Rows(i).Item("LoadType")) = "PROGRAM" Then
                    'Get path from registry
                    For i = 0 To (dsVersion.Tables(0).Rows.Count - 1)
                        sValue = My.Computer.Registry.GetValue(dsVersion.Tables(0).Rows(i).Item("RegistryKey"), dsVersion.Tables(0).Rows(i).Item("RegistryString"), Nothing)
                        sSoftware = aSoftware(j)
                        If sValue <> Nothing Then
                            Exit For
                        End If
                    Next
                ElseIf UCase(dsVersion.Tables(0).Rows(i).Item("LoadType")) = "ADDON" Then
                    'Get core software
                    dsCoreSoftware = SCEAccess.GetRowsDB(JEGCore.SCECoreDB, "LUSoftware", "RegistryKey", "Software", aSoftware(j), "", "", False, True)
                End If
            End If
            Exit For
        Next
SkipVersion:
        'Check file exists
        If My.Computer.FileSystem.FileExists(sValue) = False Then
            For i = 0 To UBound(aSoftware)
                If i = 0 Then
                    sMessage = aSoftware(i)
                Else
                    sMessage = Chr(10) & aSoftware(i)
                End If
            Next
            MsgBox("One of the following software is required for this client:" & Chr(10) & Chr(10) & _
                sMessage, MsgBoxStyle.Exclamation, JEGCore.MsgBoxCaption)
            sValue = Nothing
        End If
        Return sValue
        dsVersion = Nothing
    End Function

    Public Shared Function IsProjectWiseInstalled() As Boolean
        'return if projectwise is installed
        Dim bReturn As Boolean = False

        If My.Computer.FileSystem.FileExists(GetProjectWisePath() & "bin\pwc.exe") = True Then
            bReturn = True
        End If

        Return bReturn
    End Function

    Public Shared Function GetProjectWisePath() As String
        'Return path to installed software
        Dim iIndex As Integer = -1
        Dim sReturn As String = ""

        iIndex = GetIndex("ProjectWise", True)

        If iIndex > -1 Then
            sReturn = My.Computer.Registry.GetValue(dsSoftware.Tables(0).Rows(iIndex).Item("RegistryKey"), dsSoftware.Tables(0).Rows(iIndex).Item("RegistryString"), "")
        End If

        Return sReturn
    End Function

    '    Public Shared Sub SCEUpdate()
    '        'check for updates and extract from zip to public user location
    '        'pjr - should only be required until next profile refresh
    '        Dim sFile As String
    '        Dim sFileName As String
    '        Dim sToolsSourcePath As String = "Z:\SKM CAD\Bentley\SCE\Source\Tools\Cell Finder"
    '        Dim sToolsDestPath As String = "C:\Users\Public\Jacobs\MicroStation\Tools"

    '        If My.Computer.FileSystem.DirectoryExists(sToolsSourcePath) = True Then
    '            My.Computer.FileSystem.CreateDirectory(sToolsDestPath)
    '            For Each sFile In My.Computer.FileSystem.GetFiles(sToolsSourcePath, FileIO.SearchOption.SearchTopLevelOnly, "*.zip")
    '                sFileName = SCEFile.FileParse(sFile, SCEFile.FileParser.FileName)
    '                My.Computer.FileSystem.CreateDirectory(sToolsDestPath & "\Source")
    '                'Check zip has been updated
    '                If My.Computer.FileSystem.FileExists(sToolsDestPath & "\Source\" & sFileName & ".zip") = True Then
    '                    If FileDateTime(sFile) = FileDateTime(sToolsDestPath & "\Source\" & sFileName & ".zip") Then
    '                        GoTo NextTool
    '                    End If
    '                End If
    '                FileCopy(sFile, sToolsDestPath & "\Source\" & sFileName & ".zip")

    '                'Unzip Tool
    '                SCEZip.Unzip(sToolsDestPath & "\Source\" & sFileName & ".zip", sToolsDestPath & "\" & sFileName, False)
    'NextTool:
    '            Next
    '        End If

    '    End Sub

End Class
