Public Class JEGCore

    Private Shared bSiteOffice As Boolean
    Private Shared sConfigPath As String
    Private Shared sPublicConfigPath As String
    Private Shared sCoreDB As String
    Public Shared sLocalTempPath As String
    Private Shared sMSTNSCEKey As String
    Private Shared sMSTNSCEToolsKey As String
    Public Shared sNetworkPath As String
    Public Shared sNetworkClientPath As String
    Private Shared sSCEPath As String
    Private Shared sToolsPath As String
    Private Shared sUpdatePath As String
    Private Shared sUserDB As String
    Private Shared sUserPath As String
    Private Shared sUserTempPath As String
    Private Shared sWikiDB As String
    Private Shared sWikiPath As String
    Private Shared sClientPath As String
    Private Shared sXMLFileName As String
    Private Shared sEmailDomain As String
    Private Shared sIntranetPath As String
    Private Shared sDialogTitle As String
    Private Shared bAuthorisedSecret As Boolean
    Public Shared sLocalPath As String
    Private Shared sNetworkUpdatePath As String
    'Private Shared sPrestartPath As String
    'Private Shared sOffice As String
    Private Shared sSettingsFile As String
    'Private Shared sSettingsVersion As String
    'Public Shared oMSTN As MicroStationDGN.Application

    ''Public Shared cConfigurationFilePath = "C:\Program Files\SKMCAD\SCE_R17\Support"
    Public Shared MsgBoxCaption = My.Settings.MsgBoxCaption
    Public Shared dsSCE As DataSet
    Public Shared DocumentManagementSystem As String
    Public Shared SCERegKey As RegistryKey
    Public Shared Loaded As Boolean
    Public Shared bDebug As Boolean
    Public Shared Location As String

    Shared Sub identity(ByVal id As Long)
        If id = "654916916" Then
            bAuthorisedSecret = True
        End If
    End Sub

    'Shared Function Authoriseold() As Boolean
    '    'return boolean whether machine/user is authorised
    '    Dim bResult As Boolean

    '    If bAuthorisedSecret = False Then
    '        bResult = Locations.Check

    '        ''If My.Computer.FileSystem.DirectoryExists("C:\Program Files\ManageSoft\Launcher\Cache\Common\MicroStation SCE") = True Then
    '        ''    bResult = True
    '        ''Else
    '        ''    If Environment.UserName = "pripp" Then
    '        ''        bResult = True
    '        ''    End If
    '        ''End If
    '        If bResult = True Then
    '            Location = Locations.LocationName
    '        Else
    '            MsgBox("You are not authorised to run this software.", MsgBoxStyle.Critical, SCECore.MsgBoxCaption)
    '        End If
    '    Else
    '        bResult = True
    '    End If

    '    Return bResult
    'End Function

    Shared Function Authorise(Optional ByVal sLicensefile As String = "") As Boolean
        'Authorise use of the software
        Dim bReturn As Boolean = False

        AddLog("BulkModify", "Information", "PC Domain = " & My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\Tcpip\Parameters", "NV Domain", ""))

        Select Case UCase(My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\Tcpip\Parameters", "NV Domain", ""))
            Case UCase("Jacobs.com")
                bReturn = True
            Case UCase("Europe.Jacobs.com")
                bReturn = True
            Case UCase("jade.jacobs.com ")
                bReturn = True
            Case UCase("jec.jacobs.com")
                bReturn = True
            Case UCase("skmconsulting.com")
                bReturn = True
        End Select

        'Check license file
        If bReturn = False And sLicensefile <> "" Then
            bReturn = JEGLicense.IsValid(sLicensefile)
        End If

        Return bReturn
    End Function

    Public Shared Function GetSettings(Optional ByVal sFile As String = "") As Array
        'Return settings from file
        Dim sLine As String
        Dim aReturn() As String = Nothing
        Dim i As Integer = -1
        Dim sReturn As String = ""
        Dim j As Integer = 0
        Dim aSplit As Array = Nothing

        If sFile = "" Then
            sSettingsFile = My.Application.Info.DirectoryPath & "\Settings.ini"
        Else
            sSettingsFile = sFile
        End If

        'Get settings
        If My.Computer.FileSystem.FileExists(sSettingsFile) = True Then
            FileOpen(1, sSettingsFile, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
            Do While Not EOF(1)
                sLine = UCase(Trim(LineInput(1)))
                If SCEVB.IsCommented(sLine) = False Then
                    i = i + 1
                    ReDim Preserve aReturn(i)
                    aReturn(i) = sLine
                End If
            Loop
            FileClose(1)
        End If

        'Evaluate variables
        For i = 0 To UBound(aReturn)
            If InStr(aReturn(i), "[USERTEMP]") > 0 Then
                aReturn(i) = Replace(aReturn(i), "[USERTEMP]", JEGCore.UserTempPath)
            End If

            'If InStr(aReturn(i), "[") > 0 Then
            '    For j = 0 To UBound(aReturn)
            '        aSplit = Split(aReturn(i), "=")
            '        If InStr(aReturn(i), "[" & Trim(aSplit(0)) & "]") > 0 Then 'variable found
            '            aReturn(i) = Replace(aReturn(i), "[" & Trim(aSplit(0)) & "]", Trim(aSplit(0)))
            '        End If
            '    Next
            'End If
        Next

        Return aReturn
    End Function

    Public Shared Function GetSetting(ByVal sFile As String, ByVal sSetting As String) As String
        'Return setting from file
        Dim sLine As String
        Dim sReturn As String = ""
        Dim aSplit As Array = Nothing

        'Get settings
        If My.Computer.FileSystem.FileExists(sFile) = True Then
            FileOpen(1, sFile, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
            Do While Not EOF(1)
                sLine = UCase(Trim(LineInput(1)))
                If SCEVB.IsCommented(sLine) = False Then
                    If InStr(sLine, UCase(sSetting) & "=") > 0 Then 'setting found
                        aSplit = Split(sLine, "=")
                        sReturn = Trim(aSplit(1))
                    End If
                End If
            Loop
            FileClose(1)
        End If

        'Evaluate variables
        If InStr(sReturn, "[USERTEMP]") > 0 Then
            sReturn = Replace(sReturn, "[USERTEMP]", JEGCore.UserTempPath)
        End If

        Return sReturn
    End Function

    'Public Shared Sub GetSettingsTEMP()
    '    'get settings
    '    Dim sLine As String

    '    sSettingsFile = My.Application.Info.DirectoryPath & "\Settings.ini"

    '    If My.Computer.FileSystem.FileExists(sSettingsFile) = True Then
    '        FileOpen(1, sSettingsFile, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
    '        Do While Not EOF(1)
    '            sLine = UCase(Trim(LineInput(1)))
    '            If SCEVB.IsCommented(sLine) = False Then
    '                If InStr(UCase(sLine), "VERSION") > 0 Then
    '                    sLine = Replace(sLine, "VERSION", "")
    '                    sLine = Trim(Replace(sLine, "=", ""))
    '                    sSettingsVersion = sLine
    '                ElseIf InStr(UCase(sLine), "OFFICE") > 0 Then
    '                    sLine = Replace(sLine, "OFFICE", "")
    '                    sLine = Trim(Replace(sLine, "=", ""))
    '                    sOffice = sLine
    '                ElseIf InStr(UCase(sLine), "NETWORKPATH") > 0 Then
    '                    sLine = Replace(sLine, "NETWORKPATH", "")
    '                    sLine = Trim(Replace(sLine, "=", ""))
    '                    sNetworkPath = sLine
    '                ElseIf InStr(UCase(sLine), "LOCALPATH") > 0 Then
    '                    sLine = Replace(sLine, "LOCALPATH", "")
    '                    sLine = Trim(Replace(sLine, "=", ""))
    '                    sLocalPath = sLine
    '                ElseIf InStr(UCase(sLine), "UPDATEPATH") > 0 Then
    '                    sLine = Replace(sLine, "UPDATEPATH", "")
    '                    sLine = Trim(Replace(sLine, "=", ""))
    '                    sNetworkUpdatePath = sLine
    '                ElseIf InStr(UCase(sLine), "LOCALTEMP") > 0 Then
    '                    sLine = Replace(sLine, "LOCALTEMP", "")
    '                    sLine = Trim(Replace(sLine, "=", ""))
    '                    sLocalTempPath = sLine
    '                ElseIf InStr(UCase(sLine), "PRESTART") > 0 Then
    '                    sLine = Replace(sLine, "PRESTART", "")
    '                    sLine = Trim(Replace(sLine, "=", ""))
    '                    sPrestartPath = sLine
    '                End If
    '            End If
    '        Loop
    '        FileClose(1)
    '    End If
    'End Sub

    Public Shared Sub AddLog(ByVal sProgram As String, ByVal sSeverity As String, ByVal sDescription As String)
        'add specified description to log file
        'time/date,WARNING/ERROR/Information,Description
        Dim sLogFile As String = LogFile(sProgram)

        My.Computer.FileSystem.CreateDirectory(SCEFile.FileParse(sLogFile, SCEFile.FileParser.Path))

        FileOpen(99, sLogFile, OpenMode.Append, OpenAccess.Write, OpenShare.LockWrite)
        PrintLine(99, Now & "," & sSeverity & "," & sDescription)
        FileClose(99)
    End Sub

    Public Shared Sub Initialise()
        If JEGCore.bDebug = True Then JEGCore.RecordSequence(System.Reflection.MethodBase.GetCurrentMethod.Module.Name, System.Reflection.MethodBase.GetCurrentMethod.Name())

        sSCEPath = LocalJMEPath()
        sCoreDB = SCECoreDB()
        sUpdatePath = JEGCore.sUpdatePath
        sToolsPath = ToolsPath()
        bSiteOffice = SiteOffice(True)

        SCEClients.Initialise()
        SCESoftware.Initialise()
        SCEFavourite.LoadClients()
        Loaded = True
    End Sub

    'Shared Function SettingsVersion()
    '    'Return version
    '    Return sSettingsVersion
    'End Function

    'Shared Function Office()
    '    'Return office
    '    Return sOffice
    'End Function

    'Shared Function PrestartPath()
    '    'Return prestart path
    '    Return sPrestartPath
    'End Function

    Shared Function SettingsFilePath()
        'Return settings file path
        Return sSettingsFile
    End Function

    Shared Function SiteOffice(ByVal bInit As Boolean) As Boolean
        'Return if current office is offsite
        Dim sSiteOffice As String = ""

        If bInit = True Then
            sSiteOffice = Trim(SCEXML.GetConfigurationValue("MSTNSiteOffice"))

            If sSiteOffice <> "" Then
                bSiteOffice = True
            End If
        End If
        If JEGCore.bDebug = True Then JEGCore.RecordSequence(System.Reflection.MethodBase.GetCurrentMethod.Module.Name, System.Reflection.MethodBase.GetCurrentMethod.Name(), bSiteOffice)
        Return bSiteOffice
    End Function

    Shared Function Password() As String
        'Return SCE password
        Dim sReturn As String = ""

        If Authorise() = True Then
            sReturn = My.Settings.Password
        End If

        Return sReturn
    End Function

    Public Shared Function EmailDomain() As String
        '@GlobalSKM.com
        If sEmailDomain = "" Then
            sEmailDomain = SCEXML.GetConfigurationValue("EmailDomain")
        End If

        ' ''pjrxxx - testing only, EmailDomain needs to be added to XML
        ''If sEmailDomain = "" Then
        ''    sEmailDomain = "@GlobalSKM.com"
        ''End If
        If JEGCore.bDebug = True Then JEGCore.RecordSequence(System.Reflection.MethodBase.GetCurrentMethod.Module.Name, System.Reflection.MethodBase.GetCurrentMethod.Name(), sEmailDomain)
        Return sEmailDomain
    End Function

    'Shared Function SCEConfigPath() As String
    '    'C:\ProgramData\Jacobs\MicroStation Environment

    '    If sConfigPath = "" Then
    '        sConfigPath = "C:\ProgramData\Jacobs\MicroStation Environment" 'Installed path
    '    End If
    '    If JEGCore.bDebug = True Then JEGCore.RecordSequence(System.Reflection.MethodBase.GetCurrentMethod.Module.Name, System.Reflection.MethodBase.GetCurrentMethod.Name(), sConfigPath)
    '    Return sConfigPath
    'End Function

    Shared Function SCEPublicConfigPath() As String
        'C:\Users\Public\SKMCAD\MicroStation SCE

        If sPublicConfigPath = "" Then
            sPublicConfigPath = "C:\Users\Public\Jacobs\MicroStation Environment"
        End If
        If JEGCore.bDebug = True Then JEGCore.RecordSequence(System.Reflection.MethodBase.GetCurrentMethod.Module.Name, System.Reflection.MethodBase.GetCurrentMethod.Name(), sPublicConfigPath)
        Return sPublicConfigPath
    End Function

    Shared Function LocalTempPath()
        'C:\Users\pripp\AppData\Local\Temp\SKMCAD\MicroStation SCE
        If sLocalTempPath = "" Then
            sLocalTempPath = SCEXML.GetConfigurationValue("MSTNLocalTemp")
            If My.Computer.FileSystem.DirectoryExists(sLocalTempPath) = False Then
                Try
                    My.Computer.FileSystem.CreateDirectory(sLocalTempPath)
                Catch ex As Exception

                End Try
            End If
        End If
        If JEGCore.bDebug = True Then JEGCore.RecordSequence(System.Reflection.MethodBase.GetCurrentMethod.Module.Name, System.Reflection.MethodBase.GetCurrentMethod.Name(), sLocalTempPath)
        Return sLocalTempPath
    End Function

    Shared Function NetworkAccess() As Boolean
        'Return boolean if network access is available
        Return My.Computer.FileSystem.DirectoryExists(JEGCore.NetworkUpdatePath)
    End Function

    Shared Function NetworkPath()
        'N:\SKMCADCFG\MicroStation
        If sNetworkPath = "" Then
            sNetworkPath = SCEXML.GetConfigurationValue("MSTNNetworkPath")
        End If
        If JEGCore.bDebug = True Then JEGCore.RecordSequence(System.Reflection.MethodBase.GetCurrentMethod.Module.Name, System.Reflection.MethodBase.GetCurrentMethod.Name(), sNetworkPath)
        Return sNetworkPath
    End Function

    Shared Function NetworkClientPath()
        'N:\SKMCADCFG\MicroStation\Clients
        'If sNetworkPath = "" Then
        '    sNetworkPath = SCEXML.GetConfigurationValue("MSTNNetworkPath")
        'End If
        'If JEGCore.bDebug = True Then JEGCore.RecordSequence(System.Reflection.MethodBase.GetCurrentMethod.Module.Name, System.Reflection.MethodBase.GetCurrentMethod.Name(), sNetworkPath)
        Return sNetworkClientPath
    End Function

    Shared Function NetworkUpdatePath()
        'N:\CAD\MSTN\Update
        If sNetworkUpdatePath = "" Then
            sNetworkUpdatePath = SCEXML.GetConfigurationValue("MSTNNetworkPath") & "\Update"
        End If

        Return sNetworkUpdatePath
    End Function

    ' ''Shared Function NetworkClientPath()
    ' ''    'N:\CAD\MSTN\Working
    ' ''    Return SCEXML.GetConfigurationValue("MSTNNetworkPath") & "\Working"
    ' ''End Function

    'Shared Function RegionPath() As String
    '    'D:\Documents and Settings\skmcad\MSTN\Region
    '    Dim sReturn As String

    '    sReturn = SCEConfigPath() & "\Region"

    '    Return sReturn
    'End Function

    'Shared Function SCEPath() As String
    '    'replace by LocalJMEPath
    '    'C:\Program Files\SKMCAD\MicroStation SCE

    '    If sSCEPath = "" Then
    '        sSCEPath = SCEXML.GetConfigurationValue("MSTNSCERoot")
    '    End If
    '    If JEGCore.bDebug = True Then JEGCore.RecordSequence(System.Reflection.MethodBase.GetCurrentMethod.Module.Name, System.Reflection.MethodBase.GetCurrentMethod.Name(), sSCEPath)
    '    Return sSCEPath
    'End Function

    'Shared Function ProjectWise() As Boolean
    '    'Display ProjectWise controls
    '    Dim bReturn As Boolean = False

    '    If UCase(SCEXML.GetConfigurationValue("ProjectWise")) = "TRUE" Then
    '        bReturn = True
    '    End If

    '    Return bReturn
    'End Function

    Shared Function ToolsPath() As String
        'C:\ProgramData\Jacobs\MicroStation Environment\Tools

        If sToolsPath = "" Then
            sToolsPath = LocalJMEPath() & "\Tools"
        End If
        If JEGCore.bDebug = True Then JEGCore.RecordSequence(System.Reflection.MethodBase.GetCurrentMethod.Module.Name, System.Reflection.MethodBase.GetCurrentMethod.Name(), sToolsPath)
        Return sToolsPath
    End Function

    Shared Function UserPath() As String
        'C:\Users\pripp\AppData\Local\SKMCAD\MicroStation SCE"

        If sUserPath = "" Then
            sUserPath = Environment.GetEnvironmentVariable("LOCALAPPDATA") & "\Jacobs\MicroStation Environment"
            'sUserPath = "C:\Users\" & Environment.UserName & "\AppData\Local\Jacobs\MicroStation Environment"
        End If
        If JEGCore.bDebug = True Then JEGCore.RecordSequence(System.Reflection.MethodBase.GetCurrentMethod.Module.Name, System.Reflection.MethodBase.GetCurrentMethod.Name(), sUserPath)
        Return sUserPath
    End Function

    Shared Function UserTempPath() As String
        'D:\Documents and Settings\skmcad\MSTN\User

        If sUserTempPath = "" Then
            sUserTempPath = Environment.GetEnvironmentVariable("TEMP") & "\Jacobs\MicroStation Environment"
            'sUserTempPath = "C:\Users\" & Environment.UserName & "\AppData\Local\Temp\Jacobs\MicroStation Environment"
        End If
        If JEGCore.bDebug = True Then JEGCore.RecordSequence(System.Reflection.MethodBase.GetCurrentMethod.Module.Name, System.Reflection.MethodBase.GetCurrentMethod.Name(), sUserTempPath)
        Return sUserTempPath
    End Function

    Shared Function WikiPath()
        '"\Tools\SCE Documentation\Wiki"
        If sWikiPath = "" Then
            sWikiPath = "C:\Users\pripp\Desktop\Team Site"
        End If

        Return sWikiPath
    End Function

    Shared Function ClientPath() As String
        'C:\ProgramData\Jacobs\MicroStation Environment\Clients

        If sClientPath = "" Then
            sClientPath = LocalJMEPath() & "\Clients"
        End If
        If JEGCore.bDebug = True Then JEGCore.RecordSequence(System.Reflection.MethodBase.GetCurrentMethod.Module.Name, System.Reflection.MethodBase.GetCurrentMethod.Name(), sClientPath)
        Return sClientPath
    End Function

    ''Shared Function NetworkUpdateINI() As String
    ''    'N:\CAD\MSTN\Update\MicroStation SCE.ini
    ''    Return NetworkPath() & "\Update\MicroStation SCE.ini"
    ''End Function

    Shared Function SCECoreDB() As String
        'C:\Program Files\SKMCAD\MicroStation SCE\Database\SCE Core.accdb

        If sCoreDB = "" Then
            'sCoreDB = SCEXML.GetConfigurationValue("MSTNSCEDB")
            sCoreDB = LocalJMEPath() & "Database\Core.accdb"
        End If
        If JEGCore.bDebug = True Then JEGCore.RecordSequence(System.Reflection.MethodBase.GetCurrentMethod.Module.Name, System.Reflection.MethodBase.GetCurrentMethod.Name(), sCoreDB)
        Return sCoreDB
    End Function

    Shared Function DialogTitle() As String
        'Dialog title override

        If sDialogTitle = "" Then
            sDialogTitle = SCEXML.GetConfigurationValue("MSTNDialogTitle")
        End If

        Return sDialogTitle
    End Function

    Shared Function UserDatabase()
        'D:\Documents and Settings\skmcad\MSTN\User\Database\SCE User.accdb

        If sUserDB = "" Then
            sUserDB = UserPath() & "\Database\User.accdb"
        End If
        If JEGCore.bDebug = True Then JEGCore.RecordSequence(System.Reflection.MethodBase.GetCurrentMethod.Module.Name, System.Reflection.MethodBase.GetCurrentMethod.Name(), sUserDB)
        Return sUserDB
    End Function

    Shared Function WikiDatabase()
        '"\Tools\SCE Documentation\Wiki"
        If sWikiDB = "" Then
            sWikiDB = "C:\Users\pripp\Desktop\Team Site\MicroStation Team Site.accdb"
        End If

        Return sWikiDB
    End Function

    Shared Function MSTNSCEKey() As String
        'MSTN SCE Registry key
        'HKEY_CURRENT_USER\Software\SKMCAD\MicroStation\SCE

        If sMSTNSCEKey = "" Then
            sMSTNSCEKey = My.Settings.MSTNSCEKey
        End If
        If JEGCore.bDebug = True Then JEGCore.RecordSequence(System.Reflection.MethodBase.GetCurrentMethod.Module.Name, System.Reflection.MethodBase.GetCurrentMethod.Name(), sMSTNSCEKey)
        Return sMSTNSCEKey
    End Function

    Shared Function SKMCADKey() As String
        'SKM CAD Registry key
        'HKEY_CURRENT_USER\Software\SKMCAD
        If JEGCore.bDebug = True Then JEGCore.RecordSequence(System.Reflection.MethodBase.GetCurrentMethod.Module.Name, System.Reflection.MethodBase.GetCurrentMethod.Name(), My.Settings.JacobsKey)
        Return My.Settings.JacobsKey
    End Function

    Shared Function MSTNToolsKey() As String
        'MSTN Tools Registry key root
        'HKEY_CURRENT_USER\Software\SKMCAD\MicroStation\SCE\Tools
        If sMSTNSCEToolsKey = "" Then
            sMSTNSCEToolsKey = MSTNSCEKey() & "\Tools"
        End If

        Return sMSTNSCEToolsKey
    End Function

    Public Shared Function LocalJMEPath() As String
        'return local file path
        'C:\ProgramData\Jacobs\MicroStation Environment
        'C:\Users\Public\Jacobs\MicroStation Environment
        'Dim sReturn As String = ""

        'If sLocalPath = "" Then
        '    If My.Computer.FileSystem.DirectoryExists(My.Settings.LocalFilePath32Bit) = True Then
        '        sLocalPath = My.Settings.LocalFilePath32Bit
        '    Else
        '        sLocalPath = My.Settings.LocalFilePath64Bit
        '    End If
        'End If

        'If JEGCore.bDebug = True Then JEGCore.RecordSequence(System.Reflection.MethodBase.GetCurrentMethod.Module.Name, System.Reflection.MethodBase.GetCurrentMethod.Name(), sLocalPath)
        Return sLocalPath
    End Function

    Public Shared Function LocalPublicPath() As String
        'return local public path
        'C:\ProgramData\Jacobs\MicroStation Environment
        'Dim sReturn As String = ""

        Return "C:\Users\Public\Jacobs\MicroStation Environment"

        'If JEGCore.bDebug = True Then JEGCore.RecordSequence(System.Reflection.MethodBase.GetCurrentMethod.Module.Name, System.Reflection.MethodBase.GetCurrentMethod.Name(), sLocalPath)
        Return sLocalPath
    End Function

    Public Shared Function LogFile(ByVal sProgram As String) As String
        'log file full path
        Return "C:\Users\public\Jacobs\Logs\" & sProgram & " Log.csv"
    End Function

    Shared Function XMLFileName() As String
        'C:\Program Files\SKMCAD\SCE_R17\Support\ConfigurationII.XML
        If sXMLFileName = "" Then
            sXMLFileName = LocalJMEPath() & My.Settings.UpdateFile
            sXMLFileName = Replace(sXMLFileName, "[ProgramFiles]", My.Computer.FileSystem.SpecialDirectories.ProgramFiles, , , Microsoft.VisualBasic.CompareMethod.Text)
            ''If My.Computer.FileSystem.FileExists(sXMLFileName) = False Then
            ''    sXMLFileName = My.Settings.LocalFilePath & My.Settings.UpdateFile1
            ''    sXMLFileName = Replace(sXMLFileName, "[ProgramFiles]", My.Computer.FileSystem.SpecialDirectories.ProgramFiles, , , Microsoft.VisualBasic.CompareMethod.Text)
            ''End If

            Dim DocumentsAndSettingsFolder As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            DocumentsAndSettingsFolder = Replace(DocumentsAndSettingsFolder, Environment.UserName & "\My Documents", "", , , Microsoft.VisualBasic.CompareMethod.Text)
            sXMLFileName = Replace(sXMLFileName, "[DocumentsAndSettings]", DocumentsAndSettingsFolder)
        End If

        If My.Computer.FileSystem.FileExists(sXMLFileName) = False Then
            sXMLFileName = My.Application.Info.DirectoryPath & "\" & My.Settings.UpdateFile
        End If

        Return sXMLFileName
    End Function

    Shared Function IsAdmin(ByVal sUserName As String) As Boolean
        'Return boolean if specified user is an admin (from CoP site)
        Dim dsAdmins As DataSet
        Dim bReturn As Boolean = False
        Dim i As Integer

        dsAdmins = SCEAccess.GetColumnDB(SCECoreDB, "Jacobs MicroStation Administrators", "Name", "Title", "Administrator", True, False, True)

        If SCEVB.IsNullDataSet(dsAdmins) = False Then
            For i = 0 To dsAdmins.Tables(0).Rows.Count - 1
                If UCase(sUserName) = UCase(dsAdmins.Tables(0).Rows(i).Item("Name")) Then 'Title
                    bReturn = True
                    Exit For
                End If
            Next
        End If

        Return bReturn
    End Function

    Shared Function ShutdownFile() As String
        'Shutdown file
        Return "Shutdown.txt"
    End Function

    Shared Function IntranetPath() As String
        'http:\\... - SKM
        'P:\SKM_Homepage\index.htm - Riotinto Perth

        If sIntranetPath = "" Then
            sIntranetPath = SCEXML.GetConfigurationValue("MSTNIntranetPath")
        End If

        Return sIntranetPath
    End Function

    Public Shared Function ClientDBFullName(ByVal sRegion As String, ByVal sClient As String) As String
        'return client database fullname
        Dim sReturn As String = JEGCore.ClientPath & "\" & sRegion & "\" & sClient & "\Database\SCE Client - " & sClient & ".accdb"

        Return sReturn
    End Function

    ''' <summary>
    ''' add log of procedures executed
    ''' </summary>
    ''' <param name="AssemblyName"></param>
    ''' <param name="ModuleName"></param>
    ''' <param name="msg"></param>
    ''' <remarks></remarks>
    Public Shared Sub RecordSequence(ByVal AssemblyName As String, ByVal ModuleName As String, Optional ByVal msg As String = "")
        'add following command to all/specific assemblies
        'SKM.Utilities.RecordSequence(System.Reflection.MethodBase.GetCurrentMethod.Module.Name, System.Reflection.MethodBase.GetCurrentMethod.Name())

        Try
            Dim slogfile As String = JEGCore.UserTempPath & "\Record Sequence.csv"

            If Not System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(slogfile)) Then
                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(slogfile))
            End If

            '' Check the to see if the file is more than 10 days old and delete it if thats the case.
            ''If System.IO.File.Exists(slogfile) Then
            ''    Dim filedate As Date = System.IO.File.GetCreationTime(slogfile)
            ''    Dim TimeElapsed As TimeSpan = DateTime.Now.Subtract(filedate)
            ''    If TimeElapsed.Days > 10 Then
            ''        System.IO.File.Delete(slogfile)
            ''    End If
            ''End If

            Dim now As Date = DateAndTime.Now

            System.IO.File.AppendAllText(slogfile, now.Year.ToString & "-" & now.Month.ToString & "-" & now.Day.ToString & _
                                         "," & now.TimeOfDay.ToString & "," & ModuleName & "," & AssemblyName & "," & msg & vbCrLf)
        Catch ex As Exception
            'SKM.Utilities.ExceptionMessageDisplay(ex, System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod.Module.Name)
        End Try
    End Sub

    Public Shared Function TemplateDB(Optional ByVal sOverrideSCEPath As String = "") As String
        'full path to empty Access Database
        If sOverrideSCEPath = "" Then
            Return "C:\Program Files (x86)\Jacobs\MicroStation\Templates\Template.accdb"
        Else
            Return sOverrideSCEPath & "\Templates\Template.accdb"
        End If
    End Function

    Public Shared Function GetTempIndex() As String
        'return unique index based on current time date 
        Dim sReturn As String = Format(Now, "yyyyMMddhhmmtt")

        Return sReturn
    End Function

    Private Shared Function SCEConfigPath() As String
        Throw New NotImplementedException
    End Function

End Class
