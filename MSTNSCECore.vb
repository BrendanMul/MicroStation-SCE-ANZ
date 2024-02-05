Imports MicroStationSCE.SCEAccess1

<ComClass(MSTNSCECore.ClassId, MSTNSCECore.InterfaceId, MSTNSCECore.EventsId)> _
Public Class MSTNSCECore

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "cf7a97b2-5ad5-4005-a69f-b9242b91dec4"
    Public Const InterfaceId As String = "80d57221-e4e1-4749-9d51-fe5427df816b"
    Public Const EventsId As String = "912a00a9-fdb4-4fdc-b454-3668ad1b1752"
#End Region

    'Module
    Public Shared iProjectIndex As Long
    Public Shared SCELoaded As Boolean

    Public SCECells As New MSTNClientCells
    Public ClientCells As New MSTNClientCells
    Public ClientSettings As New ClientSettings
    Public ClientText As New MSTNClientText
    Public ClientTools As New ClientTools
    Public Drawing As New Drawing
    ''Public SCE As New SCE
    Public SCESettings As New SCESettings
    Public SCETables As New SCETables
    Public Software As New SCESoftware
    Public StampGuide As New StampGuide
    Public TitleBorder As New TitleBorder
    Public TitleBorderPosition As New TitleBorderPosition
    Public TitleBorders As New TitleBorders
    Public LocalUserName As New UserName

    Public Const cSKMCADKey = "HKEY_LOCAL_MACHINE\SOFTWARE\SKMCAD"
    Public Const cSCEKey = "HKEY_LOCAL_MACHINE\SOFTWARE\SKMCAD\MicroStation\SCE"
    Public Const cLocalSKMCADKey = "HKEY_CURRENT_USER\SOFTWARE\SKMCAD"
    Public Const cLocalSCEKey = "HKEY_CURRENT_USER\Software\SKMCAD\MicroStation\SCE"

    'SCE Class
    Public Shared Mode As String
    Public Shared Loaded As Boolean

    'Client
    Public Shared Client As String
    Public Shared Region As String

    'Databases
    Public Shared ClientDB As String
    Public Shared CoreDB As String
    Public Shared NetworkDB As String
    Public Shared UserDB As String

    'Paths
    Public Shared ClientWorkspacePath As String
    Public Shared VBAPath As String
    Public Shared SCERegionPath As String
    Public Shared WorkingPath As String
    Public Shared UserPath As String

    Private Shared _myInstance As MicroStationDGN.Application
    Public Shared oMSTN As MicroStationDGN.Application = GetInstance()

    Public Shared Function GetInstance() As MicroStationDGN.Application
        If _myInstance Is Nothing Then
            _myInstance = New MicroStationDGN.Application 'TODO: Add check if App is already(running)
        End If
        Return _myInstance
    End Function

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Public Function Password() As String
        'SCE password
        Dim sReturn As String = ""
        Return SCECore.Password
    End Function

    Sub Load()

        Dim j As Long
        ''Dim ProjectPath As New ProjectPath
        Dim sDBValue As String
        Dim sDBExtension As String

        'Delete AutoSCE file
        If My.Computer.FileSystem.FileExists(oMSTN.ActiveWorkspace.ConfigurationVariableValue("_USTN_APPL") & "AutoSCE.cfg") = True Then
            My.Computer.FileSystem.DeleteFile(oMSTN.ActiveWorkspace.ConfigurationVariableValue("_USTN_APPL") & "AutoSCE.cfg")
        End If

        'Core
        CoreDB = oMSTN.ActiveWorkspace.ConfigurationVariableValue("_SCE_DB")
        SCERegionPath = oMSTN.ActiveWorkspace.ConfigurationVariableValue("SCE_REGION")
        WorkingPath = oMSTN.ActiveWorkspace.ConfigurationVariableValue("SCE_WORKINGROOT")
        VBAPath = oMSTN.ActiveWorkspace.ConfigurationVariableValue("SCE_VBA")
        UserPath = oMSTN.ActiveWorkspace.ConfigurationVariableValue("SCE_USER")
        ClientWorkspacePath = oMSTN.ActiveWorkspace.ConfigurationVariableValue("SCE_CLIENT_WORKSPACE")

        SCETables = Nothing
        SCETables.Load()

        'Client
        Region = oMSTN.ActiveWorkspace.ConfigurationVariableValue("SCE_CLIENT_REGION")
        Client = oMSTN.ActiveWorkspace.ConfigurationVariableValue("SCE_CLIENT_NAME")
        ClientDB = oMSTN.ActiveWorkspace.ConfigurationVariableValue("_SCE_CLIENT_DB")

        If My.Computer.FileSystem.FileExists(ClientDB) = False Then
            ClientDB = oMSTN.ActiveWorkspace.ConfigurationVariableValue("_SCE_CLIENT_DBPATH") & "SCE Client - " & oMSTN.ActiveWorkspace.ConfigurationVariableValue("SCE_CLIENT_NAME") & ".ACCDB"
        End If

        'Network
        NetworkDB = oMSTN.ActiveWorkspace.ConfigurationVariableValue("_SCE_NET_DB")

        'User
        UserDB = oMSTN.ActiveWorkspace.ConfigurationVariableValue("_SCE_USER_DB")

        LocalUserName = Nothing
        LocalUserName.Load()
        Drawing = Nothing
        SCELoaded = True
    End Sub

    Function GetDatabase(ByVal sTable As String)
        'Return database fullname using specified database type
        Dim i As Long
        Dim vDBs As Object
        Dim sDBFullName As String

        vDBs = SCETables.GetDatabases(sTable)

        For i = 0 To UBound(vDBs)
            'Get DB
            Select Case UCase(vDBs(i))
                Case "NETWORK"
                    sDBFullName = NetworkDB
                Case "CLIENT"
                    sDBFullName = ClientDB
                Case "USER"
                    sDBFullName = UserDB
                Case "CORE"
                    sDBFullName = CoreDB
            End Select

            'Check exists
            If My.Computer.FileSystem.FileExists(sDBFullName) = True Then
                If TableExists(sDBFullName, sTable) = True Then
                    GetDatabase = sDBFullName
                    Exit For
                End If
            End If
        Next
        Erase vDBs
    End Function

    Function Regions()
        Regions = SCEFile.GetFolders(SCERegionPath, all)
    End Function

    Function Clients(ByVal sRegion As String)
        Clients = SCEFile.GetFolders(SCERegionPath & "\" & sRegion, all)
    End Function

End Class


