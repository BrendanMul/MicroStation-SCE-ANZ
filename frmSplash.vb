'Imports SCECoreDLL.File

Public NotInheritable Class frmSplash

    Private bLoading As Boolean
    Private bHide As Boolean

    'Private Sub tmrTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrTimer.Tick
    '    'If bLoading = True Then
    '    If Me.Opacity = 0 Then
    '        tmrTimer.Enabled = False
    '        If bHide = False Then
    '            sceclientselector.frmclientselector.Show()
    '            SCE.tmrDisplay.Enabled = True
    '        End If
    '        Me.Opacity = 0
    '        Me.Hide()
    '    Else
    '        Me.Opacity = Me.Opacity - 0.1
    '        tmrTimer.Enabled = False
    '        tmrTimer.Interval = 20
    '        tmrTimer.Enabled = True
    '    End If
    '    'End If
    'End Sub

    Public Sub New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
    End Sub

    Private Sub frmSplash_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim dsFavourites As DataSet
        'Dim sUser As String
        Dim bNewUser As Boolean = False
        'Dim i As Integer
        Dim sLine As String = ""

        ' ''Check for updates
        ''If My.Computer.FileSystem.FileExists(Core.NetworkUpdatePath & "\MicroStation SCE\Version.txt") = True Then
        ''    FileOpen(1, Core.NetworkUpdatePath & "\MicroStation SCE\Version.txt", OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
        ''    Do While Not EOF(1)
        ''        sLine = Trim(LineInput(1))
        ''        If Trim(sLine) <> "" Then
        ''            Exit Do
        ''        End If
        ''    Loop
        ''    FileClose(1)
        ''    If My.Computer.Registry.GetValue(Core.MSTNSCEKey, "SCEVersion", "") <> sLine Then
        ''        If My.Computer.FileSystem.FileExists(Core.SCEPath & "\Update MicroStation SCE.bat") = True Then
        ''            My.Computer.Registry.SetValue(Core.MSTNSCEKey, "SCEVersion", sLine)
        ''            Shell(Core.SCEPath & "\Update MicroStation SCE.bat", AppWinStyle.MinimizedNoFocus)
        ''            End
        ''        End If
        ''    End If
        ''End If

        ''InitialiseDocumentData()

        ''Version.Text = System.String.Format(Version.Text, My.Application.Info.Version.Major, My.Application.Info.Version.Minor)
        ''SCE.sCoreDB = Core.SCEDatabase
        ''sUser = Environment.UserName
        ''SCE.sUserPath = Core.UserPath()
        ''SCE.sHelpFilePath = sSCEPath & "\Documentation\MicroStation SCE - User Guide.chm"
        ''If FileExists(SCE.sHelpFilePath) = False Then
        ''    SCE.sHelpFilePath = ""
        ''End If

        ' ''Check user is configured
        ''If FolderExists(SCE.sUserPath) = False Then
        ''    bNewUser = True
        ''End If
        ''ConfigureUser()

        ''sProjectDrive = XML.GetConfigurationValue("ProjectDrive")

        'Favourites
        ''SCECore.sUserDB = Core.UserDatabase
        ''SCE.sWikiDB = Core.WikiDatabase
        ''sRegionPath = Core.RegionPath
        ''sWikiPath = Core.WikiPath
        ''dsFavourites = SCEAccess.GetColumnDB(SCE.sUserDB, "LUFavourites", "CustomName", "", "", True, False, True)

        ''If FolderExists(Core.NetworkClientPath & "\Database") = True Then
        ''    SCECore.ClientPath = Core.NetworkClientPath
        ''    bNetwork = True
        ''Else
        ''    SCECore.ClientPath = Core.ClientPath
        ''    bNetwork = False
        ''End If

        'Check client builds exist
        If My.Computer.FileSystem.DirectoryExists(JEGCore.ClientPath) = False Then
            'MsgBox("No client builds were located." & Chr(10) & Chr(10) & "MicroStation SCE will now exit.", MsgBoxStyle.Critical, SCECore.MsgBoxCaption)
            'End
            GoTo ExitSub
        End If

        'Extract databases if network access is not found
        ''If bNetwork = False Then
        ''    lblStatus.Text = "Initialising databases..."
        ''    My.Application.DoEvents()
        ''    Me.Cursor = Cursors.WaitCursor
        ''    Zip.UnzipDatabases()
        ''    Me.Cursor = Cursors.Default
        ''    lblStatus.Text = ""
        ''    My.Application.DoEvents()
        ''End If

        'Update the SCE if necessary
        ''lblStatus.Text = "Updating SCE functionality..."
        ''My.Application.DoEvents()
        ''Me.Cursor = Cursors.WaitCursor
        ''UpdateSCE()
        ''Me.Cursor = Cursors.Default
        ''lblStatus.Text = ""

        'pjr - not required??
        'If bNewUser = True Then
        '    frmUser.Show()
        'End If
        dsFavourites = Nothing
        bLoading = True

        'Create local temp path otherwise users can't use CTRL C for copy
        If My.Computer.FileSystem.DirectoryExists(JEGCore.LocalTempPath) = False Then
            My.Computer.FileSystem.CreateDirectory(JEGCore.LocalTempPath)
        End If

        'Process command line switches
        'If My.Application.CommandLineArgs.Count = 0 Then
        '    'Normal mode
        '    tmrTimer.Enabled = False
        '    tmrTimer.Enabled = True
        'Else
        '    For i = 0 To (My.Application.CommandLineArgs.Count - 1)
        '        If UCase(Trim(My.Application.CommandLineArgs.Item(i))) = "/Q" Then 'Quiet mode
        '            tmrTimer.Enabled = False
        '            bHide = True
        '        End If
        '    Next
        'End If
        Me.Dispose()
ExitSub:
    End Sub

    Private Sub frmSplash_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        If bHide = True Then
            tmrTimer.Enabled = False
            tmrTimer.Enabled = True
        End If
    End Sub

    Private Sub MainLayoutPanel_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MainLayoutPanel.Paint

    End Sub
End Class
