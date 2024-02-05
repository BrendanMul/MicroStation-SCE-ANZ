Imports JMECore.SCEAccess
Imports JMECore.SCEFile
Imports JMECore.SCEXML

Public Class SCEWiki

    Public Shared sRunWikiNetwork As String
    Public Shared sWikiWebAddress As String = "http://km.skmconsulting.com/cop/cad/Microstation/MicroStation%20Wiki/Forms/AllPages.aspx"

    Shared Function SearchWebAddress(ByVal sSearchText As String) As String
        'Return web address to search wiki for specified text
        Return "http://km.skmconsulting.com/cop/cad/Microstation/_layouts/OSSSearchResults.aspx?k=" & sSearchText & "&cs=This%20List&u=http%3A%2F%2Fkm.skmconsulting.com%2Fcop%2Fcad%2FMicrostation%2FMicroStation%20Wiki"
    End Function

    Shared Function HasWiki(ByVal sProgram As String) As Boolean
        'Return if specified program has current wiki's
        Dim dsWiki As DataSet = Nothing
        Dim bReturn As Boolean = False

        If IsWikisAvailable() = True Then
            dsWiki = SCEAccess.GetColumnDB(JEGCore.WikiDatabase, "MicroStation Wiki", "Keywords", "Keywords", "Program:" & sProgram, True, False, False)
            If dsWiki.Tables(0).Rows.Count > 0 Then
                bReturn = True
            End If
        Else
            dsWiki = SCEAccess.GetColumnDB(JEGCore.WikiDatabase, "MicroStation Wiki Cache", "Keywords1", "Keywords1", "Program:" & sProgram, True, False, False)
            If dsWiki.Tables(0).Rows.Count > 0 Then
                bReturn = True
            End If
        End If

        Return bReturn
    End Function

    Shared Function IsWikisAvailable() As Boolean
        'Test speed of wiki info return and configure todays setting
        Dim dsWiki As DataSet
        Dim oStartTime As Date = Now
        Dim oEndTime As Date
        Dim bResult As Boolean = False
        Dim sLine As String = ""

        If sRunWikiNetwork = "" Then
            If My.Computer.FileSystem.FileExists(JEGCore.WikiPath & "\Wiki.ini") = True Then
                Dim oDate As Date = FileDateTime(JEGCore.WikiPath & "\Wiki.ini")
                If Format(oDate, "dd/MM/yyyy") = Format(Now, "dd/MM/yyyy") Then 'only get variable if file updated today
                    FileClose(1)
                    FileOpen(1, JEGCore.WikiPath & "\Wiki.ini", OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
                    Do While Not EOF(1)
                        sLine = Trim(LineInput(1))
                        Exit Do
                    Loop
                    FileClose(1)

                    If UCase(sLine) = "NETWORK" Then
                        bResult = True
                    End If
                Else
                    GoTo TestWiki
                End If
            Else 'Test wiki connection speed
TestWiki:
                If AccessSharePoint() = True Then
                    dsWiki = GetColumnDB(JEGCore.WikiDatabase, "MicroStation Wiki", "Keywords", "Keywords", "Program", True, False, False)
                    oEndTime = Now
                    If DateDiff(DateInterval.Second, oStartTime, oEndTime) < 3 Then
                        bResult = True
                    End If
                End If
                If My.Computer.FileSystem.DirectoryExists(JEGCore.WikiPath & "\Wiki") = False Then
                    UpdateWikiCache()
                End If
                FileOpen(2, JEGCore.WikiPath & "\Wiki.ini", OpenMode.Output, OpenAccess.Write, OpenShare.LockWrite)
                If bResult = True Then
                    Print(2, "Network")
                Else
                    Print(2, "Local")
                End If
                FileClose(2)
            End If
            sRunWikiNetwork = bResult
        ElseIf sRunWikiNetwork = "True" Then
            bResult = True
        Else
            bResult = True
        End If

        dsWiki = Nothing

        Return bResult
    End Function

    Shared Sub ViewWiki(ByVal sProgram As String)
        'View Wiki's relating to specified program
        Dim dsSoftware As DataRow
        Dim sWikiWebsite As String = GetConfigurationValue("MSTNWikiWebsite")

        'If AccessSharePoint() = True Then
        dsSoftware = SCEAccess.GetRowDB(JEGCore.SCECoreDB, "LUSoftware", "Software", "Software", "Internet Explorer", "", "", False, True)
        '    If sRunWikiNetwork = "True" Then
        Shell(dsSoftware.Item("Location") & Space(1) & Chr(34) & sWikiWebsite & "/" & sProgram & ".aspx" & Chr(34), AppWinStyle.NormalFocus)
        '    Else
        '        frmHelp.webHelp.Navigate("file:///" & sWikiPath & "\" & sProgram & ".html")
        '    End If
        '    dsSoftware = Nothing
        'End If

        'Dim RetVal
        'Dim sDocPath As String
        'Dim sWikiWebsite As String = XML.GetConfigurationValue("MSTNWikiWebsite")

        'If AccessSharePoint() = True Then
        '    If sWikiWebsite <> "" Then
        '        sDocPath = sWikiWebsite & "/" & sProgram & ".aspx"
        '        My.Computer.Registry.SetValue(My.Settings.cSKMCADKey & "\MiscTools\SCE Documentation", "Documentation", sDocPath)
        '        My.Computer.Registry.SetValue(My.Settings.cSKMCADKey & "\MiscTools\SCE Documentation", "Title", "MicroStation SCE Wiki")
        '        RetVal = Shell(Chr(34) & XML.GetConfigurationValue("MSTNSCERoot") & "\Wiki Viewer.exe" & Chr(34), vbNormalFocus)
        '    End If
        'End If
    End Sub

    Shared Function AccessSharePoint() As Boolean
        'Return boolean if SharePoint server is accessible
        Dim bReturn As Boolean = False

        Try
            If My.Computer.Network.IsAvailable = True Then
                Return My.Computer.Network.Ping(GetConfigurationValue("SharePointServer"), 200)
            End If
        Catch
        End Try

        Return bReturn
    End Function

    Shared Sub UpdateWikiCache()
        'Update local wiki cache
        Dim dsWiki As DataSet
        Dim sWiki As String
        Dim sFields As String = ""
        Dim dsFields As DataSet = Nothing
        Dim sValues As String = ""
        Dim i As Integer
        Dim sWikiContent As String

        'If FolderExists(sWikiPath) = True Then
        '    DeleteDir(sWikiPath)
        'End If
        My.Computer.FileSystem.CreateDirectory(JEGCore.WikiPath & "\Images")
        UpdateWikiImages()
        'If Wiki.IsWikisAvailable = False Then GoTo ExitSub
        dsWiki = GetAllDB(JEGCore.WikiDatabase, "MicroStation Wiki", "ID")

        'frmCatalogue.TSStatusLabel.Text = "Please wait, updating Wiki data cache..."
        'Application.DoEvents()
        'frmCatalogue.Cursor = Cursors.WaitCursor
        'frmCatalogue.TSProgressBar.Visible = True
        'frmCatalogue.TSProgressBar.ProgressBar.Maximum = dsWiki.Tables(0).Rows.Count
        'Create Database cache of Wiki's
        If TableExists(JEGCore.WikiDatabase, "MicroStation Wiki Cache") = True Then
            DeleteTable(JEGCore.WikiDatabase, "MicroStation Wiki Cache")
        End If

        sFields = "Name1|Keywords1|Modified1|Encoded Absolute URL"
        If TableExists(JEGCore.WikiDatabase, "MicroStation Wiki Cache") = False Then
            CreateTable(JEGCore.WikiDatabase, "MicroStation Wiki Cache", sFields)
        End If
        'Create local cache of Wiki's
        For i = 0 To dsWiki.Tables(0).Rows.Count - 1
            'frmCatalogue.TSProgressBar.ProgressBar.Value = i
            sValues = Replace(FileParse(dsWiki.Tables(0).Rows(i).Item("Name"), FileParser.FullFileName), "%20", " ")
            sValues = sValues & "|" & dsWiki.Tables(0).Rows(i).Item("Keywords")
            sValues = sValues & "|" & dsWiki.Tables(0).Rows(i).Item("Modified")
            sValues = sValues & "|" & dsWiki.Tables(0).Rows(i).Item("Encoded Absolute URL")
            SetDB(JEGCore.WikiDatabase, "MicroStation Wiki Cache", sFields, sValues)
            sWiki = Replace(FileParse(dsWiki.Tables(0).Rows(i).Item("Name"), FileParser.FullFileName), "%20", " ")
            sWiki = FileParse(sWiki, FileParser.FileName)
            sWiki = Replace(sWiki, "..", ".")
            FileOpen(2, JEGCore.WikiPath & "\" & sWiki & ".html", OpenMode.Output, OpenAccess.Write, OpenShare.LockWrite)
            PrintLine(2, "<p align=left><font face=Calibri size=5>" & sWiki & "</font></p>")
            sWikiContent = Replace(dsWiki.Tables(0).Rows(i).Item("Wiki Content"), "/cop/cad/Microstation/Images1/Wiki/", "Images/")
            'sWikiContent = Replace(sWikiContent, "/cop/cad/Microstation/MicroStation%20Wiki/", "\Images/")
            sWikiContent = Replace(sWikiContent, "/cop/cad/Microstation/MicroStation%20Wiki/Forms/", "")
            sWikiContent = Replace(sWikiContent, "aspx", "html")
            PrintLine(2, sWikiContent)
            FileClose(2)
        Next
        'frmCatalogue.TSStatusLabel.Text = ""
        'frmCatalogue.TSProgressBar.Visible = False
        'frmCatalogue.Cursor = Cursors.Default
ExitSub:
    End Sub

    Shared Sub UpdateTeamSiteWikiCache()
        'Update local wiki cache
        Dim dsWiki As DataSet
        Dim sWiki As String
        Dim sFields As String = ""
        Dim dsFields As DataSet = Nothing
        Dim sValues As String = ""
        Dim i As Integer
        Dim sWikiContent As String

        'If FolderExists(sWikiPath) = True Then
        '    DeleteDir(sWikiPath)
        'End If

        My.Computer.FileSystem.CreateDirectory(JEGCore.WikiPath & "\Images")
        ''UpdateWikiImages()

        'If Wiki.IsWikisAvailable = False Then GoTo ExitSub
        dsWiki = GetAllDB(JEGCore.WikiDatabase, "MicroStationWiki", "ID")

        'frmCatalogue.TSStatusLabel.Text = "Please wait, updating Wiki data cache..."
        'Application.DoEvents()
        'frmCatalogue.Cursor = Cursors.WaitCursor
        'frmCatalogue.TSProgressBar.Visible = True
        'frmCatalogue.TSProgressBar.ProgressBar.Maximum = dsWiki.Tables(0).Rows.Count
        'Create Database cache of Wiki's
        If TableExists(JEGCore.WikiDatabase, "MicroStation Wiki Cache") = True Then
            DeleteTable(JEGCore.WikiDatabase, "MicroStation Wiki Cache")
        End If

        sFields = "Name1|Keywords1|Modified1|Encoded Absolute URL"
        If TableExists(JEGCore.WikiDatabase, "MicroStation Wiki Cache") = False Then
            CreateTable(JEGCore.WikiDatabase, "MicroStation Wiki Cache", sFields)
        End If
        'Create local cache of Wiki's
        For i = 0 To dsWiki.Tables(0).Rows.Count - 1
            'frmCatalogue.TSProgressBar.ProgressBar.Value = i
            sValues = Replace(FileParse(dsWiki.Tables(0).Rows(i).Item("Name"), FileParser.FullFileName), "%20", " ")
            sValues = sValues & "|" & dsWiki.Tables(0).Rows(i).Item("Keywords")
            sValues = sValues & "|" & dsWiki.Tables(0).Rows(i).Item("Modified")
            sValues = sValues & "|" & dsWiki.Tables(0).Rows(i).Item("Encoded Absolute URL")
            SetDB(JEGCore.WikiDatabase, "MicroStation Wiki Cache", sFields, sValues)
            sWiki = Replace(FileParse(dsWiki.Tables(0).Rows(i).Item("Name"), FileParser.FullFileName), "%20", " ")
            sWiki = FileParse(sWiki, FileParser.FileName)
            sWiki = Replace(sWiki, "..", ".")
            FileOpen(2, JEGCore.WikiPath & "\" & sWiki & ".html", OpenMode.Output, OpenAccess.Write, OpenShare.LockWrite)
            PrintLine(2, "<p align=left><font face=Calibri size=5>" & sWiki & "</font></p>")
            sWikiContent = Replace(dsWiki.Tables(0).Rows(i).Item("Wiki Content"), "/cop/cad/Microstation/Images1/Wiki/", "Images/")
            'sWikiContent = Replace(sWikiContent, "/cop/cad/Microstation/MicroStation%20Wiki/", "\Images/")
            sWikiContent = Replace(sWikiContent, "/cop/cad/Microstation/MicroStation%20Wiki/Forms/", "")
            sWikiContent = Replace(sWikiContent, "aspx", "html")
            PrintLine(2, sWikiContent)
            FileClose(2)
        Next
        'frmCatalogue.TSStatusLabel.Text = ""
        'frmCatalogue.TSProgressBar.Visible = False
        'frmCatalogue.Cursor = Cursors.Default
ExitSub:
    End Sub

    '    Shared Sub UpdateDWSWikiCache()
    '        'Update local DWS wiki cache
    '        Dim dsWiki As DataSet
    '        Dim sWiki As String
    '        Dim sFields As String = ""
    '        Dim dsFields As DataSet = Nothing
    '        Dim sValues As String = ""
    '        Dim i As Integer
    '        Dim sWikiContent As String

    '        'If FolderExists(sWikiPath) = True Then
    '        '    DeleteDir(sWikiPath)
    '        'End If
    '        My.Computer.FileSystem.CreateDirectory(SCECore.ToolsPath() & "\SCE Documentation\DWS Wiki\Images")
    '        UpdateDWSWikiImages()
    '        'If Wiki.IsWikisAvailable = False Then GoTo ExitSub
    '        dsWiki = GetAllDB(SCECore.ToolsPath() & "\SCE Documentation\DWS Wiki\DWS Wiki.accdb", "DWS Wiki", "ID")

    '        'frmCatalogue.TSStatusLabel.Text = "Please wait, updating Wiki data cache..."
    '        'Application.DoEvents()
    '        'frmCatalogue.Cursor = Cursors.WaitCursor
    '        'frmCatalogue.TSProgressBar.Visible = True
    '        'frmCatalogue.TSProgressBar.ProgressBar.Maximum = dsWiki.Tables(0).Rows.Count
    '        'Create Database cache of Wiki's
    '        If TableExists(SCECore.ToolsPath() & "\SCE Documentation\DWS Wiki\DWS Wiki.accdb", "MicroStation Wiki Cache") = True Then
    '            DeleteTable(SCECore.ToolsPath() & "\SCE Documentation\DWS Wiki\DWS Wiki.accdb", "MicroStation Wiki Cache")
    '        End If

    '        sFields = "Name1|Keywords1|Modified1|Encoded Absolute URL"
    '        If TableExists(SCECore.ToolsPath() & "\SCE Documentation\DWS Wiki\DWS Wiki.accdb", "MicroStation Wiki Cache") = False Then
    '            CreateTable(SCECore.ToolsPath() & "\SCE Documentation\DWS Wiki\DWS Wiki.accdb", "MicroStation Wiki Cache", sFields)
    '        End If
    '        'Create local cache of Wiki's
    '        For i = 0 To dsWiki.Tables(0).Rows.Count - 1
    '            'frmCatalogue.TSProgressBar.ProgressBar.Value = i
    '            sValues = Replace(FileParse(dsWiki.Tables(0).Rows(i).Item("Name"), FileParser.FullFileName), "%20", " ")
    '            sValues = sValues & "|" & dsWiki.Tables(0).Rows(i).Item("Keywords")
    '            sValues = sValues & "|" & dsWiki.Tables(0).Rows(i).Item("Modified")
    '            sValues = sValues & "|" & dsWiki.Tables(0).Rows(i).Item("Encoded Absolute URL")
    '            SetDB(SCECore.ToolsPath() & "\SCE Documentation\DWS Wiki\DWS Wiki.accdb", "MicroStation Wiki Cache", sFields, sValues)
    '            sWiki = Replace(FileParse(dsWiki.Tables(0).Rows(i).Item("Name"), FileParser.FullFileName), "%20", " ")
    '            sWiki = Replace(sWiki, "%27", "'")
    '            sWiki = FileParse(sWiki, FileParser.FileName)
    '            sWiki = Replace(sWiki, "..", ".")
    '            FileOpen(2, SCECore.ToolsPath() & "\SCE Documentation\DWS Wiki\" & sWiki & ".html", OpenMode.Output, OpenAccess.Write, OpenShare.LockWrite)
    '            PrintLine(2, "<p align=left><font face=Calibri size=5>" & sWiki & "</font></p>")
    '            sWikiContent = Replace(dsWiki.Tables(0).Rows(i).Item("Wiki Content"), "/cop/cad/Microstation/Images1/DWS%20Wiki/", "Images/")
    '            'sWikiContent = Replace(sWikiContent, "/cop/cad/Microstation/MicroStation%20Wiki/", "\Images/")
    '            sWikiContent = Replace(sWikiContent, "/cop/cad/Microstation/DWS%20Wiki/Forms/", "")
    '            sWikiContent = Replace(sWikiContent, "aspx", "html")
    '            PrintLine(2, sWikiContent)
    '            FileClose(2)
    '        Next
    '        'frmCatalogue.TSStatusLabel.Text = ""
    '        'frmCatalogue.TSProgressBar.Visible = False
    '        'frmCatalogue.Cursor = Cursors.Default
    'ExitSub:
    '        MsgBox("DWS Wiki Cache Update is now complete.", MsgBoxStyle.OkOnly, My.Settings.MsgBoxCaption)
    '    End Sub

    Shared Sub UpdateWikiImages()
        'Update Wiki Images
        ''Dim aFiles As Array
        ''Dim i As Integer
        Dim sFile As String

        'If Wiki.IsWikisAvailable = False Then GoTo ExitSub

        'frmCatalogue.TSStatusLabel.Text = "Please wait, updating Wiki image cache..."
        'Application.DoEvents()
        'frmCatalogue.Cursor = Cursors.WaitCursor
        'frmCatalogue.TSProgressBar.Visible = True
        'frmCatalogue.TSProgressBar.ProgressBar.Maximum = UBound(aFiles)
        For Each sFile In My.Computer.FileSystem.GetFiles("\\" & GetConfigurationValue("SharePointServer") & "\cop\cad\microstation\images1\wiki", FileIO.SearchOption.SearchTopLevelOnly, "*")
            'frmCatalogue.TSProgressBar.Value = i
            If My.Computer.FileSystem.FileExists("\\" & GetConfigurationValue("SharePointServer") & "\cop\cad\microstation\images1\wiki\" & sFile) = True Then
                If My.Computer.FileSystem.FileExists(JEGCore.WikiPath & "\Images\" & sFile) = True Then
                    If FileDateTime("\\" & GetConfigurationValue("SharePointServer") & "\cop\cad\microstation\images1\wiki\" & sFile) > FileDateTime(JEGCore.WikiPath & "\Images\" & sFile) Then
                        FileCopy("\\" & GetConfigurationValue("SharePointServer") & "\cop\cad\microstation\images1\wiki\" & sFile, JEGCore.WikiPath & "\Images\" & sFile)
                    End If
                End If
            End If
        Next
        'frmCatalogue.TSStatusLabel.Text = "Please wait, updating Wiki image cache..."
        'frmCatalogue.TSProgressBar.Visible = False
        'frmCatalogue.Cursor = Cursors.Default
ExitSub:
    End Sub

    ''    Shared Sub UpdateDWSWikiImages()
    ''        'Update DWS Wiki Images
    ''        Dim aFiles As Array
    ''        Dim i As Integer

    ''        'If Wiki.IsWikisAvailable = False Then GoTo ExitSub

    ''        'frmCatalogue.TSStatusLabel.Text = "Please wait, updating Wiki image cache..."
    ''        'Application.DoEvents()
    ''        'frmCatalogue.Cursor = Cursors.WaitCursor
    ''        aFiles = GetFiles("\\" & GetConfigurationValue("SharePointServer") & "\cop\cad\microstation\images1\dws wiki", "*")
    ''        'frmCatalogue.TSProgressBar.Visible = True
    ''        'frmCatalogue.TSProgressBar.ProgressBar.Maximum = UBound(aFiles)
    ''        For i = 0 To UBound(aFiles)
    ''            'frmCatalogue.TSProgressBar.Value = i
    ''            If My.Computer.FileSystem.FileExists("\\" & GetConfigurationValue("SharePointServer") & "\cop\cad\microstation\images1\dws wiki\" & aFiles(i)) = True Then
    ''                If My.Computer.FileSystem.FileExists(SCECore.ToolsPath() & "\SCE Documentation\DWS Wiki\Images\" & aFiles(i)) = True Then
    ''                    If FileDateTime("\\" & GetConfigurationValue("SharePointServer") & "\cop\cad\microstation\images1\dws wiki\" & aFiles(i)) > FileDateTime(SCECore.ToolsPath() & "\SCE Documentation\DWS Wiki\Images\" & aFiles(i)) Then
    ''                        FileCopy("\\" & GetConfigurationValue("SharePointServer") & "\cop\cad\microstation\images1\dws wiki\" & aFiles(i), SCECore.ToolsPath() & "\SCE Documentation\DWS Wiki\Images\" & aFiles(i))
    ''                    End If
    ''                Else
    ''                    FileCopy("\\" & GetConfigurationValue("SharePointServer") & "\cop\cad\microstation\images1\dws wiki\" & aFiles(i), SCECore.ToolsPath() & "\SCE Documentation\DWS Wiki\Images\" & aFiles(i))
    ''                End If
    ''            End If
    ''        Next
    ''        'frmCatalogue.TSStatusLabel.Text = "Please wait, updating Wiki image cache..."
    ''        'frmCatalogue.TSProgressBar.Visible = False
    ''        'frmCatalogue.Cursor = Cursors.Default
    ''ExitSub:
    ''    End Sub

    'Shared Function HasWiki(ByVal sProgram As String) As Boolean
    '    'Return if specified program has current wiki's
    '    'Author: Paul Ripp
    '    Dim dsWiki As DataSet = Nothing

    '    If AccessSharePoint() = True Then
    '        dsWiki = GetColumnDB(sUserDB, "MicroStation Wiki", "Keywords", "Keywords", "Program:" & sProgram, True, False, False)
    '        If dsWiki.Tables(0).Rows.Count = 0 Then

    '            Return False
    '        Else
    '            Return True
    '        End If
    '    End If
    'End Function

    'Shared Sub ViewWiki(ByVal sProgram As String)
    '    'View Wiki's relating to specified program
    '    'Author: Paul Ripp
    '    Dim dsSoftware As DataSet
    '    Dim sWikiWebsite As String = xml.GetConfigurationValue("MSTNWikiWebsite")

    '    If AccessSharePoint() = True Then
    '        dsSoftware = Access.GetRowDB(sCoreDB, "LUSoftware", "Software", "Software", "Internet Explorer", "", "", False, True)
    '        Shell(dsSoftware.Tables(0).Rows(0).Item("Location") & Space(1) & Chr(34) & sWikiWebsite & "/" & sProgram & ".aspx" & Chr(34), AppWinStyle.NormalFocus)
    '        dsSoftware = Nothing
    '    End If
    'End Sub

    'Shared Function AccessSharePoint() As Boolean
    '    'Return boolean if SharePoint server is accessible
    '    'Author: Paul Ripp
    '    If My.Computer.Network.IsAvailable = True Then
    '        Return My.Computer.Network.Ping(GetConfigurationValue("SharePointServer"), 200)
    '    Else
    '        Return False
    '    End If
    'End Function

End Class


