Imports MicroStationSCE.SCESoftware

<ComClass(MSTNApplications.ClassId, MSTNApplications.InterfaceId, MSTNApplications.EventsId)> _
Public Class MSTNApplications

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "0d08bcc3-d762-4dee-9f60-8b72560b4e06"
    Public Const InterfaceId As String = "2b047fc7-d153-403d-b1da-5fb41ea8914f"
    Public Const EventsId As String = "78a3405c-0299-4356-b905-431febe2c553"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Private Shared _myInstance As MicroStationDGN.Application
    Public Shared oMSTN As MicroStationDGN.Application = GetInstance()

    Public Shared Function GetInstance() As MicroStationDGN.Application
        If _myInstance Is Nothing Then
            _myInstance = New MicroStationDGN.Application 'TODO: Add check if App is already(running)
        End If
        Return _myInstance
    End Function

    Sub ExploreBackupPath()
        'Open Windows Explorer from Backup location

        If My.Computer.FileSystem.DirectoryExists(oMSTN.ActiveWorkspace.ConfigurationVariableValue("MS_BACKUP")) = True Then
            Shell("explorer " & oMSTN.ActiveWorkspace.ConfigurationVariableValue("MS_BACKUP"), vbNormalFocus)
        Else
            MsgBox("Backup folder " & Chr(34) & oMSTN.ActiveWorkspace.ConfigurationVariableValue("MS_BACKUP") & Chr(34) & " does not exist", vbExclamation, "SCE - Backup Drawing")
        End If
    End Sub

    Sub ExploreDrawingPath()
        'Open Windows Explorer from DGN file location
        Shell("explorer " & oMSTN.ActiveDesignFile.Path, vbNormalFocus)
    End Sub

    Sub Convertor()
        'Open convertor program
        Dim sSCESupport As String

        sSCESupport = MSTNXML.GetXMLSetting(oMSTN.ActiveWorkspace.ConfigurationVariableValue("_SCE_XMLCONFIGFILE"), "SCERoot")
        sSCESupport = Replace(UCase(sSCESupport), "[PROGRAMFILES]", "C:\Program Files")

        If My.Computer.FileSystem.FileExists(sSCESupport & "\Support\convertor.exe") = True Then
            Shell(sSCESupport & "\Support\convertor.exe", vbNormalFocus)
        End If
    End Sub

    Sub SKMStandardDrawings()
        'Open convertor program
        oMSTN.CadInputQueue.SendCommand("vba run [SCECore]Applications.Web " & Chr(34) & "http://intranet.skmconsulting.com/FunctionalUnits/BusinessProcesses/CAD/skm_cad_std_drgs.htm" & Chr(34))
    End Sub

    Sub MicroStationCoP()
        'Open MicroStation Community of Practise
        oMSTN.CadInputQueue.SendCommand("vba run [SCECore]Applications.Web " & Chr(34) & "http://km.skmconsulting.com/cop/cad/Microstation/Lists/Team%20Discussion/AllItems.aspx" & Chr(34))
    End Sub

    Sub MicroStationWiki()
        'Open MicroStation Wiki's
        If My.Computer.FileSystem.FileExists(oMSTN.ActiveWorkspace.ConfigurationVariableValue("SCE_ROOT") & "Wiki Viewer.exe") = True Then
            Shell(oMSTN.ActiveWorkspace.ConfigurationVariableValue("SCE_ROOT") & "Wiki Viewer.exe", vbNormalFocus)
        Else
            If SCEWiki.AccessSharePoint = True Then
                oMSTN.CadInputQueue.SendCommand("vba run [SCECore]Applications.Web " & Chr(34) & "http://km.skmconsulting.com/cop/cad/Microstation/MicroStation%20Wiki/Forms/AllPages.aspx" & Chr(34))
            End If
        End If
    End Sub

    Sub ExploreIssuedDrawingPath()
        'Open Windows Explorer from Issued Drawing file location
        Dim sIssuePath As String = ""
        Dim vFolders As Array
        Dim i As Long
        Dim j As Long

        If InStr(UCase(oMSTN.ActiveDesignFile.Path), "DELIVERABLES\DRAWINGS") > 0 Then
            sIssuePath = ""
            vFolders = Split(oMSTN.ActiveDesignFile.Path, "\")

            For i = 0 To UBound(vFolders)
                If UCase(vFolders(i)) = "DRAWINGS" Then
                    'Build path
                    For j = 0 To i
                        sIssuePath = sIssuePath & vFolders(j) & "\"
                    Next
                    If UBound(vFolders) = (i + 1) Then
                        sIssuePath = sIssuePath & "Issued\Drawings\" & vFolders(i + 1) 'Area/Discipline
                    ElseIf UBound(vFolders) = (i + 2) Then
                        sIssuePath = sIssuePath & "Issued\Drawings\" & vFolders(i + 1) & "\" & vFolders(i + 2) 'Area and Discipline
                    Else 'something else
                        sIssuePath = sIssuePath & "Issued\Drawings"
                    End If
                    Exit For
                End If
            Next
        End If

        If My.Computer.FileSystem.DirectoryExists(sIssuePath) = True Then
            Shell("explorer " & sIssuePath, vbNormalFocus)
        Else
            MsgBox("Issue path " & Chr(34) & sIssuePath & Chr(34) & " does not exist", vbExclamation, "SCE - Open Explorer")
        End If
    End Sub

    Sub ExploreIssuedPDFPath()
        'Open Windows Explorer from Issued PDF file location
        Dim sIssuePath As String
        Dim vFolders As Object
        Dim i As Long
        Dim j As Long

        If InStr(UCase(oMSTN.ActiveDesignFile.Path), "DELIVERABLES\DRAWINGS") > 0 Then
            sIssuePath = ""
            vFolders = Split(oMSTN.ActiveDesignFile.Path, "\")
            For i = 0 To UBound(vFolders)
                If UCase(vFolders(i)) = "DRAWINGS" Then
                    'Build path
                    For j = 0 To i
                        sIssuePath = sIssuePath & vFolders(j) & "\"
                    Next
                    If UBound(vFolders) = (i + 1) Then
                        sIssuePath = sIssuePath & "Issued\PDF\" & vFolders(i + 1) 'Area/Discipline
                    ElseIf UBound(vFolders) = (i + 2) Then
                        sIssuePath = sIssuePath & "Issued\PDF\" & vFolders(i + 1) & "\" & vFolders(i + 2) 'Area and Discipline
                    Else 'something else
                        sIssuePath = sIssuePath & "Issued\PDF"
                    End If
                    Exit For
                End If
            Next
        End If

        If My.Computer.FileSystem.DirectoryExists(sIssuePath) = True Then
            Shell("explorer " & sIssuePath, vbNormalFocus)
        Else
            MsgBox("Issue path " & Chr(34) & sIssuePath & Chr(34) & " does not exist", vbExclamation, "SCE - Open Explorer")
        End If
    End Sub

    Sub Web()
        'Open webpage in IE
        'Keyin: vba run Web http://intranet.skm.com.au
        Dim sFullName As String
        Dim WebSite As String

        WebSite = oMSTN.KeyinArguments
        sFullName = SCESoftware.FullName("Internet Explorer")
        'Open Internet Explorer and access the inputed website
        If My.Computer.FileSystem.FileExists(sFullName) = True Then
            Shell(sFullName & " " & WebSite, vbNormalFocus)
        End If
    End Sub

    Sub OpenPDF()
        'Launch Adobe Acrobat Reader
        Dim sFullName As String
        Dim sPDFFullName As String

        If sPDFFullName = "" Then sPDFFullName = oMSTN.KeyinArguments
        sPDFFullName = oMSTN.ActiveWorkspace.ExpandConfigurationVariable(sPDFFullName)
        If My.Computer.FileSystem.FileExists(sPDFFullName) = True Then
            sFullName = SCESoftware.FullName("Acrobat Reader")
            If My.Computer.FileSystem.FileExists(sFullName) = True Then
                Shell(sFullName & Space(1) & Chr(34) & sPDFFullName & Chr(34), vbNormalFocus)
            Else
                oMSTN.ShowMessage("Adobe Acrobat Reader was not found.", sFullName)
            End If
        Else
            oMSTN.ShowMessage("PDF does not exist.", sPDFFullName)
        End If
    End Sub

    Sub ClientDocumentation()
        'Open Client documentation program
        'Dim Reg As New RegistryFunctions
        Dim RetVal
        Dim sDocPath As String

        sDocPath = oMSTN.ActiveWorkspace.ConfigurationVariableValue("SCE_CLIENT_DOCUMENTATION")
        If Right(sDocPath, 1) = "\" Then
            sDocPath = Left(sDocPath, Len(sDocPath) - 1)
            sDocPath = Replace(sDocPath, "/", "\")
        End If
        If Reg.KeyExist(cLocalSKMCADKey & "\MiscTools\SCE Documentation") = False Then
            Reg.CreateKey(cLocalSKMCADKey & "\MiscTools\SCE Documentation")
        End If
        Reg.SetStringValue(cLocalSKMCADKey & "\MiscTools\SCE Documentation", "Documentation", sDocPath)
        Reg.SetStringValue(cLocalSKMCADKey & "\MiscTools\SCE Documentation", "Title", ClientSettings.FullName & Space(1) & "Documentation")
        RetVal = Shell(Chr(34) & oMSTN.ActiveWorkspace.ConfigurationVariableValue("SCE_ROOT") & "SCE Documentation.exe" & Chr(34), vbNormalFocus)
    End Sub

    Sub CellCatalogue()
        'Open the Cell Catalogue

        Shell(oMSTN.ActiveWorkspace.ConfigurationVariableValue("SCE_ROOT") & "cellcatalogue.exe", vbNormalFocus)
    End Sub

    Sub CellCataloguePlaceCell()
        Dim vSplit

        'Show placement dialog
        'frmPlaceCell.Show

        vSplit = Split(oMSTN.KeyinArguments, "|")
        oMSTN.ActiveWorkspace.AddConfigurationVariable("_TMP_CELLLIB", vSplit(0))
        oMSTN.ActiveWorkspace.AddConfigurationVariable("_TMP_CELLNAME", vSplit(1))
        oMSTN.ActiveWorkspace.AddConfigurationVariable("_TMP_CELLLEVELOVERRIDE", vSplit(2))
        oMSTN.ActiveWorkspace.AddConfigurationVariable("_TMP_CELLLEVEL", vSplit(3))

        CommandState.StartPrimitive(New AutoPlaceCell)
    End Sub


End Class


