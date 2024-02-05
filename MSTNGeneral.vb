Imports SCECoreDLL

<ComClass(MSTNGeneral.ClassId, MSTNGeneral.InterfaceId, MSTNGeneral.EventsId)> _
Public Class MSTNGeneral

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "5e7fd657-c8d1-48f3-a456-2241eef6b4c7"
    Public Const InterfaceId As String = "43f4ca21-84c0-4ec3-8159-40a33e01e67a"
    Public Const EventsId As String = "9a6971be-eb77-43c2-b51a-6e741894b6fd"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Private Shared _myInstance As MicroStationDGN.Application
    Public Shared oMSTN As MicroStationDGN.Application = SCEMicroStationSession.GetInstance()

    'Public Shared Function GetInstance() As MicroStationDGN.Application
    '    If _myInstance Is Nothing Then
    '        _myInstance = New MicroStationDGN.Application 'TODO: Add check if App is already(running)
    '    End If
    '    Return _myInstance
    'End Function

    Sub ExploreBackupPath()
        'Open Windows Explorer from Backup location
        SCESoftware.ExplorePath(oMSTN.ActiveWorkspace.ConfigurationVariableValue("MS_BACKUP"))
    End Sub

    Sub ExploreDrawingPath()
        'Open Windows Explorer from DGN file location
        SCESoftware.ExplorePath(oMSTN.ActiveDesignFile.Path)
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

    'Sub ClientDocumentation()
    '    'Open Client documentation program
    '    Dim Reg As New RegistryFunctions
    '    Dim RetVal
    '    Dim sDocPath As String

    '    sDocPath = oMSTN.ActiveWorkspace.ConfigurationVariableValue("SCE_CLIENT_DOCUMENTATION")
    '    If Right(sDocPath, 1) = "\" Then
    '        sDocPath = Left(sDocPath, Len(sDocPath) - 1)
    '        sDocPath = Replace(sDocPath, "/", "\")
    '    End If
    '    If Reg.KeyExist(My.Settings.LocalSKMCADKey & "\MiscTools\SCE Documentation") = False Then
    '        Reg.CreateKey(My.Settings.LocalSKMCADKey & "\MiscTools\SCE Documentation")
    '    End If
    '    Reg.SetStringValue(My.Settings.LocalSKMCADKey & "\MiscTools\SCE Documentation", "Documentation", sDocPath)
    '    Reg.SetStringValue(My.Settings.LocalSKMCADKey & "\MiscTools\SCE Documentation", "Title", ClientSettings.FullName & Space(1) & "Documentation")
    '    RetVal = Shell(Chr(34) & oMSTN.ActiveWorkspace.ConfigurationVariableValue("SCE_ROOT") & "SCE Documentation.exe" & Chr(34), vbNormalFocus)
    'End Sub

    Sub CellCataloguePlaceCell()
        Dim vSplit

        vSplit = Split(oMSTN.KeyinArguments, "|")
        oMSTN.ActiveWorkspace.AddConfigurationVariable("_TMP_CELLLIB", vSplit(0))
        oMSTN.ActiveWorkspace.AddConfigurationVariable("_TMP_CELLNAME", vSplit(1))
        oMSTN.ActiveWorkspace.AddConfigurationVariable("_TMP_CELLLEVELOVERRIDE", vSplit(2))
        oMSTN.ActiveWorkspace.AddConfigurationVariable("_TMP_CELLLEVEL", vSplit(3))
        'pjrxxx - uncomment for final
        'oMSTN.CommandState.StartPrimitive(New AutoPlaceCell)
    End Sub

End Class


