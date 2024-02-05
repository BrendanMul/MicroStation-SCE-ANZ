Public Class SCEPrint

    Public Shared dsPrint As DataSet
    Public Shared Loaded As Boolean
    Public Shared GeneralCount As Integer
    Public Shared PlotterCount As Integer

    Shared Sub Initialise()
        'Load print settings
        Dim oTable As DataTable
        Dim dsTempPrint As DataSet

        If Loaded = True Then
            GoTo ExitSub
        End If

        'General print information
        dsTempPrint = SCEAccess.GetAllDB(JEGCore.SCECoreDB, "LUPrint", "PrintType")

        If SCEVB.IsNullDataSet(dsTempPrint) = True Then
            GoTo ExitSub
        End If

        dsPrint = New DataSet
        oTable = dsTempPrint.Tables(0)
        oTable.TableName = "General"
        dsPrint.Tables.Add(oTable.Clone)
        GeneralCount = oTable.Rows.Count

        'Plotter information
        'dsTempPrint = SCEAccess.GetAllDB(SCECore.SCECoreDB, "LUPlotter", "Description")

        'If SCEVB.IsNullDataSet(dsTempPrint) = True Then
        '    GoTo ExitSub
        'End If

        'oTable = dsTempPrint.Tables(0)
        'oTable.TableName = "Plotters"
        'dsPrint.Tables.Add(oTable.Clone)
        'PlotterCount = oTable.Rows.Count

        'status issue settings are added when client is loaded

        dsTempPrint.Dispose()
        Loaded = True
ExitSub:
    End Sub

    Shared Function GetGeneralIndex(ByVal sOffice As String) As Integer
        'get index of specified print type
        Dim iReturn As Integer = -1
        Dim i As Integer

        Initialise()

        For i = 0 To GeneralCount - 1
            If UCase(sOffice) = UCase(dsPrint.Tables(0).Rows(i).Item("Office")) Then
                iReturn = i
                Exit For
            End If
        Next

        Return iReturn
    End Function

    Shared Function GetPlotterIndex(ByVal sOffice As String) As Integer
        'get index of specified print type
        Dim iReturn As Integer = -1
        Dim i As Integer

        Initialise()

        For i = 0 To GeneralCount - 1
            If UCase(sOffice) = UCase(dsPrint.Tables(1).Rows(i).Item("Office")) Then
                iReturn = i
                Exit For
            End If
        Next

        Return iReturn
    End Function

    'Public Shared Sub SilentCreateThumbnailJPGs(oMSTN As MicroStationDGN.Application, oDesign As DesignFile, ByVal sOutputPath As String)
    '    'Dim oDesign As DesignFile
    '    Dim oModel As ModelReference

    '    'oMSTN.Visible = False
    '    'oDesign = oMSTN.OpenDesignFile(sFileName, True)

    '    For Each oModel In oDesign.Models
    '        If oModel.Type = MsdModelType.msdModelTypeDefault Or oModel.Type = MsdModelType.msdModelTypeDrawing Or oModel.Type = MsdModelType.msdModelTypeNormal Or oModel.Type = MsdModelType.msdModelTypeSheet Then
    '            oModel.Activate()
    '            oMSTN.CadInputQueue.SendCommand("FIT ALL;SELVIEW 1")
    '            oMSTN.CadInputQueue.SendCommand("print driver C:\Program Files (x86)\Jacobs\MicroStation\MicroStation Core\Workspace\PlotDrv\CellCatalogueJPG.plt")
    '            oMSTN.CadInputQueue.SendCommand("print papername A4")
    '            oMSTN.CadInputQueue.SendCommand("print fullsheet on")
    '            oMSTN.CadInputQueue.SendCommand("print attributes datafields on")
    '            oMSTN.CadInputQueue.SendCommand("print attributes fill on")
    '            oMSTN.CadInputQueue.SendCommand("print attributes transparency on")
    '            oMSTN.CadInputQueue.SendCommand("print boundary all")
    '            oMSTN.CadInputQueue.SendCommand("print execute " & sOutputPath & "\" & oModel.Name & ".jpg")
    '        End If
    '    Next
    '    'oDesign.Close()
    '    'oMSTN.Quit
    'End Sub

    Public Shared Sub SilentCreateThumbnailJPG(ByVal sFilename As String, Optional ByVal sOutputPath As String = "")
        '        If My.Computer.FileSystem.FileExists(sFilename) = False Then
        '            GoTo ExitSub
        '        End If

        '        Dim oDesign As MicroStationDGN.DesignFile
        '        Dim oMSTN As New MicroStationDGN.Application
        '        'Dim oModel As MicroStationDGN.ModelReference

        '        If sOutputPath = "" Then
        '            sOutputPath = SCEFile.FileParse(sFilename, SCEFile.FileParser.Path)
        '        End If

        '        oMSTN.Visible = False
        '        oDesign = oMSTN.OpenDesignFile(sFilename, True) 'oMSTN.OpenDesignFile("C:\Users\pripp\Documents\vicroads.dgn", True)

        '        'For Each oModel In oDesign.Models
        '        '    If oModel.Type = MicroStationDGN.MsdModelType.msdModelTypeDefault Or oModel.Type = MicroStationDGN.MsdModelType.msdModelTypeDrawing Or oModel.Type = MicroStationDGN.MsdModelType.msdModelTypeNormal Or oModel.Type = MicroStationDGN.MsdModelType.msdModelTypeSheet Then
        '        '        oModel.Activate()

        '        'oMSTN.CadInputQueue.SendCommand("print driver C:\Program Files (x86)\SKMCAD\MicroStation SCE\MicroStation Core\Workspace\PlotDrv\CellCatalogueJPG.plt")
        '        oMSTN.ActiveWorkspace.AddConfigurationVariable("_SKM_PDF", sOutputPath & "\" & SCEFile.FileParse(sFilename, SCEFile.FileParser.FileName))
        '        oMSTN.CadInputQueue.SendCommand("print driver C:\ProgramData\Jacobs\MicroStation\Working\WA\RioTintoExpansionProjects\Workspace\plotdrv\PDF_Mech.plt")
        '        'oMSTN.CadInputQueue.SendCommand("print execute " & sOutputPath & "\" & oModel.Name & ".jpg")
        '        oMSTN.CadInputQueue.SendCommand("print papername b1")
        '        oMSTN.CadInputQueue.SendKeyin("PRINT BOUNDARY FIT ALL")
        '        oMSTN.CadInputQueue.SendCommand("print fullsheet on")
        '        oMSTN.CadInputQueue.SendCommand("print attributes datafields on")
        '        oMSTN.CadInputQueue.SendCommand("print attributes fill on")
        '        oMSTN.CadInputQueue.SendCommand("print attributes transparency on")
        '        oMSTN.CadInputQueue.SendCommand("print attributes textnodes off")
        '        oMSTN.CadInputQueue.SendKeyin("PRINT MAXIMIZE")
        '        oMSTN.CadInputQueue.SendCommand("print execute")
        '        '    End If
        '        'Next
        '        oDesign.Close()
        '        oMSTN.Quit()
        'ExitSub:
    End Sub

    'Sub CreateSilentModelList()
    '    Dim oDesign As DesignFile
    '    Dim oMSTN As New MicroStationDGN.Application
    '    Dim oModel As ModelReference
    '    Dim oAttach As Attachment
    '    Dim oAttachNest As Attachment
    '    Dim oAttachNest1 As Attachment

    '    oMSTN.Visible = False
    '    oDesign = oMSTN.OpenDesignFile("C:\Users\pripp\Documents\vicroads.dgn", True)

    '    For Each oModel In oDesign.Models
    '        If oModel.Type = msdModelTypeDefault Or oModel.Type = msdModelTypeDrawing Or oModel.Type = msdModelTypeNormal Or oModel.Type = msdModelTypeSheet Then
    '            For Each oAttach In oModel.Attachments
    '                For Each oAttachNest In oAttach.Attachments 'Nest level 1
    '                    For Each oAttachNest1 In oAttachNest.Attachments 'Next level 2

    '                    Next
    '                Next
    '            Next
    '        End If
    '    Next
    '    oDesign.Close()
    '    'oMSTN.Quit
    'End Sub
End Class
