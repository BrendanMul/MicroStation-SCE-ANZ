''Imports MicroStationSCE.SCEFile
''Imports MicroStationSCE.sceaccess

''<ComClass(MSTNCell.ClassId, MSTNCell.InterfaceId, MSTNCell.EventsId)> _
''Public Class MSTNCell

''#Region "COM GUIDs"
''    ' These  GUIDs provide the COM identity for this class 
''    ' and its COM interfaces. If you change them, existing 
''    ' clients will no longer be able to access the class.
''    Public Const ClassId As String = "3011d455-c6ab-4ea0-8fb6-e0db83b34649"
''    Public Const InterfaceId As String = "36af9644-53ac-4fab-b590-66c2fc93d1ee"
''    Public Const EventsId As String = "aa338cc0-0b08-49a0-aeea-fd6979666d78"
''#End Region

''    Public Shared dsCells As DataSet

''    Private Shared _myInstance As MicroStationDGN.Application
''    Public Shared oMSTN As MicroStationDGN.Application = GetInstance()

''    Public Shared Function GetInstance() As MicroStationDGN.Application
''        If _myInstance Is Nothing Then
''            _myInstance = New MicroStationDGN.Application 'TODO: Add check if App is already(running)
''        End If
''        Return _myInstance
''    End Function

''    ' A creatable COM class must have a Public Sub New() 
''    ' with no parameters, otherwise, the class will not be 
''    ' registered in the COM registry and cannot be created 
''    ' via CreateObject.
''    Public Sub New()
''        MyBase.New()
''    End Sub


''    '--------------------------------------------------------------------------------
''    '
''    '   Cell Tools
''    '
''    '   - Manipulate Cells within the drawing
''    '
''    '---------------------------------------------------------------------------------

''    Private Const cMsgBoxCaption = "SCE - Cell Tools"

''    Public Enum eCellTypes
''        tClientCells = 0
''        tAllClientCells = 1
''        tNetworkCells = 2
''        tProjectCells = 3
''        tUserCells = 4
''    End Enum

''Sub InsertCell(ByVal sCellName As String, pPoint3D As Point3d, Optional ByVal sCellScale As String)
''        'Places cell based on user input
''        Dim vValues As Array
''        Dim sCellAngle As String

''        ''CurrentScale = oMSTN.ActiveSettings.Scale.X

''        'If called from command line get values
''        If oMSTN.KeyinArguments <> "" Then
''            vValues = Split(oMSTN.KeyinArguments, ",")
''            If UBound(vValues) = 5 Then
''                sCellName = vValues(0)
''                pPoint3D.X = vValues(1)
''                pPoint3D.Y = vValues(2)
''                pPoint3D.Z = vValues(3)
''                sCellAngle = vValues(4)
''                sCellScale = vValues(5)
''            Else
''                MsgBox("Incorrect number of values", vbCritical, cMsgBoxCaption)
''                Exit Sub
''            End If
''        Else
''            If sCellScale <> "" Then
''                pPoint3D.X = pPoint3D.X * sCellScale
''                pPoint3D.Y = pPoint3D.Y * sCellScale
''                pPoint3D.Z = pPoint3D.Z * sCellScale
''            End If
''        End If
''        'Insert Cell
''        oMSTN.CadInputQueue.SendKeyin("ac=" & sCellName & ";as=" & sCellScale)
''        oMSTN.CadInputQueue.SendDataPoint(pPoint3D, 1)
''        oMSTN.CadInputQueue.SendKeyin("as=" & sCellScale)
''        oMSTN.CadInputQueue.SendKeyin(";nocommand")
''        oMSTN.CommandState.StartDefaultCommand()
''    End Sub

''Sub InsertCellPoints(Optional ByVal sCellName As String, Optional ByVal sX As String, Optional ByVal sY As String, Optional ByVal sZ As String, Optional ByVal sAngle As String)

''        'Places cell based on user input
''        'Keyin: vba run InsertCellPoints CellName,001,002,003,004
''        'Place cell at location 001,002,003 at angle 004
''        'must have 3 digits for X, Y, Z and angle
''        'sets scale to Active Scale

''        Dim point As Point3d
''        Dim vValues As Object

''        'If called from command line get values
''        If oMSTN.KeyinArguments <> "" Then

''            vValues = Split(oMSTN.KeyinArguments, ",")

''            If UBound(vValues) = 4 Then

''                sCellName = vValues(0)
''                sX = vValues(1)
''                sY = vValues(2)
''                sZ = vValues(3)
''                sAngle = vValues(4)

''            Else

''                MsgBox("Incorrect number of values", vbCritical, cMsgBoxCaption)
''                Exit Sub

''            End If

''        End If

''        point.X = sX * oMSTN.ActiveSettings.Scale.X
''        point.Y = sY * oMSTN.ActiveSettings.Scale.Y
''        point.Z = sZ

''        'Insert Cell
''        oMSTN.CadInputQueue.SendKeyin("ac=" & sCellName & ";aa=" & sAngle)
''        oMSTN.CadInputQueue.SendDataPoint(point)
''        oMSTN.CadInputQueue.SendKeyin(";nocommand")
''        oMSTN.CommandState.StartDefaultCommand()

''    End Sub

''Sub InsertCellPointsScale(Optional ByVal sCellName As String, Optional ByVal sX As String, Optional ByVal sY As String, Optional ByVal sZ As String, Optional ByVal sCellAngle As String, Optional ByVal sCellScale As String)
''        'Places cell based on user input
''        'Keyin: vba run InsertCellPoints CellName,001,002,003,004,0.5
''        'Place cell at location 001,002,003 at angle 004 and scale .5
''        'must have 3 characters for X, Y, Z, angle and scale (scale includes point)
''        Dim point As Point3d
''        Dim vValues As Object

''        ''CurrentScale = oMSTN.ActiveSettings.Scale.X

''        'If called from command line get values
''        If oMSTN.KeyinArguments <> "" Then
''            vValues = Split(oMSTN.KeyinArguments, ",")
''            If UBound(vValues) = 5 Then
''                sCellName = vValues(0)
''                point.X = vValues(1)
''                point.Y = vValues(2)
''                point.Z = vValues(3)
''                sCellAngle = vValues(4)
''                sCellScale = vValues(5)
''            Else
''                MsgBox("Incorrect number of values", vbCritical, cMsgBoxCaption)
''                Exit Sub
''            End If
''        Else
''            point.X = sX
''            point.Y = sY
''            point.Z = sZ
''        End If

''        'Insert Cell
''        oMSTN.CadInputQueue.SendKeyin("ac=" & sCellName & ";aa=" & sCellAngle & ";as=" & sCellScale)
''        oMSTN.CadInputQueue.SendDataPoint(point, 1)
''        oMSTN.CadInputQueue.SendKeyin("as=" & sCellScale)
''        oMSTN.CadInputQueue.SendKeyin(";nocommand")
''        oMSTN.CommandState.StartDefaultCommand()
''    End Sub

''Sub DeleteCell(Optional ByVal sCellName As String, Optional oException As CellElement)
''        'Searches drawing for cell and deletes all instances if found
''        'Keyin: vba run DeleteCell CellName
''        Dim ee As ElementEnumerator

''        If Trim(sCellName) = "" Then GoTo ExitSub

''        'Find and delete Cell
''        ee = oMSTN.ActiveModelReference.Scan

''        Do While ee.MoveNext

''            If ee.Current.IsCellElement Then

''                If UCase(ee.Current.AsCellElement.Name) = UCase(sCellName) Then

''                    If oException Is Nothing Then

''                        oMSTN.ActiveModelReference.RemoveElement(ee.Current)

''                    Else 'check if exception cell equals current cell

''                        If ee.Current.AsCellElement.FilePosition <> oException.FilePosition Then

''                            oMSTN.ActiveModelReference.RemoveElement(ee.Current)

''                        End If

''                    End If

''                End If

''            End If

''        Loop

''ExitSub:
''    End Sub

''Function FindCell(Optional ByVal sCellName As String, Optional ByVal bReturnCellObject As Boolean, Optional oModel As ModelReference, Optional bCount As Boolean)

''        'Searches drawing for specified Cell in particular model
''        'If bReturnCellObject = True then the Cell Object is returned, otherwise a boolean expression is returned
''        'Keyin: vba run FindCell CellName

''        Dim ee As ElementEnumerator
''        Dim oSC As New ElementScanCriteria

''        oSC.ExcludeAllTypes()
''        oSC.IncludeType(oMSTN.msdElementTypeCellHeader)
''        oSC.IncludeType(oMSTN.msdElementTypeSharedCell)

''        'If called from command line get value
''        If sCellName = "" Then
''            sCellName = oMSTN.KeyinArguments
''        End If

''        'Scan appropriate Model for Cells
''        If oModel Is Nothing Then
''            ee = oMSTN.ActiveModelReference.GraphicalElementCache.Scan(oSC)
''        Else
''            ee = oModel.GraphicalElementCache.Scan(oSC)

''        End If

''        Do While ee.MoveNext

''            If ee.Current.IsCellElement Then

''                If UCase(ee.Current.AsCellElement.Name) = UCase(sCellName) Then

''                    'Cell found return object/value
''                    If bReturnCellObject = True Then

''                        FindCell = ee.Current
''                        GoTo ExitFunction

''                    Else

''                        If bCount = True Then

''                            If FindCell = "" Then

''                                FindCell = 1

''                            Else

''                                FindCell = FindCell + 1

''                            End If

''                        Else

''                            FindCell = True
''                            GoTo ExitFunction

''                        End If

''                    End If

''                    If bCount = False Then GoTo ExitFunction

''                End If

''            ElseIf ee.Current.IsSharedCellElement Then

''                If UCase(ee.Current.AsSharedCellElement.Name) = UCase(sCellName) Then

''                    'Cell found return object/value
''                    If bReturnCellObject = True Then

''                        FindCell = ee.Current

''                    Else

''                        If bCount = True Then

''                            If FindCell = "" Then

''                                FindCell = 1

''                            Else

''                                FindCell = FindCell + 1

''                            End If

''                        Else

''                            FindCell = True
''                            GoTo ExitFunction

''                        End If

''                    End If

''                    If bCount = False Then GoTo ExitFunction

''                End If

''            End If

''        Loop

''        'Cell not found return object/value
''        If bCount = False Then

''            If bReturnCellObject = True Then

''                FindCell = Nothing

''            Else

''                FindCell = False

''            End If

''        End If

''ExitFunction:

''    End Function

''    Sub PlaceRevTri()

''        'Place revision triangle
''        'Keyin: vba run PlaceRevTri

''        Dim ee As ElementEnumerator
''        Dim ee1 As ElementEnumerator
''        Dim textText As TextElement
''        Dim strText As String
''        Dim pnt As CadInputMessage
''        Dim celCell As CellElement
''        Dim point As Point3d
''        Dim InpRev As String
''        Dim UserInput As String

''        InpRev = InputBox("Enter revision number", "Drawing Revision")

''        oMSTN.CadInputQueue.SendCommand("ac=revtri")

''        pnt = oMSTN.CadInputQueue.GetInput(oMSTN.msddatapoint)
''        oMSTN.CadInputQueue.SendDataPoint(pnt.Point, 1)
''        oMSTN.CadInputQueue.SendReset()

''        point = pnt.Point

''        ee = oMSTN.ActiveModelReference.Scan

''        Do While ee.MoveNext

''            If ee.Current.IsCellElement Then

''                celCell = ee.Current

''                If celCell.Origin.X = point.X Then

''                    oMSTN.ActiveModelReference.SelectElement(celCell)
''                    ee1 = ee.Current.AsComplexElement.GetSubElements

''                    Do While ee1.MoveNext

''                        If ee1.Current.IsTextElement Then

''                            ee1.Current.AsTextElement.Text = InpRev
''                            ee1.Current.AsTextElement.Rewrite()
''                            'CadInputQueue.SendCommand "update view" 'XM
''                            oMSTN.CadInputQueue.SendCommand("selview all")
''                            oMSTN.CadInputQueue.SendCommand("choose element")

''                        End If

''                    Loop

''                End If

''            End If

''        Loop

''        oMSTN.ActiveModelReference.UnselectAllElements()

''    End Sub

''Sub ReplaceCellText(ByVal sCellName As String, ByVal sText As String, Optional ByVal iColour As Integer)

''        'Replace text and/or change colour within a cell

''        Dim ee As ElementEnumerator
''        Dim ee1 As ElementEnumerator
''        Dim vLines As Object
''        Dim i As Long
''        Dim oTextNode As TextNodeElement
''        Dim oCell As CellElement

''        sCellName = UCase(sCellName)
''        ee = oMSTN.ActiveModelReference.Scan

''        Do While ee.MoveNext

''            If ee.Current.IsCellElement Then

''                With ee.Current.AsCellElement

''                    If UCase(.Name) = sCellName Then

''                        CellFound = True
''                        ee1 = ee.Current.AsCellElement.GetSubElements

''                        Do While ee1.MoveNext

''                            If ee1.Current.IsTextElement Then

''                                ee1.Current.AsTextElement.Text = sText
''                                If iColour <> "" Then ee1.Current.Color = iColour
''                                ee1.Current.Rewrite()

''                                '                        ElseIf ee1.Current.IsTextNodeElement Then
''                                '
''                                '                            Set oTextNode = ee1.Current
''                                '                            vLines = Split(sText, "|")
''                                '                            For i = 0 To UBound(vLines)
''                                '                                If oTextNode.TextLine(i + 1) <> vLines(i) Then
''                                '                                    oTextNode.DeleteTextLine (i + 1)
''                                '                                    'oTextNode.AddTextLine "dsfg"
''                                '                                    oTextNode.TextLine(i + 1) = "fgh" 'vLines(i)
''                                '                                End If
''                                '                            Next
''                                '                            If iColour <> "" Then ee1.Current.Color = iColour
''                                '                            ee1.Current.Rewrite

''                                '                            Set oTextNode = ee1.Current
''                                '                            vLines = Split(sText, "|")
''                                '                            For i = 0 To UBound(vLines)
''                                '                                If oTextNode.TextLine(i + 1) <> vLines(i) Then
''                                '                                    oTextNode.re.DeleteTextLine (i + 1)
''                                '                                    'oTextNode.AddTextLine "dsfg"
''                                '                                    oTextNode.TextLine(i + 1) = "fgh" 'vLines(i)
''                                '                                End If
''                                '                            Next
''                                '                            If iColour <> "" Then ee1.Current.Color = iColour
''                                '                            ee1.Current.Rewrite

''                            End If

''                        Loop

''                        Exit Sub

''                    Else

''                        CellFound = False

''                    End If

''                End With

''            ElseIf ee.Current.IsSharedCellElement Then

''                With ee.Current.AsSharedCellElement

''                    If UCase(.Name) = sCellName Then

''                        CellFound = True
''                        ee1 = ee.Current.AsSharedCellElement.GetSubElements

''                        Do While ee1.MoveNext

''                            If ee1.Current.IsTextElement Then

''                                ee1.Current.AsTextElement.Text = sText
''                                If iColour <> "" Then ee1.Current.Color = iColour
''                                ee1.Current.Rewrite()

''                            End If

''                        Loop

''                        Exit Sub

''                    Else

''                        CellFound = False

''                    End If

''                End With

''            End If

''        Loop

''    End Sub

''    Function ReturnCell(ByVal sCellName As String)

''        'Searches drawing for specified cell and returns Cell object
''        'Keyin: vba run ReturnCell CellName

''        Dim ee As ElementEnumerator

''        If sCellName = "" Then

''            sCellName = oMSTN.KeyinArguments

''        End If

''        ee = oMSTN.ActiveModelReference.Scan

''        Do While ee.MoveNext

''            If ee.Current.IsCellElement Then

''                If ee.Current.AsCellElement.Name = UCase(sCellName) Then

''                    ReturnCell = ee.Current
''                    Exit Function

''                Else

''                    ReturnCell = False

''                End If

''            ElseIf ee.Current.IsSharedCellElement Then

''                If ee.Current.AsCellElement.Name = UCase(sCellName) Then

''                    ReturnCell = ee.Current
''                    Exit Function

''                Else

''                    ReturnCell = False

''                End If

''            End If

''        Loop

''    End Function

''    Function ReturnRef(ByVal sRefName As String)
''        'Searches drawing for specified Reference and returns Object if found
''        'Keyin: vba run ReturnRef RefName
''        Dim oAttachment As Attachment
''        Dim oAttachments As Attachments
''        Dim UserInput As String

''        If sRefName = "" Then
''            sRefName = oMSTN.KeyinArguments
''        End If

''        If UCase(Right(sRefName, 3) <> "DGN") Then sRefName = sRefName & ".DGN"

''        oAttachments = oMSTN.ActiveModelReference.Attachments
''        On Error Resume Next

''        For Each oAttachment In oAttachments

''            If UCase(oAttachment.DesignFile.Name) = UCase(sRefName) Then

''                ReturnRef = oAttachment
''                Exit For

''            Else

''                ReturnRef = False

''            End If

''        Next

''    End Function

''    Sub DeleteCells()
''        'Delete all instances of user selected Cell
''        'Dev Note: Would be good to merge this with the DeleteCell procedure
''        Dim pnt As CadInputMessage
''        Dim sCell As String
''        Dim oElement As Element
''        Dim bCellDeleted As Boolean

''        bCellDeleted = False

''        oMSTN.ActiveModelReference.UnselectAllElements()

''        'Get cell to delete
''SelectCell:
''        oMSTN.CadInputQueue.SendCommand("choose element")
''        oMSTN.ShowPrompt("Select cell to delete")

''        pnt = oMSTN.CadInputQueue.GetInput(oMSTN.msddatapoint)

''        If pnt.InputType = 2 Then
''            GoTo ExitProg
''        End If

''        oMSTN.CadInputQueue.SendDataPoint(pnt.Point, 1)

''        'Get selected cell name
''        If oMSTN.ActiveModelReference.AnyElementsSelected = True Then
''            ee = oMSTN.ActiveModelReference.GetSelectedElements
''            Do While ee.MoveNext
''                If ee.Current.IsCellElement Then
''                    sCell = ee.Current.AsCellElement.Name
''                End If
''            Loop
''        End If

''        'Check if cell selected otherwise quit
''        If sCell = "" Then
''            oMSTN.ShowStatus("No cell selected")
''            oMSTN.ActiveModelReference.UnselectAllElements()
''            GoTo SelectCell
''        End If

''        ee = oMSTN.ActiveModelReference.Scan

''        'Delete cells
''        Do While ee.MoveNext
''            If ee.Current.IsCellElement Then
''                If UCase(ee.Current.AsCellElement.Name) = UCase(sCell) Then
''                    oMSTN.ActiveModelReference.RemoveElement(ee.Current)
''                    bCellDeleted = True
''                End If
''            End If
''        Loop

''        If bCellDeleted = True Then oMSTN.ShowStatus("Cells deleted")

''        GoTo SelectCell
''ExitProg:
''        oMSTN.CadInputQueue.SendCommand(";nocommand")
''        oMSTN.ActiveModelReference.UnselectAllElements()
''        oMSTN.CommandState.StartDefaultCommand()
''        oMSTN.ShowStatus("")
''        oMSTN.ShowPrompt("DeleteCells program exited")
''    End Sub

''    ''    Sub CreateCellBarmenu()
''    ''        'Create barmenu of cell libraries
''    ''        'Author: Paul Ripp
''    ''        Dim oModel As ModelReference
''    ''        Dim oDesign As DesignFile
''    ''        Dim vFiles As Object
''    ''        Dim i As Long
''    ''        Dim sBarMenu As String
''    ''        Dim sDescription As String
''    ''        Dim j As Long
''    ''        Dim k As Long
''    ''        Dim sCellList As String
''    ''        Dim vSplit As Object
''    ''        Dim sSCEApp As String
''    ''        Dim sHeading As String

''    ''        sSCEApp = ActiveWorkspace.ConfigurationVariableValue("SCE_APP_ROOT")
''    ''        sBarMenu = SCE.WorkingPath & SCE.Region & "\" & SCE.Client & "\Workspace\Data\CellLibraries.mdf"
''    ''        If FileExists(sBarMenu) = True Then
''    ''            Delete(sBarMenu)
''    ''        End If
''    ''        sCellList = ActiveWorkspace.ConfigurationVariableValue("MS_CELL")
''    ''        vSplit = Split(sCellList, ";")

''    ''    Open sBarMenu For Output As #1
''    ''    Print #1, "Title=" & SCE.Client & " Cells"
''    ''        sHeading = SCE.Client & " Cell Libraries"
''    ''    Print #1, sHeading
''    ''    Print #1, "{"
''    ''        For l = 0 To UBound(vSplit)
''    ''            If InStr(vSplit(l), sSCEApp) > 0 Then
''    ''                GoTo NextFile
''    ''            ElseIf FolderExists(vSplit(l)) = False Then
''    ''                GoTo NextFile
''    ''            End If
''    ''            vFiles = GetFiles(vSplit(l), "cel", all, False)
''    ''            If vFiles(0) = "" And UBound(vFiles) = 0 Then
''    ''                GoTo NextFile
''    ''            End If
''    ''            If PathParse(UCase(vSplit(l)), LastFolder) <> "CELL" And PathParse(UCase(vSplit(l)), LastFolder) <> "CEL" Then
''    ''            Print #1, "}"
''    ''                sHeading = PathParse(UCase(vSplit(l)), LastFolder)
''    ''            Print #1, sHeading
''    ''            Print #1, "{"
''    ''                j = 0
''    ''                k = 0
''    ''            End If
''    ''            For i = 0 To UBound(vFiles)
''    ''                j = j + 1
''    ''                If j = 40 Then
''    ''                    j = 0
''    ''                    k = k + 1
''    ''                Print #1, "}"
''    ''                Print #1, sHeading & Space(1) & k
''    ''                Print #1, "{"
''    ''                End If
''    ''            Print #1, FileParse(vFiles(i), FileName)
''    ''            Print #1, "{"
''    ''            Print #1, "Load Library" & Chr(44) & Chr(34) & "rc=" & FileParse(vFiles(i), FileName) & ";dialog cellmaintenance" & Chr(34)
''    ''            Print #1, "-,"
''    ''                oDesign = MicroStationDGN.OpenDesignFileForProgram(vSplit(l) & "\" & vFiles(i), True)
''    ''                For Each oModel In oDesign.Models
''    ''                    If oModel.Name = "Default" Then GoTo NextModel
''    ''                    sDescription = Trim(oModel.Description)
''    ''                    If sDescription = "" Then
''    ''                        sDescription = oModel.Name
''    ''                    End If
''    ''                Print #1, sDescription & Chr(44) & Chr(34) & "rc=" & FileParse(vFiles(i), FileName) & ";ac=" & oModel.Name & ";place active cell" & Chr(34)
''    ''NextModel:
''    ''                Next
''    ''                oDesign.Close()
''    ''            Print #1, "}"
''    ''            Next
''    ''NextFile:
''    ''        Next
''    ''    Print #1, "}"
''    ''    Close #1
''    ''        If sDescription = "" Then
''    ''            Delete(sBarMenu)
''    ''            ShowMessage("No Cells were found, barmenu was not created.")
''    ''        End If
''    ''        oDesign = Nothing
''    ''        oModel = Nothing
''    ''ExitSub:
''    ''    End Sub

''    Sub CreateClientCatalogue()
''        CreateCellCatalogue(eCellTypes.tClientCells)
''    End Sub

''    Sub CreateAllClientsCatalogue()
''        CreateCellCatalogue(eCellTypes.tAllClientCells)
''    End Sub

''    Sub CreateNetworkCatalogue()
''        CreateCellCatalogue(eCellTypes.tNetworkCells)
''    End Sub

''    Sub CreateProjectCatalogue()
''        CreateCellCatalogue(eCellTypes.tProjectCells)
''    End Sub

''    Sub CreateUserCatalogue()
''        CreateCellCatalogue(eCellTypes.tUserCells)
''    End Sub

''    Sub CreateCellCatalogue(ByVal sType As eCellTypes)
''        'create cell catalogue for specified cell type
''        'Author: Paul Ripp
''        Dim i As Long
''        Dim j As Long
''        Dim r As Long
''        Dim oModel As ModelReference
''        Dim vFolders As Object
''        Dim vFiles As Object
''        Dim oDesign As DesignFile
''        Dim sCurrent As String
''        Dim sCellLibraryPath As String
''        Dim sDBFullName As String
''        Dim aFolders() As String
''        Dim vRegions As Object
''        Dim bRegions As Boolean
''        Dim bClientOnly As Boolean
''        Dim sTableName As String
''        Dim sLocation As String
''        Dim sModel As String
''        Dim sRegion As String
''        Dim sClient As String
''        Dim sDiscipline As String

''        If SCE.Loaded = False Then SCE.Load()
''        sTableName = "RGCellCatalogue"
''        sCurrent = oMSTN.ActiveDesignFile.FullName
''        Select Case sType
''            Case eCellTypes.tClientCells
''                sDBFullName = SCE.ClientDB
''                sLocation = "Client"
''                sLocation = "Project"
''                sRegion = SCE.Region
''                sClient = SCE.Client
''                vFolders = Split(oMSTN.ActiveWorkspace.ConfigurationVariableValue("MS_CELL"), ";")
''                sCellLibraryPath = oMSTN.ActiveWorkspace.ConfigurationVariableValue("SCE_CLIENT_WORKSPACE") & "Cell\SCECatalogue"
''                If FolderExists(sCellLibraryPath) = True Then
''                    DeleteDir(sCellLibraryPath)
''                End If
''                CreateFolder(sCellLibraryPath)
''                bClientOnly = True
''                ''Case eCellTypes.tAllClientCells
''                ''    sLocation = "Client"
''                ''    vRegions = GetFolders(SCE.SCERegionPath, all)
''                ''    bRegions = True
''                ''    bClientOnly = True
''                ''Case eCellTypes.tNetworkCells
''                ''    sLocation = "Network"
''                ''    sRegion = ""
''                ''    sClient = ""
''                ''    sDBFullName = SCE.NetworkDB
''                ''    ReDim aFolders(0)
''                ''    aFolders(0) = oMSTN.ActiveWorkspace.ConfigurationVariableValue("SCE_NET_WORKSPACE") & "Cell"
''                ''    vFolders = aFolders
''                ''    sCellLibraryPath = oMSTN.ActiveWorkspace.ConfigurationVariableValue("SCE_NET_WORKSPACE") & "Cell\SCECatalogue"
''                ''    If FolderExists(sCellLibraryPath) = True Then
''                ''        DeleteDir(sCellLibraryPath)
''                ''    End If
''                ''    CreateFolder(sCellLibraryPath)
''                ''Case eCellTypes.tProjectCells
''                ''    sLocation = "Project"
''                ''    sRegion = SCE.Region
''                ''    sClient = SCE.Client
''                ''    If oMSTN.ActiveWorkspace.IsConfigurationVariableDefined("SCE_PROJECT_WORKSPACE") = True Then
''                ''        sDBFullName = SCE.ProjectDB
''                ''        ReDim aFolders(0)
''                ''        aFolders(0) = oMSTN.ActiveWorkspace.ConfigurationVariableValue("SCE_PROJECT_WORKSPACE") & "Cell"
''                ''        vFolders = aFolders
''                ''        sCellLibraryPath = vFolders(0) & "\SCECatalogue"
''                ''        If FolderExists(sCellLibraryPath) = True Then
''                ''            DeleteDir(sCellLibraryPath)
''                ''        End If
''                ''        CreateFolder(sCellLibraryPath)
''                ''    Else
''                ''        ShowError("Project not configured.")
''                ''        GoTo ExitSub
''                ''    End If
''                ''Case eCellTypes.tUserCells
''                ''    sLocation = "User"
''                ''    sRegion = ""
''                ''    sClient = ""
''                ''    sDBFullName = SCE.UserDB
''                ''    ReDim aFolders(0)
''                ''    aFolders(0) = oMSTN.ActiveWorkspace.ConfigurationVariableValue("SCE_USER_WORKSPACE") & "Cell"
''                ''    vFolders = aFolders
''                ''    sCellLibraryPath = vFolders(0) & "\Cell Catalogue"
''                ''    If FolderExists(sCellLibraryPath) = True Then
''                ''        DeleteDir(sCellLibraryPath)
''                ''    End If
''                ''    CreateFolder(sCellLibraryPath)
''            Case Else
''                GoTo ExitSub
''        End Select

''        If bRegions = False Then
''            GoTo ProcessFolders
''        End If

''        For r = 0 To UBound(vRegions)
''            sDBFullName = SCE.ClientDB
''            sRegion = SCE.Region
''            sClient = SCE.Client
''            vFolders = GetFolders(SCE.SCERegionPath & vRegions(r), all)
''            sCellLibraryPath = oMSTN.ActiveWorkspace.ConfigurationVariableValue("SCE_CLIENT_WORKSPACE") & "Cell\SCECatalogue"
''            If FolderExists(sCellLibraryPath) = True Then
''                DeleteDir(sCellLibraryPath)
''            End If
''            CreateFolder(sCellLibraryPath)
''ProcessFolders:
''            If TableExists(sDBFullName, sTableName) = True Then
''                DeleteTable(sDBFullName, sTableName)
''            End If
''            CreateTable(sDBFullName, sTableName, "ID|CellName|Description|Library|Location|Region|Client|Discipline|Dimension|CellLibPath|CustomDescription|CUIPath|Rotate|Level")

''            For i = 0 To UBound(vFolders)
''                vFiles = GetFiles(vFolders(i), "cel", all, False)
''                If vFiles(0) = "" And UBound(vFiles) = 0 Then GoTo NextFolder
''                'Check if cell in client location
''                If bClientOnly = True Then
''                    If InStr(UCase(vFolders(i)), UCase(oMSTN.ActiveWorkspace.ConfigurationVariableValue("SCE_CLIENT_WORKSPACE"))) = -1 Then
''                        GoTo NextFolder
''                    End If
''                End If

''                For j = 0 To UBound(vFiles)
''                    'Index Cells
''                    oDesign = oMSTN.OpenDesignFile(vFolders(i) & "\" & vFiles(j), False, oMSTN.msdV7ActionWorkmode)
''                    For Each oModel In oDesign.Models
''                        If oModel.Name = "Default" Then GoTo NextModel
''                        oModel.Activate()
''                        oMSTN.ActiveWorkspace.AddConfigurationVariable("_SCEJPG", sCellLibraryPath & "\" & oModel.Name, False)
''                        If oModel.Is3D = True Then
''                            sModel = "3D"
''                        Else
''                            sModel = "2D"
''                        End If
''                        If InStr(UCase(oModel.Description), "PATTERN") > 0 Then
''                            sDiscipline = "Pattern"
''                        Else
''                            sDiscipline = ""
''                        End If
''                        If SCE.Region = "" Then
''                            SCE.Load()
''                            If SCE.Region = "" Then
''                                SetDB(sDBFullName, sTableName, "CellName|Description|Library|Location|Region|Client|Discipline|Dimension|CellLibPath", _
''                                    oModel.Name & "|" & Replace(oModel.Description, "'", "") & "|" & oDesign.FullName & "|" & sLocation & "|" & SCE.Region & "|" & SCE.Client & "|" & sDiscipline & "|" & sModel & "|" & oDesign.FullName)
''                                GoTo CreateImage
''                            End If
''                        End If
''                        SetDB(sDBFullName, sTableName, "CellName|Description|Library|Location|Region|Client|Discipline|Dimension|CellLibPath", _
''                            oModel.Name & "|" & Replace(oModel.Description, "'", "") & "|" & oDesign.FullName & "|" & sLocation & "|" & SCE.Region & "|" & SCE.Client & "|" & sDiscipline & "|" & sModel & "|" & oDesign.FullName)
''CreateImage:
''                        'Reset View 1
''                        oMSTN.ActiveDesignFile.Views.Item(1).DisplaysDataEntryRegions = True
''                        oMSTN.ActiveDesignFile.Views.Item(1).DisplaysFill = True
''                        oMSTN.CadInputQueue.SendCommand("FIT ALL;SELVIEW 1")
''                        oMSTN.SaveSettings()
''                        'Create thumbnail image
''                        CreateThumbNailJPG()
''                        'ActiveWorkspace.AddConfigurationVariable "_SCEJPG", sCellLibraryPath & "\" & oModel.Name & "_Small", False
''                        'CreateThumbNailJPG_Small
''NextModel:
''                    Next
''                    oDesign.Close()
''                Next
''NextFolder:
''            Next
''            CompactDB(sDBFullName)
''            If bRegions = False Then
''                Exit For
''            End If
''        Next

''        If Not oDesign Is Nothing Then
''            oDesign = Nothing
''        End If
''        oMSTN.OpenDesignFile(sCurrent, False)
''ExitSub:
''    End Sub

''    Sub CreateThumbNailJPG()
''        'Create small images for Cell Library indexing
''        oMSTN.CadInputQueue.SendCommand("print driver " & oMSTN.ActiveWorkspace.ConfigurationVariableValue("SCE_PLTDRV") & "CellCatalogueJPG.plt")  '
''        oMSTN.CadInputQueue.SendCommand("print papername A4")
''        oMSTN.CadInputQueue.SendCommand("print fullsheet on")
''        oMSTN.CadInputQueue.SendCommand("print attributes datafields on")
''        oMSTN.CadInputQueue.SendCommand("print attributes fill on")
''        oMSTN.CadInputQueue.SendCommand("print attributes transparency on")
''        oMSTN.CadInputQueue.SendCommand("print boundary view")
''        oMSTN.CadInputQueue.SendCommand("print execute")
''    End Sub

''    Sub CreateThumbNailJPG_Small()
''        'Create small images for Cell Library indexing
''        oMSTN.CadInputQueue.SendCommand("print driver " & oMSTN.ActiveWorkspace.ConfigurationVariableValue("SCE_PLTDRV") & "CellCatalogueJPG_Small.plt")
''        oMSTN.CadInputQueue.SendCommand("print papername A4")
''        oMSTN.CadInputQueue.SendCommand("print fullsheet on")
''        oMSTN.CadInputQueue.SendCommand("print attributes datafields on")
''        oMSTN.CadInputQueue.SendCommand("print attributes fill on")
''        oMSTN.CadInputQueue.SendCommand("print attributes transparency on")
''        oMSTN.CadInputQueue.SendCommand("print boundary view")
''        oMSTN.CadInputQueue.SendCommand("print execute")
''    End Sub

''    Sub ADMINCreateCellLibrary()
''        'create Cell library using all regions and clients
''        'used by admin only
''        Dim i As Long
''        Dim r As Long
''        Dim oModel As ModelReference
''        Dim vFolders As Object
''        Dim vFiles As Object
''        Dim oDesign As DesignFile
''        Dim sCurrent As String
''        Dim sCellLibraryPath As String
''        Dim sDBFullName As String
''        Dim vRegions As Object
''        Dim sFolder As String

''        sCurrent = oMSTN.ActiveDesignFile.FullName
''        sFolder = "\Workspace\Cell"
''        sCellLibraryPath = "D:\CellLibrary"
''        CreateFolder(sCellLibraryPath)
''        sDBFullName = sCellLibraryPath & "\CellLibrary.accdb"
''        CreateDB(sDBFullName)
''        If TableExists(sDBFullName, "RGCellLibrary") = True Then
''            DeleteTable(sDBFullName, "RGCellLibrary")
''        End If
''        CreateTable(sDBFullName, "RGCellLibrary", "CellName|Description|Library|Region|Client|Discipline|CellLibPath")

''        vRegions = GetFolders(SCE.SCERegionPath, all)
''        For r = 0 To UBound(vRegions)

''            vFolders = GetFolders(SCE.SCERegionPath & vRegions(r), all)
''            For i = 0 To UBound(vFolders)
''                vFiles = GetFiles(SCE.SCERegionPath & vRegions(r) & "\" & vFolders(i) & sFolder, "cel", all, False)
''                If vFiles(0) = "" And UBound(vFiles) = 0 Then GoTo NextFolder
''                CreateFolder(sCellLibraryPath & "\" & vRegions(r) & " - " & vFolders(i))
''                For j = 0 To UBound(vFiles)
''                    'Index Cells
''                    oDesign = oMSTN.OpenDesignFile(SCE.SCERegionPath & vRegions(r) & "\" & vFolders(i) & sFolder & "\" & vFiles(j), False, msdV7ActionWorkmode)
''                    For Each oModel In oDesign.Models
''                        If oModel.Name = "Default" Then GoTo NextModel
''                        oModel.Activate()
''                        oMSTN.ActiveWorkspace.AddConfigurationVariable("_SCEJPG", sCellLibraryPath & "\" & vRegions(r) & " - " & vFolders(i) & "\" & oModel.Name & ".jpg", False)
''                        If SCE.Region = "" Then
''                            SCE.Load()
''                            If SCE.Region = "" Then
''                                SetDB(sDBFullName, "RGCellLibrary", "CellName|Description|Library|Region|Client|CellLibPath", _
''                                    oModel.Name & "|" & oModel.Description & "|" & oDesign.FullName & "|" & vRegions(r) & "|" & vFolders(i) & "|" & oDesign.FullName)
''                                GoTo CreateImage
''                            End If
''                        End If
''                        SetDB(sDBFullName, "RGCellLibrary", "CellName|Description|Library|Region|Client|CellLibPath", _
''                            oModel.Name & "|" & oModel.Description & "|" & oDesign.Name & "|" & vRegions(r) & "|" & vFolders(i) & "|" & oDesign.FullName)
''CreateImage:
''                        oMSTN.CadInputQueue.SendCommand("FIT ALL;SELVIEW 1")
''                        oMSTN.SaveSettings()
''                        CreateThumbNailJPG()
''NextModel:
''                    Next
''                    oDesign.Close()
''                    'GoTo NextFolder 'tests changing clients and regions quicker
''                Next
''NextFolder:
''            Next
''        Next
''        CompactDB(sDBFullName)
''        If Not oDesign Is Nothing Then
''            oDesign = Nothing
''        End If
''        oMSTN.OpenDesignFile(sCurrent, False)
''    End Sub

''End Class


