'Imports MicroStationDGN

'<Microsoft.VisualBasic.ComClass()> Public Class MSTNAutoPlaceCell
'    Implements IPrimitiveCommandEvents
'#Region "COM GUIDs"
'    ' These  GUIDs provide the COM identity for this class 
'    ' and its COM interfaces. If you change them, existing 
'    ' clients will no longer be able to access the class.
'    Public Const ClassId As String = "8108d5b6-e8a3-4c3c-a0f9-a4b0957dcd85"
'    Public Const InterfaceId As String = "f8f90b2c-4ad1-42bd-9bf2-f308bcf79e3a"
'    Public Const EventsId As String = "7156fe29-857b-47e0-8f24-1e76af557ef2"
'#End Region

'    ' A creatable COM class must have a Public Sub New() 
'    ' with no parameters, otherwise, the class will not be 
'    ' registered in the COM registry and cannot be created 
'    ' via CreateObject.
'    Public Sub New()
'        MyBase.New()
'    End Sub

'    Private Shared _myInstance As MicroStationDGN.Application
'    Public Shared oMSTN As MicroStationDGN.Application = GetInstance()

'    Public Shared Function GetInstance() As MicroStationDGN.Application
'        If _myInstance Is Nothing Then
'            _myInstance = New MicroStationDGN.Application 'TODO: Add check if App is already(running)
'        End If
'        Return _myInstance
'    End Function



'    Private sCellName As String
'    Private oCell As CellElement

'    Private Sub iPrimitiveCommandEvents_Cleanup()
'        oMSTN.CommandState.StartDefaultCommand()
'    End Sub

'    Private Sub iPrimitiveCommandEvents_DataPoint(ByVal point As Point3d, ByVal View As View)
'        Dim ee As ElementEnumerator = Nothing
'        Dim sOverride As String
'        Dim tempoint As Point3d
'        Dim eleELE As Element
'        Dim lngTemp As Long
'        Dim sLevel As String
'        Dim oLevel As Level

'        oMSTN.ActiveModelReference.UnselectAllElements()
'        sOverride = oMSTN.ActiveWorkspace.ConfigurationVariableValue("_TMP_CELLLEVELOVERRIDE")
'        sLevel = oMSTN.ActiveWorkspace.ConfigurationVariableValue("_TMP_CELLLEVEL")

'        oMSTN.ActiveModelReference.AddElement(oCell)

'        If sOverride = "True" Then
'            ee = oMSTN.ActiveModelReference.GraphicalElementCache.Scan

'            Do While ee.MoveNext
'                If TypeOf ee.Current Is CellElement Then
'                    If ee.Current.AsCellElement.Name = oCell.Name And ee.Current.FilePosition = oMSTN.ActiveModelReference.GetLastValidGraphicalElement.FilePosition Then
'                        ee = ee.Current.AsCellElement.GetSubElements
'                        Do While ee.MoveNext
'                            ee.Current.Level = oMSTN.ActiveSettings.Level
'                            ee.Current.LineStyle = oMSTN.ByLevelLineStyle
'                            ee.Current.Rewrite()
'                            ee.Current.Redraw()
'                        Loop
'                        Exit Do
'                    End If
'                End If
'            Loop
'        End If

'        If sLevel <> "" Then
'            If TypeOf ee.Current Is CellElement Then
'                If ee.Current.AsCellElement.Name = oCell.Name And ee.Current.FilePosition = oMSTN.ActiveModelReference.GetLastValidGraphicalElement.FilePosition Then
'                    ee = ee.Current.AsCellElement.GetSubElements
'                    Do While ee.MoveNext
'                        oLevel = oMSTN.GetLevel(sLevel)
'                        ee.Current.Level = oLevel
'                        ee.Current.LineStyle = oMSTN.ByLevelLineStyle
'                        ee.Current.Rewrite()
'                        ee.Current.Redraw()
'                    Loop
'                End If
'            End If
'        End If

'        oMSTN.CadInputQueue.SendCommand("UPDATE ALL")

'        eleELE = oMSTN.ActiveModelReference.GetLastValidGraphicalElement
'        oMSTN.ActiveModelReference.SelectElement(eleELE)
'        tempoint = eleELE.AsCellElement.Origin

'        '   Start a command
'        oMSTN.CadInputQueue.SendCommand("ROTATE ICON")

'        '   Set a variable associated with a dialog box
'        '   This only modifies a few bits of the variable it changes. It first
'        '   creates a mask for clearing the bits it will change. Then it gets
'        '   the variable and uses the mask to clear those bits. Finally
'        '   it sets the desired bits in the value and saves the updated value.
'        lngTemp = Not 1
'        lngTemp = oMSTN.GetCExpressionValue("settings.aboutOrigin", "TRANSFRM") And lngTemp
'        oMSTN.SetCExpressionValue("settings.aboutOrigin", lngTemp Or 1, "TRANSFRM")
'        oMSTN.SetCExpressionValue("tcb->msToolSettings.rotate.method", 1, "TRANSFRM")

'        '   Send 2 data points to the current command
'        oMSTN.CadInputQueue.SendDataPoint(tempoint, 1)
'        oMSTN.CadInputQueue.SendDataPoint(tempoint, 1)
'        oMSTN.ActiveModelReference.UnselectAllElements()
'    End Sub

'    Private Sub iPrimitiveCommandEvents_Dynamics(ByVal point As Point3d, ByVal View As View, ByVal DrawMode As MsdDrawingMode)
'        oCell = oMSTN.CreateCellElement3(sCellName, point, True)
'        oCell.Redraw(DrawMode)
'    End Sub

'    Private Sub iPrimitiveCommandEvents_Keyin(ByVal Keyin As String)

'    End Sub

'    Private Sub iPrimitiveCommandEvents_Reset()
'        oMSTN.CommandState.StartDefaultCommand()
'    End Sub

'    Private Sub iPrimitiveCommandEvents_Start()
'        sCellName = oMSTN.ActiveWorkspace.ConfigurationVariableValue("_TMP_CELLNAME")
'        oMSTN.CommandState.StartDynamics()
'    End Sub







'End Class


