<ComClass(MSTNApplicationData.ClassId, MSTNApplicationData.InterfaceId, MSTNApplicationData.EventsId)> _
Public Class MSTNApplicationData

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "23d16b07-63f0-49f5-b2fc-06353a40efaa"
    Public Const InterfaceId As String = "5b291771-1b1c-4b4d-9c38-944eac53a8c6"
    Public Const EventsId As String = "eff2ba1c-bfcb-43b4-a283-3dac556f0aa0"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub


    '--------------------------------------------------------------------------------
    '
    '   Application Data
    '
    '   - Modify Application Data in the Drawing
    '
    '---------------------------------------------------------------------------------

    Private Const myID As Long = 56345
    Private Const cMsgBoxCaption = "SKM Application Data"

    Private Shared _myInstance As MicroStationDGN.Application
    Public Shared oMSTN As MicroStationDGN.Application = GetInstance()

    Public Shared Function GetInstance() As MicroStationDGN.Application
        If _myInstance Is Nothing Then
            _myInstance = New MicroStationDGN.Application 'TODO: Add check if App is already(running)
        End If
        Return _myInstance
    End Function

    Shared Sub SetAppData(ByVal sValue As String)
        'Writes Application Data (Type 66) to the Drawing
        Dim ae As ApplicationElement
        Dim DB As New DataBlock
        Dim vValue As Array = Nothing
        Dim bIsLocked As Boolean
        Dim i As Integer

        'Check if Model is locked and if so unlock
        bIsLocked = oMSTN.ActiveModelReference.IsLocked

        If bIsLocked = True Then oMSTN.ActiveModelReference.IsLocked = False

        'If comma delimited value is passed separate
        If sValue <> "" Then vValue = Split(sValue, ",")

        'Delete App Data if already exists
        DeleteAppData(CStr(vValue(0)))

        'Create App Data
        DB.CopyString(CStr(vValue(0)), True)
        DB.CopyInteger(CStr(vValue(1)), True)

        For i = 2 To UBound(vValue)
            If vValue(i) = "" Then
                DB.CopyString(" ", True)
            Else
                DB.CopyString(CStr(vValue(i)), True)
            End If
        Next

        ae = oMSTN.CreateApplicationElement(myID, DB)

        oMSTN.ActiveModelReference.AddElement(ae)

        'If Model was locked relock Model
        If bIsLocked = True Then oMSTN.ActiveModelReference.IsLocked = True
ExitSub:
    End Sub

    Shared Function GetAppData(Optional ByVal sValue As String = "", Optional ByVal oModel As ModelReference = Nothing) As Object
        'Return App Data
        'If no value supplied return all App Data
        Dim ee As ElementEnumerator
        Dim ae As ApplicationElement
        Dim data As Long
        Dim sTotalValue As String
        Dim iValues As Integer
        Dim sString As String = ""
        Dim vValue As Object
        Dim aTotalValue() As String
        Dim DB As DataBlock
        Dim bFound As Boolean
        Dim oSC As New ElementScanCriteria
        Dim i As Integer
        Dim x As Integer
        Dim aReturn(0) As String

        oSC.ExcludeAllTypes()
        oSC.IncludeType(oMSTN.msdElementTypeMicroStation)
        oSC.IncludeType(oMSTN.msdElementSubtypeApplicationElement)

        If oModel Is Nothing Then oModel = oMSTN.ActiveModelReference

        ee = oModel.ControlElementCache.Scan(oSC)

        ReDim aTotalValue(0)

        i = -1
        bFound = False

        On Error GoTo Err

        Do While ee.MoveNext
            If ee.Current.Type = oMSTN.msdElementTypeMicroStation And ee.Current.Subtype = oMSTN.msdElementSubtypeApplicationElement Then
                ae = ee.Current
                If ae.ApplicationID = myID Then
                    DB = ae.GetApplicationData
                    If sValue = "All" Or sValue = "" Then
                        i = i + 1
                        ReDim Preserve aTotalValue(i)

                        DB.CopyString(aTotalValue(i), False)
                        DB.CopyInteger(iValues, False)

                        For x = 1 To iValues
                            On Error GoTo Err
                            DB.CopyString(sString, False)
                            aTotalValue(i) = aTotalValue(i) & "," & sString
                        Next
                        If iValues > 0 Then bFound = True
                    Else
                        i = i + 1
                        ReDim aTotalValue(i)

                        DB.CopyString(aTotalValue(i), False)
                        DB.CopyInteger(iValues, False)

                        ReDim Preserve aTotalValue(i + iValues)

                        For x = 1 To iValues
                            i = i + 1
                            DB.CopyString(aTotalValue(i), False)
                        Next

                        If UCase(aTotalValue(0)) = UCase(sValue) Then
                            If UBound(aTotalValue) = 1 Then
                                aReturn = aTotalValue
                            Else
                                aReturn = aTotalValue
                            End If

                            GoTo ExitFunction
                        Else
                            Erase aTotalValue
                            i = -1

                            ReDim aTotalValue(0)
                        End If
                    End If
                End If
            End If
        Loop

        If bFound = True Then
            'Return App Data value(s)
            aReturn = aTotalValue
        Else
            'If no App Data found pass string "False"
            Erase aTotalValue
            ReDim aTotalValue(1)

            aTotalValue(0) = ""
            aReturn = aTotalValue
        End If

        GoTo ExitFunction

Err:

        If Err.Number = -2147218390 Then 'EOF (no App Data Found)
            Erase aTotalValue

            ReDim aTotalValue(1)

            aTotalValue(0) = ""
            aReturn = aTotalValue

            GoTo ExitFunction
        ElseIf Err.Number = 28 Then 'Out of stack space
            ReDim aTotalValue(1)

            aTotalValue(0) = ""
            aReturn = aTotalValue
            GoTo ExitFunction
        Else
            MsgBox(Err.Number & ": " & Err.Description, vbCritical, cMsgBoxCaption)
            Resume Next
        End If
ExitFunction:
        Return aReturn
    End Function

    Shared Sub DeleteAppData(Optional ByVal sValue As String = "")

        'Delete Application Data
        'If no value given, delete all Application Data

        Dim bLocked As Boolean
        Dim DB As New DataBlock
        Dim ee As ElementEnumerator
        Dim sReturnValue As String = ""

        If oMSTN.ActiveModelReference.IsLocked = True Then

            bLocked = True
            oMSTN.ActiveModelReference.IsLocked = False

        End If

        ee = oMSTN.ActiveModelReference.ControlElementCache.Scan

        Do While ee.MoveNext

            If ee.Current.IsApplicationElement Then

                With ee.Current.AsApplicationElement

                    If .ApplicationID = myID Then

                        DB = ee.Current.AsApplicationElement.GetApplicationData

                        DB.CopyString(sReturnValue, False)

                        If sValue = "" Then

                            On Error GoTo Err

                            oMSTN.ActiveModelReference.RemoveElement(ee.Current)

                            On Error GoTo 0

                        Else

                            If UCase(sReturnValue) = UCase(sValue) Then

                                On Error GoTo Err

                                oMSTN.ActiveModelReference.RemoveElement(ee.Current)

                                On Error GoTo 0

                            End If

                        End If

                    End If

                End With

            End If

        Loop

        If bLocked = True Then oMSTN.ActiveModelReference.IsLocked = True

        Exit Sub

Err:

        If Err.Number = -2147218396 Or Err.Number = -2147218394 Then

            MsgBox("File is read only or Model is locked", vbCritical, cMsgBoxCaption)

        Else

            MsgBox(Err.Number & ": " & Err.Description, vbCritical, cMsgBoxCaption)

        End If

        Exit Sub

    End Sub

    Shared Sub DeleteAllAppData()

        'Deletes all Application Data
        'NOTE: This may not be required since DelAppData does this anyway

        Dim ee As ElementEnumerator
        Dim sReturnValue As String = ""
        Dim DB As Object

        ee = oMSTN.ActiveModelReference.ControlElementCache.Scan

        Do While ee.MoveNext

            If ee.Current.IsApplicationElement Then

                With ee.Current.AsApplicationElement

                    If .ApplicationID = myID Then

                        DB = ee.Current.AsApplicationElement.GetApplicationData

                        DB.CopyString(sReturnValue, False)

                        On Error GoTo Err

                        oMSTN.ActiveModelReference.RemoveElement(ee.Current)

                        On Error GoTo 0

                    End If

                End With

            End If

        Loop

        Exit Sub

Err:

        If Err.Number = -2147218396 Or Err.Number = -2147218394 Then

            MsgBox("File read only or Model is locked", vbCritical, cMsgBoxCaption)

        Else

            MsgBox(Err.Number & ": " & Err.Description, vbCritical, cMsgBoxCaption)

        End If

        Exit Sub

    End Sub

    Shared Sub CopyAppdata()

        'Copy all Application Data from specified drawing into current drawing

        Dim oDesignFile As DesignFile
        Dim oModel As ModelReference
        Dim vAppData As Object
        Dim i As Long
        Dim vSplit As Object
        Dim RetVal

        RetVal = MsgBox("The current drawing must be the drawing you would like to copy the Application Data INTO." & Chr(10) & Chr(10) & "Would you like to proceed?", 35, cMsgBoxCaption)

        If RetVal = vbYes Then

            'Get source drawing
            RetVal = InputBox("Enter the full file name of the drawing to retreive the Application Data from:", cMsgBoxCaption)

            If My.Computer.FileSystem.FileExists(RetVal) = True Then

                'Open source drawing
                oDesignFile = oMSTN.OpenDesignFileForProgram(RetVal, True)
                oModel = oDesignFile.DefaultModelReference

                'Get Application Data
                vAppData = GetAppData("All", oModel)

                'Close source drawing
                oModel = Nothing
                oDesignFile.Close()
                oDesignFile = Nothing

                'Add Application Data to destination (current) drawing
                For i = 0 To UBound(vAppData)

                    vSplit = Split(vAppData(i), ",")

                    SetAppData(vSplit(0) & ",1," & vSplit(1))

                Next

                oMSTN.ActiveDesignFile.Save()

                MsgBox("Application Data copy has completed", vbInformation, cMsgBoxCaption)

            ElseIf Trim(RetVal) = "" Then 'cancelled inputbox

                MsgBox("No Application Data has been copied", vbInformation, cMsgBoxCaption)

            Else 'file doesn't exist

                MsgBox(Trim(RetVal) & Space(1) & "doesn't exist", vbExclamation, cMsgBoxCaption)

            End If

        End If

    End Sub

    Shared Function HasAppData(Optional ByVal oModel As ModelReference = Nothing) As Boolean
        'Return if current model has any App Data
        Dim ee As ElementEnumerator
        Dim ae As ApplicationElement
        Dim oSC As New ElementScanCriteria

        oSC.ExcludeAllTypes()
        oSC.IncludeType(oMSTN.msdElementTypeMicroStation)
        oSC.IncludeType(oMSTN.msdElementSubtypeApplicationElement)

        If oModel Is Nothing Then oModel = oMSTN.ActiveModelReference

        ee = oModel.ControlElementCache.Scan(oSC)

        HasAppData = False

        On Error GoTo Err

        Do While ee.MoveNext

            If ee.Current.Type = oMSTN.msdElementTypeMicroStation And ee.Current.Subtype = oMSTN.msdElementSubtypeApplicationElement Then

                ae = ee.Current

                If ae.ApplicationID = myID Then

                    HasAppData = True
                    Exit Do

                End If

            End If

        Loop
Err:
    End Function
End Class


