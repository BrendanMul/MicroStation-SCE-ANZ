'Public Class SCEMicroStation

'    Private Const iGeneralID As Long = 56345
'    Public Shared oMSTNApplication As New MicroStationDGN.Application
'    Public Shared oDesignfile As MicroStationDGN.DesignFile

'    Public Shared Sub OpenMicroStation(ByVal sFileName As String)
'        'open microstation session
'        Dim oMSTNAppl As New MicroStationDGN.Application
'        If My.Computer.FileSystem.FileExists(sFileName) = True Then
'            oDesignfile = oMSTNAppl.OpenDesignFileForProgram(sFileName, True)

'            Dim iID As Integer = iGeneralID
'            'Return App Data
'            'If no value supplied return all App Data
'            Dim ee As ElementEnumerator
'            Dim ae As ApplicationElement
'            Dim data As Long
'            Dim sTotalValue As String
'            Dim iValues As Integer
'            Dim sString As String = ""
'            Dim vValue As Object
'            Dim aTotalValue() As String
'            Dim DB As DataBlock
'            Dim bFound As Boolean = False
'            Dim oSC As New ElementScanCriteria
'            Dim i As Integer = -1
'            Dim x As Integer
'            Dim aReturn(0) As String
'            Dim sValue As String = "SKMArea"

'            oSC.ExcludeAllTypes()
'            'oSC.IncludeType(MicroStationDGN.msdElementTypeMicroStation)
'            oSC.IncludeType(MicroStationDGN.MsdElementSubtype.msdElementSubtypeApplicationElement)

'            'If oModel Is Nothing Then
'            '    oModel = oMSTNApplication.ActiveDesignFile.DefaultModelReference
'            'End If

'            ee = oDesignfile.DefaultModelReference.ControlElementCache.Scan(oSC)

'            ReDim aTotalValue(0)

'            On Error GoTo Err

'            Do While ee.MoveNext
'                If ee.Current.Type = oMSTNApplication.msdElementTypeMicroStation And ee.Current.Subtype = oMSTNApplication.msdElementSubtypeApplicationElement Then
'                    ae = ee.Current
'                    If ae.ApplicationID = iID Then
'                        DB = ae.GetApplicationData
'                        If sValue = "All" Or sValue = "" Then
'                            i = i + 1
'                            ReDim Preserve aTotalValue(i)

'                            DB.CopyString(aTotalValue(i), False)
'                            DB.CopyInteger(iValues, False)

'                            For x = 1 To iValues
'                                On Error GoTo Err
'                                DB.CopyString(sString, False)
'                                aTotalValue(i) = aTotalValue(i) & "," & sString
'                            Next
'                            If iValues > 0 Then bFound = True
'                        Else
'                            i = i + 1
'                            ReDim aTotalValue(i)

'                            DB.CopyString(aTotalValue(i), False)
'                            DB.CopyInteger(iValues, False)

'                            ReDim Preserve aTotalValue(i + iValues)

'                            For x = 1 To iValues
'                                i = i + 1
'                                DB.CopyString(aTotalValue(i), False)
'                            Next

'                            If UCase(aTotalValue(0)) = UCase(sValue) Then
'                                If UBound(aTotalValue) = 1 Then
'                                    aReturn = aTotalValue
'                                Else
'                                    aReturn = aTotalValue
'                                End If

'                                GoTo ExitFunction
'                            Else
'                                Erase aTotalValue
'                                i = -1
'                                ReDim aTotalValue(0)
'                            End If
'                        End If
'                    End If
'                End If
'            Loop

'            If bFound = True Then
'                aReturn = aTotalValue
'            Else
'                'If no App Data found pass string "False"
'                Erase aTotalValue
'                ReDim aTotalValue(1)

'                aTotalValue(0) = ""
'                aReturn = aTotalValue
'            End If

'            GoTo ExitFunction
'Err:
'            If Err.Number = -2147218390 Then 'EOF (no App Data Found)
'                Erase aTotalValue
'                ReDim aTotalValue(1)

'                aTotalValue(0) = ""
'                aReturn = aTotalValue

'                GoTo ExitFunction
'            ElseIf Err.Number = 28 Then 'Out of stack space
'                ReDim aTotalValue(1)

'                aTotalValue(0) = ""
'                aReturn = aTotalValue
'                GoTo ExitFunction
'            Else
'                MsgBox(Err.Number & ": " & Err.Description, vbCritical, JEGCore.MsgBoxCaption)
'                Resume Next
'            End If
'ExitFunction:
'            MsgBox(aReturn(0))


'        End If
'    End Sub

'    Public Shared Sub CloseMicroStation()
'        'open microstation session
'        oDesignfile.Close()
'        oMSTNApplication.Quit()
'        oMSTNApplication = Nothing
'    End Sub

'    Public Shared Function GetAppData(Optional ByVal sValue As String = "", Optional ByVal oModel As ModelReference = Nothing, _
'        Optional ByVal iID As Integer = iGeneralID) As String
'        'Return App Data
'        'If no value supplied return all App Data
'        Dim ee As ElementEnumerator
'        Dim ae As ApplicationElement
'        Dim data As Long
'        Dim sTotalValue As String
'        Dim iValues As Integer
'        Dim sString As String = ""
'        Dim vValue As Object
'        Dim aTotalValue() As String
'        Dim DB As DataBlock
'        Dim bFound As Boolean = False
'        Dim oSC As New ElementScanCriteria
'        Dim i As Integer = -1
'        Dim x As Integer
'        Dim aReturn(0) As String

'        oSC.ExcludeAllTypes()
'        'oSC.IncludeType(MicroStationDGN.msdElementTypeMicroStation)
'        oSC.IncludeType(MicroStationDGN.MsdElementSubtype.msdElementSubtypeApplicationElement)

'        If oModel Is Nothing Then
'            oModel = oMSTNApplication.ActiveDesignFile.DefaultModelReference
'        End If

'        ee = oModel.ControlElementCache.Scan(oSC)

'        ReDim aTotalValue(0)

'        On Error GoTo Err

'        Do While ee.MoveNext
'            If ee.Current.Type = oMSTNApplication.msdElementTypeMicroStation And ee.Current.Subtype = oMSTNApplication.msdElementSubtypeApplicationElement Then
'                ae = ee.Current
'                If ae.ApplicationID = iID Then
'                    DB = ae.GetApplicationData
'                    If sValue = "All" Or sValue = "" Then
'                        i = i + 1
'                        ReDim Preserve aTotalValue(i)

'                        DB.CopyString(aTotalValue(i), False)
'                        DB.CopyInteger(iValues, False)

'                        For x = 1 To iValues
'                            On Error GoTo Err
'                            DB.CopyString(sString, False)
'                            aTotalValue(i) = aTotalValue(i) & "," & sString
'                        Next
'                        If iValues > 0 Then bFound = True
'                    Else
'                        i = i + 1
'                        ReDim aTotalValue(i)

'                        DB.CopyString(aTotalValue(i), False)
'                        DB.CopyInteger(iValues, False)

'                        ReDim Preserve aTotalValue(i + iValues)

'                        For x = 1 To iValues
'                            i = i + 1
'                            DB.CopyString(aTotalValue(i), False)
'                        Next

'                        If UCase(aTotalValue(0)) = UCase(sValue) Then
'                            If UBound(aTotalValue) = 1 Then
'                                aReturn = aTotalValue
'                            Else
'                                aReturn = aTotalValue
'                            End If

'                            GoTo ExitFunction
'                        Else
'                            Erase aTotalValue
'                            i = -1
'                            ReDim aTotalValue(0)
'                        End If
'                    End If
'                End If
'            End If
'        Loop

'        If bFound = True Then
'            aReturn = aTotalValue
'        Else
'            'If no App Data found pass string "False"
'            Erase aTotalValue
'            ReDim aTotalValue(1)

'            aTotalValue(0) = ""
'            aReturn = aTotalValue
'        End If

'        GoTo ExitFunction
'Err:
'        If Err.Number = -2147218390 Then 'EOF (no App Data Found)
'            Erase aTotalValue
'            ReDim aTotalValue(1)

'            aTotalValue(0) = ""
'            aReturn = aTotalValue

'            GoTo ExitFunction
'        ElseIf Err.Number = 28 Then 'Out of stack space
'            ReDim aTotalValue(1)

'            aTotalValue(0) = ""
'            aReturn = aTotalValue
'            GoTo ExitFunction
'        Else
'            MsgBox(Err.Number & ": " & Err.Description, vbCritical, JEGCore.MsgBoxCaption)
'            Resume Next
'        End If
'ExitFunction:
'        Return aReturn(0)
'    End Function

'End Class
