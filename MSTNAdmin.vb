Imports VBIDE
Imports MicroStationSCE.SCEFile

<ComClass(MSTNAdmin.ClassId, MSTNAdmin.InterfaceId, MSTNAdmin.EventsId)> _
Public Class MSTNAdmin

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "f3da8976-2993-4040-abfe-6a33d8667b11"
    Public Const InterfaceId As String = "0142f2d4-69f4-4188-a26e-94c8e6a340f9"
    Public Const EventsId As String = "53601eca-df67-4b88-afc4-72b948f8a48a"
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

    Sub ExportVBComponents()

        'Exports all projects vb components (Modules, Classes, Forms) as separate files

        Dim oProject As VBProject
        Dim oProjects As VBProjects
        Dim i As Long
        Dim lCount As Long
        Dim sFolder As String = ""
        Dim sExtension As String
        Dim sMsgBoxCaption As String
        Dim sDrive As String
        Dim sPath As String
        Dim oMsg As MsgBoxResult

        sMsgBoxCaption = "DWS - Export VBA Components"

EnterPath:
        sPath = InputBox("Enter path to create folders for VB Component export.", sMsgBoxCaption, "d:\cadwork")

        If My.Computer.FileSystem.DirectoryExists(sPath) = False Then

            oMsg = MsgBox("Create folder: " & sPath & "?", 35, sMsgBoxCaption)

            If oMsg = vbYes Then

                My.Computer.FileSystem.CreateDirectory(sPath)

            Else

                GoTo EnterPath

            End If

        End If
EnterDrive:
        sDrive = InputBox("Enter drive letter where the MVBA's exist for this export", sMsgBoxCaption, "H")

        If Trim(sDrive) = "" Then

            MsgBox("A drive letter must be entered", vbExclamation, sMsgBoxCaption)
            GoTo EnterDrive

        End If

        sDrive = UCase(sDrive)

        lCount = 0
        oProjects = oMSTN.VBE.VBProjects

        For Each oProject In oProjects

            lCount = lCount + 1
            'pjrxxx - can show form work in com class
            'frmStatus.ShowStatusDialog(sMsgBoxCaption, "Exporting Project: " & oProject.Name & "...", True, oProjects.Count, lCount)

            On Error GoTo SkipProject 'VB Project hasn't been opened and has password security

            'pjr - commented due to non resolved vbext_pp_locked
            'If oProject.Protection = vbext_pp_locked Then
            '    FileOpen(1, sPath & "\" & "Export.log", OpenMode.Append, OpenAccess.Write, OpenShare.LockWrite)
            '    PrintLine(1, "Project locked: " & oProject.Name)
            '    FileClose(1)
            '    GoTo SkipProject
            'End If

            If oProject.VBComponents.Count = 0 Then

                FileOpen(1, sPath & "\" & "Export.log", OpenMode.Append, OpenAccess.Write, OpenShare.LockWrite)
                PrintLine(1, "No Components: " & oProject.Name)
                FileClose(1)
                GoTo SkipProject

            End If

            For i = 1 To oProject.VBComponents.Count

                Select Case oProject.VBComponents.Item(i).Type

                    Case 1 'Module

                        sExtension = ".bas"

                    Case 2 'Class

                        sExtension = ".cls"

                    Case 3 'Form

                        sExtension = ".frm"

                End Select

                sFolder = sPath & "\" & sDrive & " - " & oProject.Name

                If My.Computer.FileSystem.DirectoryExists(sFolder) = False Then My.Computer.FileSystem.CreateDirectory(sFolder)

                oProject.VBComponents.Item(i).Export(sFolder & "\" & oProject.VBComponents.Item(i).Name & sExtension)

            Next

            'Save References
            If oProject.References.Count > 0 Then
                FileOpen(1, sFolder & "\" & oProject.Name & ".ref", OpenMode.Output, OpenAccess.Write, OpenShare.LockWrite)
                For i = 1 To oProject.References.Count
                    PrintLine(1, oProject.References.Item(i).FullPath)
                Next
                FileClose(1)
            End If

            'Save the Description
            If oProject.Description <> "" Then
                FileOpen(1, sFolder & "\" & oProject.Name & ".dsc", OpenMode.Output, OpenAccess.Write, OpenShare.LockWrite)
                PrintLine(1, oProject.Description)
                FileClose(1)
            End If
SkipProject:
            On Error GoTo 0
        Next
        'HideStatusDialog()
        MsgBox("VB Components have been exported", vbInformation, sMsgBoxCaption)
    End Sub

    Sub ImportVBComponents()
        'Imports components (Modules, Classes, Forms) from specified path into an empty Project (prompts with available Projects)
        Dim oProject As VBProject
        Dim oProjects As VBProjects
        Dim RetVal
        Dim bProjectFound As Boolean
        Dim sPath As String
        Dim vFiles As Object
        Dim i As Long
        Dim lType As Long
        Dim sMsgBoxCaption As String
        Dim bRefFound As Boolean
        Dim sLine As String

        sMsgBoxCaption = "DWS - Import VBA Components"
SelectPath:
        sPath = InputBox("Enter the path where the VB Components are located.", sMsgBoxCaption, "H:\CompareXM")

        'Check component path exists
        If My.Computer.FileSystem.DirectoryExists(sPath) = False Then
            MsgBox(sPath & " path does not exist please enter a valid path", vbExclamation, sMsgBoxCaption)
            GoTo SelectPath
        End If

        oProjects = oMSTN.VBE.VBProjects

        'Get empty project
        For Each oProject In oProjects
            'If oProject.Protection = vbext_pp_locked Then GoTo NextProject
            If oProject.VBComponents.Count = 1 Then
                RetVal = MsgBox("Would you like to import the components into the " & oProject.Name & " Project?", 35, sMsgBoxCaption)
                If RetVal = vbYes Then
                    bProjectFound = True

                    'Get all components
                    vFiles = GetFiles(sPath, "*", GetFileType.all)

                    'Import components
                    For i = 0 To UBound(vFiles)

                        Select Case FileParse(LCase(vFiles(i)), FileParser.Extension)
                            Case "bas" 'Module
                                lType = 1
                            Case "cls" 'Class
                                lType = 2
                            Case "frm" 'Form
                                lType = 3
                            Case Else
                                GoTo NextFile
                        End Select
                        oProject.VBComponents.Import(sPath & "\" & vFiles(i))
NextFile:
                    Next 'vFiles

                    'Remove Module 1
                    If oProject.VBComponents.Count > 1 Then
                        For i = 1 To oProject.VBComponents.Count
                            If UCase(oProject.VBComponents.Item(i).Name) = "MODULE1" Then
                                oProject.VBComponents.Remove(oProject.VBComponents.Item(i))
                                Exit For
                            End If
                        Next
                    End If

                    'check for any references
                    vFiles = GetFiles(sPath, "*.ref", GetFileType.all, False)

                    If vFiles(0) <> "" Then
                        FileOpen(1, sPath & "\" & vFiles(0), OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
                        Do While Not EOF(1)
                            bRefFound = False
                            sLine = LineInput(1)

                            For i = 1 To oProject.References.Count
                                If UCase(sLine) = UCase(oProject.References.Item(i).FullPath) Then
                                    bRefFound = True
                                    Exit For
                                End If
                            Next

                            'Add reference
                            If bRefFound = False Then
                                If FileParse(UCase(sLine), FileParser.Extension) = "MVBA" Then 'MicroStation VBA
                                    If My.Computer.FileSystem.FileExists(sPath & "\" & FileParse(sLine, FileParser.FullFileName)) = True Then
                                        oProject.References.AddFromFile(sPath & "\" & FileParse(sLine, FileParser.FullFileName))
                                    Else
                                        FileOpen(2, sPath & "\" & "Import.log", OpenMode.Append, OpenAccess.Write, OpenShare.LockWrite)
                                        PrintLine(2, "Reference not found: " & sPath & "\" & FileParse(sLine, FileParser.FullFileName) & " (old reference: " & sLine & ")")
                                        FileClose(2)
                                    End If
                                Else
                                    If My.Computer.FileSystem.FileExists(sLine) = True Then
                                        oProject.References.AddFromFile(sLine)
                                    Else
                                        FileOpen(2, sPath & "\" & "Import.log", OpenMode.Append, OpenAccess.Write, OpenShare.LockWrite)
                                        PrintLine(2, "Reference not found: " & sLine)
                                        FileClose(2)
                                    End If
                                End If
                            End If
                        Loop
                        FileClose(1)
                    End If 'check for references

                    'Check for saved description
                    vFiles = GetFiles(sPath, "*.dsc", GetFileType.all, False)

                    If vFiles(0) <> "" Then
                        FileOpen(1, sPath & "\" & vFiles(0), OpenMode.Input, OpenAccess.Write, OpenShare.LockWrite)
                        Do While Not EOF(1)
                            sLine = LineInput(1)
                            oProject.Description = sLine
                            Exit Do
                        Loop
                        FileClose(1)
                    End If
                    Exit For
                End If 'use current project to import to
            End If '1 component
NextProject:
        Next 'Projects

        If bProjectFound = True Then
            MsgBox("Import is now complete", vbInformation, sMsgBoxCaption)
        Else
            MsgBox("No empty Projects were found, please create a new Project", vbExclamation, sMsgBoxCaption)
        End If
    End Sub

End Class


