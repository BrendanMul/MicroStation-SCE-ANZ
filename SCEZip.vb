'Imports JMECore.JEGCore
'Imports JMECore.SCEFile
'Imports System.IO
''Imports Shell32

'Public Class SCEZip

'    Shared Function ZipFileCount(ByVal sZipFileName As String) As Integer
'        'Return number of files in specified zip file
'        Dim strmZipInputStream As ICSharpCode.SharpZipLib.Zip.ZipInputStream
'        Dim i As Integer = 0

'        strmZipInputStream = New ICSharpCode.SharpZipLib.Zip.ZipInputStream(System.IO.File.OpenRead(sZipFileName))

'        Dim objZipEntry As ICSharpCode.SharpZipLib.Zip.ZipEntry

'        objZipEntry = strmZipInputStream.GetNextEntry
'        While Not (objZipEntry Is Nothing)
'            i = i + 1
'            objZipEntry = strmZipInputStream.GetNextEntry
'        End While
'        strmZipInputStream.Close()
'        Return i
'    End Function

'    Shared Function ZipFileExists(ByVal sZipFileName As String, ByVal sSearchFileName As String) As Boolean
'        'Return number of files in specified zip file
'        Dim strmZipInputStream As ICSharpCode.SharpZipLib.Zip.ZipInputStream
'        Dim bReturn As Boolean = False

'        strmZipInputStream = New ICSharpCode.SharpZipLib.Zip.ZipInputStream(System.IO.File.OpenRead(sZipFileName))

'        Dim objZipEntry As ICSharpCode.SharpZipLib.Zip.ZipEntry

'        objZipEntry = strmZipInputStream.GetNextEntry
'        While Not (objZipEntry Is Nothing)
'            If InStr(UCase(objZipEntry.Name), UCase(sSearchFileName)) > 0 Then
'                bReturn = True
'                Exit While
'            End If
'            objZipEntry = strmZipInputStream.GetNextEntry
'        End While
'        strmZipInputStream.Close()
'        Return bReturn
'    End Function

'    Shared Sub Unzip(ByVal sFileToExtract As String, ByVal sExtractionPath As String, Optional ByVal bDatabase As Boolean = False, Optional ByVal bZip As Boolean = False)
'        'Unzip contents
'        Dim strmZipInputStream As ICSharpCode.SharpZipLib.Zip.ZipInputStream
'        Dim bUnzipFile As Boolean = False
'        Dim i As Integer = 0

'        strmZipInputStream = New ICSharpCode.SharpZipLib.Zip.ZipInputStream(System.IO.File.OpenRead(sFileToExtract))

'        Dim objZipEntry As ICSharpCode.SharpZipLib.Zip.ZipEntry

'        'If frmCatalogue.TSProgressBar.Visible = True Then
'        '    frmCatalogue.TSProgressBar.ProgressBar.Maximum = ZipFileCount(sFileToExtract)
'        '    i = 0
'        '    frmCatalogue.TSProgressBar.ProgressBar.Value = i
'        'End If
'        objZipEntry = strmZipInputStream.GetNextEntry
'        While Not (objZipEntry Is Nothing)
'            ' Determine directory and filename
'            Dim sDirectoryName As String = Path.GetDirectoryName(objZipEntry.Name)
'            Dim sFileName As String = Path.GetFileName(objZipEntry.Name)
'            'If frmCatalogue.TSProgressBar.Visible = True Then
'            '    i = i + 1
'            '    frmCatalogue.TSProgressBar.ProgressBar.Value = i
'            'End If
'            ' Create directories required for this file
'            If bDatabase = False And bZip = False Then Directory.CreateDirectory(sExtractionPath & "\" & sDirectoryName)
'            If (sFileName <> String.Empty) Then
'                If bDatabase = True Then
'                    If FileParse(UCase(objZipEntry.Name), FileParser.Extension) = "ACCDB" Then
'                        Dim strmWriter As FileStream = System.IO.File.Create(sExtractionPath & "\" & FileParse(Replace(objZipEntry.Name, "/", "\"), FileParser.FullFileName))
'                        Dim nSize As Integer = 2048
'                        Dim bData(nSize) As Byte

'                        While (True)
'                            nSize = strmZipInputStream.Read(bData, 0, bData.Length)
'                            If (nSize > 0) Then
'                                strmWriter.Write(bData, 0, nSize)
'                            Else
'                                Exit While
'                            End If
'                        End While
'                        strmWriter.Close()
'                    ElseIf FileParse(UCase(objZipEntry.Name), FileParser.FullFileName) = FileParse(UCase(sFileToExtract), FileParser.FileName) & ".JPG" Then
'                        If My.Computer.FileSystem.FileExists(sExtractionPath & "\" & FileParse(objZipEntry.Name, FileParser.FullFileName)) = False Then
'                            Dim strmWriter As FileStream = System.IO.File.Create(sExtractionPath & "\" & FileParse(objZipEntry.Name, FileParser.FullFileName))
'                            Dim nSize As Integer = 2048
'                            Dim bData(nSize) As Byte

'                            While (True)
'                                nSize = strmZipInputStream.Read(bData, 0, bData.Length)
'                                If (nSize > 0) Then
'                                    strmWriter.Write(bData, 0, nSize)
'                                Else
'                                    Exit While
'                                End If
'                            End While
'                            strmWriter.Close()
'                        End If
'                    End If
'                ElseIf bZip = True Then
'                    'unzip cell catalogue
'                    If UCase(sFileName) = "CELL CATALOGUE.ZIP" Then
'                        Dim strmWriter As FileStream = System.IO.File.Create(sExtractionPath & "\" & FileParse(Replace(objZipEntry.Name, "/", "\"), FileParser.FullFileName))
'                        Dim nSize As Integer = 2048
'                        Dim bData(nSize) As Byte

'                        While (True)
'                            nSize = strmZipInputStream.Read(bData, 0, bData.Length)
'                            If (nSize > 0) Then
'                                strmWriter.Write(bData, 0, nSize)
'                            Else
'                                Exit While
'                            End If
'                        End While
'                        strmWriter.Close()
'                    End If
'                Else
'                    If Not My.Computer.FileSystem.FileExists(sExtractionPath & "\" & objZipEntry.Name) Then
'                        Dim strmWriter As FileStream = System.IO.File.Create(sExtractionPath & "\" & objZipEntry.Name)
'                        Dim nSize As Integer = 2048
'                        Dim bData(nSize) As Byte

'                        While (True)
'                            nSize = strmZipInputStream.Read(bData, 0, bData.Length)
'                            If (nSize > 0) Then
'                                strmWriter.Write(bData, 0, nSize)
'                            Else
'                                Exit While
'                            End If
'                        End While
'                        strmWriter.Close()
'                    End If
'                End If
'            End If
'            objZipEntry = strmZipInputStream.GetNextEntry
'        End While
'        strmZipInputStream.Close()
'    End Sub

'    'Shared Sub UnzipClient(ByVal sRegion As String, ByVal sClientName As String)
'    '    'Unzip client data ready for use
'    '    Dim bUnzip As Boolean = False
'    '    Dim sFileDateTime As String = ""
'    '    Dim sFileDateTime1 As String = ""
'    '    Dim sVersionFile As String = JEGCore.ClientPath & "\" & sRegion & "\" & sClientName & "\ClientVersion.ini"
'    '    Dim sLine As String = ""
'    '    Dim sClientZip As String = RegionPath() & "\" & sRegion & "\" & sClientName & ".zip"

'    '    'Determine if client needs to be unzipped
'    '    If My.Computer.FileSystem.FileExists(sClientZip) = True Then
'    '        If My.Computer.FileSystem.DirectoryExists(ClientPath() & "\" & sRegion & "\" & sClientName) = False Then
'    '            bUnzip = True
'    '        Else
'    '            If My.Computer.FileSystem.FileExists(sVersionFile) = True Then
'    '                FileOpen(1, sVersionFile, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
'    '                Do While Not EOF(1)
'    '                    sFileDateTime = Trim(LineInput(1))
'    '                    Exit Do
'    '                Loop
'    '                FileClose(1)
'    '                sFileDateTime1 = FileDateTime(sClientZip)
'    '                If sFileDateTime1 <> sFileDateTime Then
'    '                    bUnzip = True
'    '                End If
'    '            Else
'    '                bUnzip = True
'    '            End If
'    '        End If
'    '    End If

'    '    'unzip client
'    '    If bUnzip = True Then
'    '        My.Computer.FileSystem.CreateDirectory(ClientPath() & "\" & sRegion & "\" & sClientName)
'    '        'unzip client data
'    '        Unzip(sClientZip, ClientPath() & "\" & sRegion & "\" & sClientName)
'    '        'unzip new database
'    '        Unzip(sClientZip, ClientPath() & "\Database\" & sRegion, True)
'    '        'Create zip datetime stamp
'    '        FileOpen(1, sVersionFile, OpenMode.Output, OpenAccess.Write, OpenShare.LockWrite)
'    '        PrintLine(1, FileDateTime(JEGCore.RegionPath & "\" & sRegion & "\" & sClientName & ".zip"))
'    '        FileClose(1)
'    '    End If

'    '    'Check for additional clients to unload
'    '    If My.Computer.FileSystem.FileExists(ClientPath() & "\" & sRegion & "\" & sClientName & "\Config\" & sClientName & ".ini") = True Then
'    '        Dim sLoadClient As String = ""
'    '        Dim sLoadRegion As String = ""

'    '        FileOpen(1, ClientPath() & "\" & sRegion & "\" & sClientName & "\Config\" & sClientName & ".ini", OpenMode.Input, OpenAccess.Read, OpenShare.Default)
'    '        Do While Not EOF(1)
'    '            sLine = LineInput(1)
'    '            If InStr(UCase(sLine), "LOADREGION") > 0 Then
'    '                sLine = Replace(UCase(sLine), "LOADREGION", "")
'    '                sLine = Replace(UCase(sLine), "=", "")
'    '                sLoadRegion = Trim(sLine)
'    '            ElseIf InStr(UCase(sLine), "LOADCLIENT") > 0 Then
'    '                sLine = Replace(UCase(sLine), "LOADCLIENT", "")
'    '                sLine = Replace(UCase(sLine), "=", "")
'    '                sLoadClient = Trim(sLine)
'    '            End If
'    '            If sLoadClient <> "" And sLoadRegion <> "" Then
'    '                If My.Computer.FileSystem.FileExists(RegionPath() & "\" & sLoadRegion & "\" & sLoadClient & ".zip") = True Then
'    '                    'If FolderExists(ClientPath() & "\" & sLoadRegion & "\" & sLoadClient) = False Then
'    '                    My.Computer.FileSystem.CreateDirectory(ClientPath() & "\" & sLoadRegion & "\" & sLoadClient)
'    '                    Unzip(RegionPath() & "\" & sLoadRegion & "\" & sLoadClient & ".zip", ClientPath() & "\" & sLoadRegion & "\" & sLoadClient)
'    '                    'End If
'    '                End If
'    '            End If
'    '        Loop
'    '        FileClose(1)
'    '    End If
'    'End Sub

'    ''    Shared Sub UnzipDatabases()
'    ''        'Unzip all client databases
'    ''        Dim aRegions As Array
'    ''        Dim aClients As Array
'    ''        Dim i As Integer
'    ''        Dim j As Integer

'    ''        aRegions = GetFolders(RegionPath)
'    ''        For i = 0 To UBound(aRegions)
'    ''            aClients = GetFiles(RegionPath() & "\" & aRegions(i), "zip")
'    ''            If SCEVB.IsNullArray(aClients) = True Then
'    ''                My.Computer.FileSystem.DeleteDirectory(RegionPath() & "\" & aRegions(i), FileIO.DeleteDirectoryOption.DeleteAllContents)
'    ''                GoTo NextRegion
'    ''            End If

'    ''            If My.Computer.FileSystem.DirectoryExists(ClientPath() & "\Database\" & aRegions(i)) = False Then
'    ''                My.Computer.FileSystem.CreateDirectory(ClientPath() & "\Database\" & aRegions(i))
'    ''            End If

'    ''            For j = 0 To UBound(aClients)
'    ''                Unzip(RegionPath() & "\" & aRegions(i) & "\" & aClients(j), ClientPath() & "\Database\" & aRegions(i), True)
'    ''            Next
'    ''NextRegion:
'    ''        Next
'    ''    End Sub



'    'uses Shell32 to create the zip file
'    'Shared Sub CreateZip(ByVal sZipPath As String, ByRef sZipFile As String)
'    '    '1) Lets create an empty Zip File .
'    '    'The following data represents an empty zip file .

'    '    If My.Computer.FileSystem.FileExists(sZipFile) = True Then
'    '        My.Computer.FileSystem.DeleteFile(sZipFile)
'    '    End If

'    '    Dim startBuffer() As Byte = {80, 75, 5, 6, 0, 0, 0, 0, 0, 0, 0, _
'    '                                 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}
'    '    ' Data for an empty zip file .
'    '    FileIO.FileSystem.WriteAllBytes(sZipFile, startBuffer, False)

'    '    'We have successfully made the empty zip file .

'    '    '2) Use the Shell32 to zip your files .
'    '    ' Declare new shell class
'    '    Dim sc As New Shell32.Shell()
'    '    'Declare the folder which contains the files you want to zip .
'    '    Dim input As Shell32.Folder = sc.NameSpace(sZipPath)
'    '    'Declare  your created empty zip file as folder  .
'    '    Dim output As Shell32.Folder = sc.NameSpace(sZipFile)
'    '    'Copy the files into the empty zip file using the CopyHere command .
'    '    output.CopyHere(input.Items, 4)
'    'End Sub

'    'uses sharpziplib to create the zip file
'    Shared Function CreateZip(ByVal sZipPath As String, ByRef sZipFile As String, Optional ByVal bIncludeSubDirectories As Boolean = False) As String
'        'Create zip file
'        'CreateZip("D:\Documents and Settings\SKMCAD\MSTN\User", "D:\Documents and Settings\SKMCAD\MSTN\User.zip")
'        Dim sFileName As String = ""
'        Dim sDirectory As String
'        Dim bSuccess As Boolean = False
'        Dim cDirectoryEntries As Collection = New Collection
'        Dim objCrc32 As New ICSharpCode.SharpZipLib.Checksums.Crc32
'        Dim strmZipOutputStream As ICSharpCode.SharpZipLib.Zip.ZipOutputStream

'        ' Remove zip file (if it exists)
'        If System.IO.File.Exists(sFileName) Then
'            System.IO.File.Delete(sFileName)
'        End If

'        ' Remove zip file (if it exists)
'        If System.IO.File.Exists(sZipFile) Then
'            System.IO.File.Delete(sZipFile)
'        End If

'        ' Get all files beneath directory specified
'        sDirectory = sZipPath
'        GetDirectoryFiles(sZipPath, cDirectoryEntries, bIncludeSubDirectories)

'        ' Create destination file
'        strmZipOutputStream = New ICSharpCode.SharpZipLib.Zip.ZipOutputStream(System.IO.File.Create(sZipFile))

'        ' Set compression level (0 = no compression, 9 = maximum compression)
'        strmZipOutputStream.SetLevel(5)

'        ' Loop through files and compress
'        For Each strFile As String In cDirectoryEntries
'            Dim strmFile As FileStream = System.IO.File.OpenRead(strFile)
'            Dim abyBuffer(strmFile.Length - 1) As Byte

'            strmFile.Read(abyBuffer, 0, abyBuffer.Length)
'            Dim sFormattedName As String = strFile.Substring(sDirectory.Length + 1)

'            Dim objZipEntry As ICSharpCode.SharpZipLib.Zip.ZipEntry = New ICSharpCode.SharpZipLib.Zip.ZipEntry(sFormattedName)

'            objZipEntry.DateTime = DateTime.Now
'            objZipEntry.Size = strmFile.Length
'            strmFile.Close()
'            objCrc32.Reset()
'            objCrc32.Update(abyBuffer)
'            objZipEntry.Crc = objCrc32.Value
'            strmZipOutputStream.PutNextEntry(objZipEntry)
'            strmZipOutputStream.Write(abyBuffer, 0, abyBuffer.Length)
'        Next

'        strmZipOutputStream.Finish()
'        strmZipOutputStream.Close()

'        Return sZipFile
'    End Function

'    'Sub UnZip()
'    '    Dim sc As New Shell32.Shell()
'    '    ''UPDATE !!
'    '    'Create directory in which you will unzip your files .
'    '    IO.Directory.CreateDirectory("D:\extractedFiles")
'    '    'Declare the folder where the files will be extracted
'    '    Dim output As Shell32.Folder = sc.NameSpace("D:\extractedFiles")
'    '    'Declare your input zip file as folder  .
'    '    Dim input As Shell32.Folder = sc.NameSpace("d:\myzip.zip")
'    '    'Extract the files from the zip file using the CopyHere command .
'    '    output.CopyHere(input.Items, 4)

'    'End Sub

'    '    Shared Sub ZipClients(Optional ByVal sRegion As String = "", Optional ByVal sClientName As String = "")
'    '        'zip all clients in all regions
'    '        ''Dim aRegions As Array
'    '        ''Dim aClients As Array
'    '        'Dim i As Integer
'    '        'Dim j As Integer
'    '        Dim sRegions As String
'    '        Dim sClients As String

'    '        For Each sRegions In My.Computer.FileSystem.GetDirectories(RegionPath)
'    '            If sRegion <> "" Then
'    '                If UCase(sRegion) <> UCase(sRegions) Then
'    '                    GoTo NextRegion
'    '                End If
'    '            End If
'    '            For Each sClients In My.Computer.FileSystem.GetDirectories(RegionPath() & "\" & sRegions)
'    '                ''aClients = GetFolders(RegionPath() & "\" & aRegions(i))
'    '                'For j = 0 To UBound(aClients)
'    '                If sClientName <> "" Then
'    '                    If UCase(sClientName) <> UCase(sClients) Then
'    '                        GoTo NextClient
'    '                    End If
'    '                End If
'    '                'ShowWait("Zipping " & aClients(j) & " client data...")
'    '                SCEZip.CreateZip(RegionPath() & "\" & sRegions & "\" & sClients, RegionPath() & "\" & sRegions & "\" & sClients & ".zip")
'    '                'HideWait()
'    'NextClient:
'    '            Next
'    'NextRegion:
'    '        Next
'    '        ''        aRegions = GetFolders(RegionPath)
'    '        ''        For i = 0 To UBound(aRegions)
'    '        ''            If sRegion <> "" Then
'    '        ''                If UCase(sRegion) <> UCase(aRegions(i)) Then
'    '        ''                    GoTo NextRegion
'    '        ''                End If
'    '        ''            End If
'    '        ''            aClients = GetFolders(RegionPath() & "\" & aRegions(i))
'    '        ''            For j = 0 To UBound(aClients)
'    '        ''                If sClientName <> "" Then
'    '        ''                    If UCase(sClientName) <> UCase(aClients(j)) Then
'    '        ''                        GoTo NextClient
'    '        ''                    End If
'    '        ''                End If
'    '        ''                'ShowWait("Zipping " & aClients(j) & " client data...")
'    '        ''                SCEZip.CreateZip(RegionPath() & "\" & aRegions(i) & "\" & aClients(j), RegionPath() & "\" & aRegions(i) & "\" & aClients(j) & ".zip")
'    '        ''                'HideWait()
'    '        ''NextClient:
'    '        ''            Next
'    '        ''NextRegion:
'    '        ''        Next
'    '        If sClientName = "" Then
'    '            MsgBox("Clients have been zipped", MsgBoxStyle.OkOnly, JEGCore.MsgBoxCaption)
'    '        Else
'    '            MsgBox("Client " & sClientName & " data has been zipped", MsgBoxStyle.OkOnly, JEGCore.MsgBoxCaption)
'    '        End If
'    '    End Sub

'    'Shared Sub UnzipClientBuild(ByVal sRegion As String, ByVal sClientName As String)
'    '    'Unzip client data ready for use
'    '    If My.Computer.FileSystem.FileExists(RegionPath() & "\" & sRegion & "\" & sClientName & ".zip") = True Then
'    '        If My.Computer.FileSystem.DirectoryExists(ClientPath() & "\" & sRegion & "\" & sClientName) = False Then
'    '            'ShowWait("Extracting " & sClientName & " client data...")
'    '            My.Computer.FileSystem.CreateDirectory(ClientPath() & "\" & sRegion & "\" & sClientName)
'    '            UnZip(RegionPath() & "\" & sRegion & "\" & sClientName & ".zip", ClientPath() & "\" & sRegion & "\" & sClientName)
'    '            'HideWait()
'    '        End If
'    '    End If
'    '    'Check for additional clients to unload
'    '    If My.Computer.FileSystem.FileExists(ClientPath() & "\" & sRegion & "\" & sClientName & "\Config\" & sClientName & ".ini") = True Then
'    '        Dim sLine As String = ""
'    '        Dim sLoadClient As String = ""
'    '        Dim sLoadRegion As String = ""

'    '        FileOpen(1, ClientPath() & "\" & sRegion & "\" & sClientName & "\Config\" & sClientName & ".ini", OpenMode.Input, OpenAccess.Read, OpenShare.Default)
'    '        Do While Not EOF(1)
'    '            sLine = LineInput(1)
'    '            If InStr(UCase(sLine), "LOADREGION") > 0 Then
'    '                sLine = Replace(UCase(sLine), "LOADREGION", "")
'    '                sLine = Replace(UCase(sLine), "=", "")
'    '                sLoadRegion = Trim(sLine)
'    '            ElseIf InStr(UCase(sLine), "LOADCLIENT") > 0 Then
'    '                sLine = Replace(UCase(sLine), "LOADCLIENT", "")
'    '                sLine = Replace(UCase(sLine), "=", "")
'    '                sLoadClient = Trim(sLine)
'    '            End If
'    '            If sLoadClient <> "" And sLoadRegion <> "" Then
'    '                If My.Computer.FileSystem.FileExists(RegionPath() & "\" & "\" & sLoadRegion & "\" & sLoadClient & ".zip") = True Then
'    '                    If My.Computer.FileSystem.DirectoryExists(ClientPath() & "\" & sLoadRegion & "\" & sLoadClient) = False Then
'    '                        'ShowWait("Extracting " & sClientName & " client data...")
'    '                        My.Computer.FileSystem.CreateDirectory(ClientPath() & "\" & sLoadRegion & "\" & sLoadClient)
'    '                        UnZip(RegionPath() & "\" & "\" & sLoadRegion & "\" & sLoadClient & ".zip", ClientPath() & "\" & sLoadRegion & "\" & sLoadClient)
'    '                        'HideWait()
'    '                    End If
'    '                End If
'    '            End If
'    '        Loop
'    '        FileClose(1)
'    '    End If
'    'End Sub

'End Class
