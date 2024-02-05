''Public Class SCEFTP

''    Public Shared LocalFTPPath As String = SCECore.LocalTempPath & "\FTP"
''    Public Shared LocalTempFTPPath As String = SCECore.LocalTempPath & "\FTPTemp"

''    Public Enum tFTPAction
''        Upload = 0
''        Download = 1
''        List = 2
''        Delete = 3
''    End Enum

''    Shared Sub FTPOperation(ByVal oAction As tFTPAction, ByVal sPath As String, Optional ByVal sFileName As String = "", _
''        Optional ByVal sListFileName As String = "", Optional ByVal sLogFileName As String = "", Optional ByVal sDownloadPath As String = "", _
''        Optional ByVal sUploadPath As String = "", Optional ByVal sZipFile As String = "")
''        'FTP operations
''        Dim sScriptFileName As String = SCEFTP.LocalFTPPath & "\FTPOperation.txt"
''        Dim ProcessInfo As New System.Diagnostics.ProcessStartInfo
''        Dim aSplit() As String
''        Dim i As Integer
''        Dim sTempDownloadPath As String = ""

''        My.Computer.FileSystem.CreateDirectory(LocalFTPPath)

''        'Create batch and script files
''        FileOpen(1, LocalFTPPath & "\FTPOperation.bat", OpenMode.Output, OpenAccess.Write, OpenShare.LockWrite)
''        FileOpen(2, sScriptFileName, OpenMode.Output, OpenAccess.Write, OpenShare.LockWrite)
''        PrintLine(1, SCEFile.FileParse(sUploadPath, SCEFile.FileParser.Drive) & ":")
''        If sDownloadPath = "" Then
''            If oAction = tFTPAction.Upload Then
''                If InStr(sUploadPath, ":\") > 0 Then
''                    PrintLine(1, "cd\" & Chr(34) & Right(sUploadPath, Len(sUploadPath) - 3) & Chr(34))
''                Else
''                    PrintLine(1, "cd\" & Chr(34) & sUploadPath & Chr(34))
''                End If
''            Else
''                PrintLine(1, "cd\" & Chr(34) & Right(LocalFTPPath, Len(LocalFTPPath) - 3) & Chr(34))
''            End If
''        Else
''            My.Computer.FileSystem.CreateDirectory(sDownloadPath)
''            If InStr(sDownloadPath, ":\") > 0 Then
''                PrintLine(1, Left(sDownloadPath, 2))
''                sTempDownloadPath = Right(sDownloadPath, Len(sDownloadPath) - 3)
''            Else
''                sTempDownloadPath = sDownloadPath
''            End If
''            PrintLine(1, "cd\" & Chr(34) & sTempDownloadPath & Chr(34))
''        End If
''        If sLogFileName = "" Then
''            PrintLine(1, "ftp -v -i -s:" & Chr(34) & sScriptFileName & Chr(34))
''        Else
''            PrintLine(1, "ftp -v -i -s:" & Chr(34) & sScriptFileName & Chr(34) & " >" & sLogFileName)
''        End If

''        'open ftp
''        PrintLine(2, "open " & Activity.FTPPath)
''        PrintLine(2, Activity.FTPUserName)
''        PrintLine(2, Activity.FTPPassword)

''        If oAction = tFTPAction.Upload Then
''            aSplit = Split(sPath, "\")
''            For i = 0 To UBound(aSplit)
''                PrintLine(2, "mkdir " & Chr(34) & aSplit(i) & Chr(34))
''                PrintLine(2, "cd " & Chr(34) & aSplit(i) & Chr(34))
''            Next
''        Else
''            PrintLine(2, "cd " & Chr(34) & sPath & Chr(34))
''        End If

''        'upload action
''        Select Case oAction
''            Case tFTPAction.Download
''                aSplit = Split(sFileName, ",")
''                For i = 0 To UBound(aSplit)
''                    PrintLine(2, "get " & Chr(34) & aSplit(i) & Chr(34))
''                Next
''            Case tFTPAction.List
''                PrintLine(2, "ls -l " & Chr(34) & sListFileName & Chr(34))
''            Case tFTPAction.Upload
''                aSplit = Split(sFileName, ",")
''                For i = 0 To UBound(aSplit)
''                    PrintLine(2, "put " & Chr(34) & aSplit(i) & Chr(34))
''                Next
''            Case tFTPAction.Delete
''                aSplit = Split(sFileName, ",")
''                For i = 0 To UBound(aSplit)
''                    PrintLine(2, "delete " & Chr(34) & sZipFile & Chr(34))
''                    PrintLine(2, "cd .. ")
''                    PrintLine(2, "rmdir " & Chr(34) & sPath & Chr(34))
''                Next
''        End Select

''        'close FTP
''        PrintLine(2, "disconnect")
''        PrintLine(2, "bye")
''        PrintLine(1, "exit")
''        FileClose(1)
''        FileClose(2)

''        'Process FTP batch file
''        ProcessInfo.FileName = LocalFTPPath & "\FTPOperation.bat" 'Name of the program
''        ProcessInfo.WindowStyle = ProcessWindowStyle.Hidden
''        System.Diagnostics.Process.Start(ProcessInfo).WaitForExit()  'Let's start the process using the given info
''    End Sub

''    Shared Function FTPVerify(ByVal oAction As tFTPAction, ByVal sPath As String) As Boolean
''        'Verify files exist in FTP
''        Dim bVerify As Boolean = False
''        Dim i As Integer
''        Dim j As Integer
''        Dim aDates() As String = Nothing
''        Dim aTimes() As String = Nothing
''        Dim aFiles() As String = Nothing
''        Dim aFolders() As String = Nothing
''        Dim sFileName As String = ""
''        Dim aLocalFiles() As String
''        Dim sFolder As String = ""
''        Dim sFile As String

''        FTPGetFolderContents(sPath, LocalFTPPath & "\FTPList.log")
''        If My.Computer.FileSystem.FileExists(LocalFTPPath & "\FTPList.log") = True Then
''            GetFileDetailList(LocalFTPPath & "\FTPList.log", aDates, aTimes, aFiles, aFolders)
''        End If

''        Select Case oAction
''            Case tFTPAction.Download
''                sFolder = "Download"
''                bVerify = Not SCEVB.IsNullArray(aFiles)
''                ''aLocalFiles = SCEFile.GetFiles(LocalTempFTPPath, "zip", SCEFile.GetFileType.all)

''                For Each sFile In My.Computer.FileSystem.GetFiles(LocalTempFTPPath, FileIO.SearchOption.SearchTopLevelOnly, "zip")
''                    sFile = SCEFile.FileParse(sFile, SCEFile.FileParser.FullFileName)
''                    If bVerify = True Then
''                        For i = 0 To UBound(aFiles)
''                            If UCase(sFile) = UCase(aFiles(i)) Then
''                                GoTo NextDLocalFile
''                            End If
''                            If sFileName = "" Then
''                                sFileName = aFiles(i)
''                            Else
''                                sFileName = sFileName & "," & aFiles(i)
''                            End If
''                            bVerify = False
''NextDLocalFile:
''                        Next
''                        If bVerify = False Then
''                            AddLog(Activity.CurrentOffice, "General", "Error - Some files were not downloaded from the FTP site: " & sFileName)
''                        End If
''                    Else
''                        AddLog(Activity.CurrentOffice, "General", "No files were downloaded from the FTP path: " & sPath)
''                    End If
''                Next
''                If bVerify = True Then
''                    AddLog(Activity.CurrentOffice, "General", "Files downloaded successfully.")
''                End If
''            Case tFTPAction.Upload
''                bVerify = Not SCEVB.IsNullArray(aFiles)
''                ''aLocalFiles = SCEFile.GetFiles(sLocalTempFTPPath, "zip", GetFileType.All)
''                If bVerify = True Then
''                    ''For i = 0 To UBound(aLocalFiles)
''                    For Each sFile In My.Computer.FileSystem.GetFiles(LocalTempFTPPath, FileIO.SearchOption.SearchTopLevelOnly, "zip")
''                        sFile = SCEFile.FileParse(sFile, SCEFile.FileParser.FullFileName)
''                        For j = 0 To UBound(aFiles)
''                            If UCase(sFile) = UCase(aFiles(j)) Then
''                                GoTo NextULocalFile
''                            End If
''                        Next
''                        If sFileName = "" Then
''                            sFileName = aLocalFiles(i)
''                        Else
''                            sFileName = sFileName & "," & aLocalFiles(i)
''                        End If
''                        bVerify = False
''NextULocalFile:
''                    Next
''                    If bVerify = False Then
''                        AddLog(Activity.CurrentOffice, "General", "Error - Some files were not uploaded to the FTP site: " & sFileName)
''                    End If
''                Else
''                    AddLog(Activity.CurrentOffice, "General", "No files were uploaded to the FTP path: " & sPath)
''                End If
''                If bVerify = True Then
''                    AddLog(Activity.CurrentOffice, "General", "Files uploaded successfully.")
''                End If
''            Case Else
''        End Select
''ExitFunction:
''        Return bVerify
''    End Function

''    Shared Sub FTPUpload(ByVal sPath As String, ByVal sFileName As String, Optional ByVal sLogFileName As String = "", Optional ByVal sUploadPath As String = "")
''        'upload selected file(s) to FTP
''        Dim bVerify As Boolean = False
''        FTPOperation(tFTPAction.Upload, sPath, sFileName, , , , sUploadPath)
''        bVerify = FTPVerify(tFTPAction.Upload, sPath)
''        If bVerify = True Then
''            AddLog(Activity.CurrentOffice, "General", "Successful")
''        Else
''            AddLog(Activity.CurrentOffice, "General", "Verify Failed")
''        End If
''    End Sub

''    Shared Sub FTPDownload(ByVal sPath As String, ByVal sFileName As String, Optional ByVal sLogFileName As String = "", Optional ByVal sDownloadPath As String = "")
''        'download all files from FTP
''        Dim j As Integer = -1
''        Dim aSplit As Array

''        'Download files
''        FTPOperation(tFTPAction.Download, sPath, sFileName, , , sDownloadPath)

''        'Move to correct local location
''        aSplit = Split(sFileName, ",")
''        sFileName = ""
''        sPath = ""
''        sFileName = ""
''ExitSub:
''        AddLog(Activity.CurrentOffice, "General", "Download complete.")
''        Exit Sub
''    End Sub

''    Shared Sub FTPGetFolderContents(ByVal sPath As String, ByVal sFileName As String, Optional ByVal sLogFileName As String = "")
''        'list contents of selected folder from FTP
''        FTPOperation(tFTPAction.List, sPath, , sFileName)
''    End Sub

''    Shared Function GetFileList(ByVal sLogFile As String, ByRef aPaths() As String, ByRef aFiles() As String)
''        'retrieve file list from FTP log
''        Dim sLine As String
''        Dim i As Integer = -1
''        Dim aSplit As Array

''        If My.Computer.FileSystem.FileExists(sLogFile) = True Then
''            FileOpen(2, sLogFile, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
''            Do While Not EOF(2)
''                sLine = Trim(LineInput(2))
''                If sLine <> "" Then
''                    i = i + 1
''                    ReDim Preserve aPaths(i)
''                    ReDim Preserve aFiles(i)

''                    aSplit = Split(sLine, "/")
''                    aPaths(i) = aSplit(0)
''                    aFiles(i) = aSplit(UBound(aSplit))
''                End If
''NextLine:
''            Loop
''            FileClose(2)
''        End If
''        Return aFiles
''    End Function

''    Shared Function GetFileDetailList(ByVal sLogFile As String, ByRef aDates() As String, ByRef aTimes() As String, ByRef aFiles() As String, ByRef aFolders() As String)
''        'retrieve file list from FTP log
''        Dim sLine As String
''        Dim i As Integer = -1
''        Dim j As Integer = -1
''        Dim aSplit As Array

''        If My.Computer.FileSystem.FileExists(sLogFile) = True Then
''            FileOpen(2, sLogFile, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
''            Do While Not EOF(2)
''                sLine = Trim(LineInput(2))
''                If sLine <> "" Then
''                    aSplit = Split(sLine, "  ")
''                    If Trim(aSplit(4)) = "<DIR>" Then
''                        j = j + 1
''                        ReDim Preserve aFolders(j)
''                        aFolders(j) = aSplit(UBound(aSplit))
''                    Else
''                        i = i + 1
''                        ReDim Preserve aDates(i)
''                        ReDim Preserve aTimes(i)
''                        ReDim Preserve aFiles(i)
''                        aDates(i) = aSplit(0)
''                        aTimes(i) = aSplit(1)
''                        aFiles(i) = aSplit(UBound(aSplit))
''                        aFiles(i) = Trim(Right(aFiles(i), Len(aFiles(i)) - InStr(Trim(aFiles(i)), " ")))
''                    End If
''                End If
''NextLine:
''            Loop
''            FileClose(2)
''        End If
''        Return aFiles
''    End Function

''    Shared Function GetEntireList()
''        'Return entire file list (full file names) from FTP
''        Dim aDates() As String = Nothing
''        Dim aTimes() As String = Nothing
''        Dim aFiles() As String = Nothing
''        Dim aTempFiles() As String = Nothing
''        Dim aFolders() As String = Nothing
''        Dim sFolder As String = ""
''        Dim i As Integer

''        'Get inital file/folder structure
''        FTPOperation(tFTPAction.List, sFTPRootFolder & "\" & sDownload, , LocalFTPPath & "\Download.lst")
''        GetFileDetailList(LocalFTPPath & "\Download.lst", aDates, aTimes, aTempFiles, aFolders)
''        If My.Computer.FileSystem.FileExists(LocalFTPPath & "\Download.lst") = True Then
''            My.Computer.FileSystem.DeleteFile(LocalFTPPath & "\Download.lst")
''        End If
''        aFiles = PopulateFullFileList(aFiles, aTempFiles, aFolders, sFTPRootFolder & "\" & sDownload)
''ReRun:
''        For i = 0 To UBound(aFiles)
''            If aFiles(i) = "" Then GoTo NextItem
''            If SCEFile.IsFolder(aFiles(i)) = True Then
''                sFolder = aFiles(i)
''                aFiles(i) = ""
''                FTPOperation(tFTPAction.List, sFolder, , LocalFTPPath & "\Download.lst")
''                GetFileDetailList(LocalFTPPath & "\Download.lst", aDates, aTimes, aTempFiles, aFolders)
''                aFiles = PopulateFullFileList(aFiles, aTempFiles, aFolders, sFolder, True)
''                Erase aTempFiles
''                Erase aFolders
''                GoTo ReRun
''            End If
''NextItem:
''        Next

''        'trim results (remove spaces from array)
''        aFiles = SCEVB.TrimArray(aFiles)

''        Return aFiles
''    End Function

''    Shared Function PopulateFullFileList(ByVal aFiles() As String, ByVal aTempFiles() As String, ByVal aFolders() As String, ByVal sFolder As String, _
''        Optional ByVal bAppend As Boolean = False) As Array
''        'populate file list with full folder path
''        Dim i As Integer
''        Dim j As Integer

''        If SCEVB.IsNullArray(aTempFiles) = False Then
''            If SCEVB.IsNullArray(aFiles) = True Then
''                For i = 0 To UBound(aTempFiles)
''                    aFiles(i) = sFolder & "\" & aTempFiles(i)
''                Next
''            Else
''                j = UBound(aFiles)
''                For i = 0 To UBound(aTempFiles)
''                    j = j + 1
''                    ReDim Preserve aFiles(j)
''                    aFiles(j) = sFolder & "\" & aTempFiles(i)
''                Next
''            End If
''        End If

''        If SCEVB.IsNullArray(aFolders) = False Then
''            If SCEVB.IsNullArray(aFiles) = True Then
''                For i = 0 To UBound(aFolders)
''                    If i = 0 Then
''                        ReDim aFiles(0)
''                    Else
''                        ReDim Preserve aFiles(i)
''                    End If
''                    aFiles(i) = sFolder & "\" & aFolders(i)
''                Next
''            Else
''                j = UBound(aFiles)
''                For i = 0 To UBound(aFolders)
''                    j = j + 1
''                    ReDim Preserve aFiles(j)
''                    aFiles(j) = sFolder & "\" & aFolders(i)
''                Next
''            End If
''        End If

''        Return aFiles
''    End Function

''End Class
