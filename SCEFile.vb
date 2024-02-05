Imports System.IO
Imports System.Security.Principal
Imports System.Management
Imports IWshRuntimeLibrary

Public Class SCEFile

    Public Enum FileParser
        FullName = 0
        Drive = 1
        Path = 2
        FullFileName = 3
        FileName = 4
        Extension = 5
    End Enum

    Public Enum PathParser
        FullPath = 0
        Drive = 1
        Job = 2
        Deliverables = 3
        Drawings = 4
        Last = 5
        SecondLast = 6
    End Enum

    Public Enum GetFileType
        all = 0
        First = 1
    End Enum

    Shared Function FileParse(ByVal sFileName As String, ByVal sType As FileParser) As String
        'Simulates BASIC FileParse function
        'Author: Paul Ripp
        '0 Full name c:\sheets\test.dat
        '1 Drive c
        '2 Path c:\sheets
        '3 Name test.dat
        '4 Root test
        '5 Extension dat
        Dim aSplit As Array
        Dim sReturn As String = ""

        If Trim(sFileName) = "" Then
            GoTo ExitFunction
        End If

        On Error GoTo Err
        Select Case sType
            Case 0
                sReturn = sFileName
            Case 1
                sReturn = Left(sFileName, 1)
            Case 2
                sReturn = Left(sFileName, InStrRev(sFileName, "\") - 1)
            Case 3
                sReturn = Right(Replace(sFileName, "/", "\"), (Len(sFileName) - InStrRev(Replace(sFileName, "/", "\"), "\")))
            Case 4
                sReturn = Right(sFileName, (Len(sFileName) - InStrRev(sFileName, "\")))
                If InStr(sReturn, ".") > 0 Then
                    aSplit = Split(sReturn, ".")
                    sReturn = Left(sReturn, (Len(sReturn) - Len(aSplit(UBound(aSplit))) - 1))
                End If
            Case 5
                sReturn = Right(sFileName, (Len(sFileName) - InStrRev(sFileName, "\")))
                If InStr(sReturn, ".") > 0 Then
                    sReturn = Right(sReturn, Len(sReturn) - (InStr(sReturn, ".")))
                End If
        End Select
ExitFunction:

        Return sReturn
        Exit Function
Err:
        Return sFileName
        On Error GoTo 0
    End Function

    Shared Function PathParse(ByVal sFileName As String, ByVal sType As PathParser) As String
        'Returns path Sections in similar fashion as FileParse
        '0 Full Path P:\Job\Deliverables\Drawings\Area\Discipline
        '1 Drive P
        '2 Folder Job
        '3 Folder Deliverables
        'etc
        Dim aSplit As Array
        Dim sReturn As String = ""

        sFileName = Replace(sFileName, "/", "\")
        aSplit = Split(sFileName, "\")
        Select Case sType
            Case 0
                sReturn = sFileName
            Case 1
                sReturn = Left(aSplit(0), 1)
            Case PathParser.Last
                sReturn = aSplit(UBound(aSplit))
            Case PathParser.SecondLast
                sReturn = aSplit(UBound(aSplit) - 1)
            Case Else
                On Error GoTo ExitFunction
                sReturn = aSplit(sType - 1)
                On Error GoTo 0
        End Select
ExitFunction:
        Return sReturn
    End Function

    'Shared Function GetFile(ByVal sPath As String, ByVal sExtension As String) As String
    '    'Return first file name in specified path
    '    Dim sReturn As String
    '    Dim i As Long = -1

    '    If sExtension = "" Then
    '        sExtension = "*.*"
    '    Else
    '        sExtension = "*." & sExtension
    '    End If

    '    sReturn = Dir(sPath & "\" & sExtension, vbNormal)
    '    Do While sReturn <> ""
    '        Exit Do
    '    Loop

    '    Return sReturn
    'End Function

    'Shared Function GetFiles(ByVal sPath As String, ByVal sFilename As String, Optional ByVal sType As GetFileType = GetFileType.all, _
    '    Optional ByVal bSubDirectories As Boolean = False) As Array
    '    'Return list of filename(s) in specified path
    '    Dim aFiles() As String = Nothing
    '    Dim sFileName1 As String = Dir(sPath & "\*" & sFilename, vbNormal)
    '    Dim i As Integer = -1

    '    Do While sFileName1 <> ""
    '        If sFileName1 <> "." And sFileName1 <> ".." Then
    '            i = i + 1
    '            ReDim Preserve aFiles(i)
    '            aFiles(i) = sFileName1
    '        End If
    '        sFileName1 = Dir()
    '    Loop

    '    If i = -1 Then ReDim aFiles(0)

    '    Return aFiles
    'End Function

    'Shared Function GetFolder(ByVal sPath As String) As String
    '    'Return first folder in specified path
    '    Dim sReturn As String = Dir(sPath, vbDirectory)

    '    Do While sReturn <> ""
    '        Exit Do
    '    Loop

    '    Return sReturn
    'End Function

    '    Shared Function GetFolders(ByVal sPath As String) As Array
    '        'Return folders in specified path
    '        Dim sFolderName As String
    '        Dim aReturn() As String = Nothing
    '        Dim i As Long = -1

    '        If InStr(Right(sPath, 5), ".") > 0 Or My.Computer.FileSystem.DirectoryExists(sPath) = False Then
    '            GoTo ExitFunction
    '        End If
    '        If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
    '        sFolderName = Dir(sPath, vbDirectory)
    '        Do While sFolderName <> ""
    '            If sFolderName <> "." And sFolderName <> ".." Then
    '                If (GetAttr(sPath & sFolderName) And vbDirectory) = vbDirectory Then
    '                    i = i + 1
    '                    ReDim Preserve aReturn(i)
    '                    aReturn(i) = sFolderName
    '                End If
    '            End If
    '            sFolderName = Dir()
    '        Loop
    'ExitFunction:
    '        'no results found
    '        If i = -1 Then
    '            ReDim aReturn(0)
    '        End If

    '        Return aReturn
    '    End Function

    'Shared Sub CreateShortcut(ByVal sShortcutFile As String, ByVal sTargetFile As String, ByVal sDescription As String, Optional ByVal sIconEXEFile As String = "", Optional ByVal iIconIndex As Integer = 0)
    '    'Create shortcut file
    '    If My.Computer.FileSystem.FileExists(sShortcutFile) = False Then
    '        Dim shell As WshShell = New WshShellClass
    '        Dim shortcut As WshShortcut = shell.CreateShortcut(sShortcutFile)
    '        shortcut.TargetPath = sTargetFile
    '        If sIconEXEFile <> "" Then
    '            shortcut.IconLocation = sIconEXEFile & "," & iIconIndex
    '        End If
    '        shortcut.Description = sDescription
    '        shortcut.Save()
    '    End If
    'End Sub

    Shared Sub GetDirectoryFiles(ByVal sDirectory As String, ByRef cDirectoryEntries As Collection, Optional ByVal bIncludeSubDirectories As Boolean = False)
        Dim sFileEntries As String() = System.IO.Directory.GetFiles(sDirectory)

        ' Process the list of files found in the directory
        For Each sFileName As String In sFileEntries
            cDirectoryEntries.Add(sFileName)
        Next

        Dim sSubdirectoryEntries As String() = System.IO.Directory.GetDirectories(sDirectory)

        ' Recurse into subdirectories of this directory
        If bIncludeSubDirectories = True Then
            For Each sSubdirectory As String In sSubdirectoryEntries
                GetDirectoryFiles(sSubdirectory, cDirectoryEntries)
            Next
        End If
    End Sub

    Public Shared Function IsFolder(ByVal sPath As String) As Boolean
        'return boolean if specified path is folder
        Dim aSplit As Array
        Dim bReturn As Boolean = False

        aSplit = Split(sPath, "\")
        If InStr(aSplit(UBound(aSplit)), ".") > 0 Then
            bReturn = False
        Else
            bReturn = True
        End If

        Return bReturn
    End Function

    Public Shared Function IsOpen(ByVal sFullFileName As String) As Integer  'returns -1 if open, 0 if readable
        'This command only opens files which are not actively open.
        'It will open files that are referenced to another open file

        '    DGNstatus = IsOpen(filetoprocess(countint)) ' call file chk function
        Dim iReturn As Integer = 0 '0 = False

        Try
            FileOpen(1, sFullFileName, OpenMode.Binary, OpenAccess.Write, OpenShare.LockWrite)
            FileClose(1)
        Catch
            If Err.Number = 70 Then            'error message for write-locked files. Error 0 = readable files
                iReturn = 1             'True
            End If
        End Try

        Return iReturn
    End Function

    Public Shared Function GetOwner(ByVal sFileName As String) As String
        'return owner of specified file
        Dim sReturn As String = System.IO.File.GetAccessControl(sFileName).GetOwner(GetType(NTAccount)).Value

        If InStr(sReturn, "\") > 0 Then
            Dim aSplit As Array = Split(sReturn, "\")
            sReturn = aSplit(UBound(aSplit))
        End If

        Return sReturn
    End Function

    Public Shared Function ConvertSize(ByVal fileSize As Long) As String
        Dim sizeOfKB As Long = 1024                 ' Actual size in bytes of 1KB       
        Dim sizeOfMB As Long = 1048576              ' 1MB
        Dim sizeOfGB As Long = 1073741824           ' 1GB
        Dim sizeOfTB As Long = 1099511627776        ' 1TB
        Dim sizeofPB As Long = 1125899906842624     ' 1PB
        Dim tempFileSize As Double
        Dim tempFileSizeString As String
        Dim myArr() As Char = {CChar("0"), CChar(".")}

        'Characters to strip off the end of our string after formating   
        If fileSize < sizeOfKB Then 'Filesize is in Bytes        
            tempFileSize = ConvertBytes(fileSize, convTo.B)
            If tempFileSize = -1 Then Return Nothing
            'Invalid conversion attempted so exit
            tempFileSizeString = Format(fileSize, "Standard").TrimEnd(myArr) ' Strip the 0's and 1's off the end of the string
            Return Math.Round(tempFileSize) & " bytes" 'Return our converted value
        ElseIf fileSize >= sizeOfKB And fileSize < sizeOfMB Then 'Filesize is in Kilobytes
            tempFileSize = ConvertBytes(fileSize, convTo.KB)
            If tempFileSize = -1 Then Return Nothing 'Invalid conversion attempted so exit
            tempFileSizeString = Format(fileSize, "Standard").TrimEnd(myArr)
            Return Math.Round(tempFileSize) & " KB"
        ElseIf fileSize >= sizeOfMB And fileSize < sizeOfGB Then ' Filesize is in Megabytes
            tempFileSize = ConvertBytes(fileSize, convTo.MB)
            If tempFileSize = -1 Then Return Nothing 'Invalid conversion attempted so exit
            tempFileSizeString = Format(fileSize, "Standard").TrimEnd(myArr)
            Return Math.Round(tempFileSize, 1) & " MB"
        ElseIf fileSize >= sizeOfGB And fileSize < sizeOfTB Then 'Filesize is in Gigabytes
            tempFileSize = ConvertBytes(fileSize, convTo.GB)
            If tempFileSize = -1 Then Return Nothing
            tempFileSizeString = Format(fileSize, "Standard").TrimEnd(myArr)
            Return Math.Round(tempFileSize, 1) & " GB"
        ElseIf fileSize >= sizeOfTB And fileSize < sizeofPB Then 'Filesize is in Terabytes
            tempFileSize = ConvertBytes(fileSize, convTo.TB)
            If tempFileSize = -1 Then Return Nothing
            tempFileSizeString = Format(fileSize, "Standard").TrimEnd(myArr)
            Return Math.Round(tempFileSize, 1) & " TB"
            'Anything bigger than that is silly ;)
        Else
            Return Nothing 'Invalid filesize so return Nothing
        End If
    End Function

    Public Shared Function ConvertBytes(ByVal bytes As Long, ByVal convertTo As convTo) As Double
        If convTo.IsDefined(GetType(convTo), convertTo) Then
            Return bytes / (1024 ^ convertTo)
        Else
            Return -1 'An invalid value was passed to this function so exit
        End If
    End Function

    Public Enum convTo
        B = 0
        KB = 1
        MB = 2
        GB = 3  'Enumerations for file size conversions
        TB = 4
        PB = 5
        EB = 6
        ZI = 7
        YI = 8
    End Enum

    Public Function GetUNC(ByVal sFilePath As String) As String
        GetUNC = ""
        Dim searcher As New ManagementObjectSearcher("SELECT RemoteName FROM win32_NetworkConnection WHERE LocalName = '" + sFilePath.Substring(0, 2) + "'")
        For Each managementObject As ManagementObject In searcher.[Get]()
            Dim sRemoteName As String = TryCast(managementObject("RemoteName"), String)
            sRemoteName += sFilePath.Substring(2)
            Return sRemoteName.ToString
        Next

    End Function

    '    Public Shared Sub VBFileCopy(ByVal sSourcePath As String, ByVal sDestinationPath As String, Optional ByVal bOverwrite As Boolean = False, Optional ByVal prgProgressBar As ProgressBar = Nothing)
    '        'VB copying of files with optional progress bar
    '        Dim sSourceFile As String = ""
    '        Dim sDestinationFile As String = ""
    '        Dim iCount As Integer = 0
    '        Dim i As Integer = 0
    '        Dim bVisible As Boolean = False

    '        If Not prgProgressBar Is Nothing Then
    '            bVisible = prgProgressBar.Visible
    '        End If

    '        If My.Computer.FileSystem.DirectoryExists(sSourcePath) = False Then
    '            GoTo ExitSub
    '        End If

    '        If My.Computer.FileSystem.DirectoryExists(sDestinationPath) = False Then
    '            My.Computer.FileSystem.CreateDirectory(sDestinationPath)
    '        End If

    '        'Get file count
    '        For Each sSourceFile In My.Computer.FileSystem.GetFiles(sSourcePath, FileIO.SearchOption.SearchAllSubDirectories, "*.*")
    '            iCount = iCount + 1
    '        Next

    '        'Initialise progress bar
    '        prgProgressBar.Maximum = iCount
    '        prgProgressBar.Value = 0
    '        If prgProgressBar.Visible = False Then
    '            prgProgressBar.Visible = True
    '            prgProgressBar.Refresh()
    '        End If

    '        'Copy files
    '        For Each sSourceFile In My.Computer.FileSystem.GetFiles(sSourcePath, FileIO.SearchOption.SearchAllSubDirectories, "*.*")
    '            i = i + 1
    '            prgProgressBar.Value = i
    '            prgProgressBar.Refresh()
    '            sDestinationFile = Replace(sSourceFile, sSourcePath, sDestinationPath, , , Microsoft.VisualBasic.CompareMethod.Binary)
    '            'pjrxxx - add destination, check folder exists and create?
    '            My.Computer.FileSystem.CopyFile(sSourceFile, sDestinationFile, bOverwrite)
    '        Next
    'ExitSub:
    '        If Not prgProgressBar Is Nothing Then
    '            prgProgressBar.Visible = bVisible
    '            prgProgressBar.Refresh()
    '        End If
    '    End Sub
End Class
