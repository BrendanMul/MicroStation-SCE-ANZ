Imports JMECore.SCEXML
Imports JMECore.SCEVB

Public Class AutoCADSCE
    Public Shared aACADConfigs() As String
    Public Shared aACADRegions() As String
    Public Shared aACADClients() As String
    Public Shared aACADVersions() As String
    Public Shared Loaded As Boolean
    Public Shared Count As Integer

    Shared Sub LoadACADSCE()
        Dim sACADClientLibPath As String = "C:\ProgramData\Jacobs\Configurations\Client" 'GetConfigurationValue("LocalClientLibraryPath")
        Dim aFiles As Array = Nothing
        Dim j As Integer = -1
        Dim aSplit As Array
        Dim sVersion As String = ""
        Dim sFile As String

        If My.Computer.FileSystem.DirectoryExists(sACADClientLibPath) = True Then
            For Each sFile In My.Computer.FileSystem.GetFiles(sACADClientLibPath, FileIO.SearchOption.SearchTopLevelOnly, "*.zip")
                'aFiles = SCEFile.GetFiles(sACADClientLibPath, "zip")
                'If IsNullArray(aFiles) = False Then
                'For i = 0 To UBound(aFiles)

                ''pjrxxx check if sfile is fullpath
                sFile = SCEFile.FileParse(sFile, SCEFile.FileParser.FullFileName)
                aSplit = Split(Replace(sFile, ".zip", "", , , CompareMethod.Text), "-")
                If UBound(aSplit) = 4 Then
                    j = j + 1
                    ReDim Preserve aACADConfigs(j)
                    ReDim Preserve aACADRegions(j)
                    ReDim Preserve aACADClients(j)
                    ReDim Preserve aACADVersions(j)
                    aACADConfigs(j) = sFile 'aFiles(i)
                    aACADRegions(j) = RemoveVersion(aSplit(3), sVersion)
                    aSplit = Split(aSplit(4), "_")
                    aACADClients(j) = RemoveVersion(aSplit(0), sVersion)
                    aACADVersions(j) = aSplit(1)
                End If
                'Next
                'End If 'afiles array
            Next
        End If 'acad client lib path 

        Loaded = True
        Count = j + 1
    End Sub

    Shared Function RemoveVersion(ByVal sItem As String, ByRef sVersion As String) As String
        Dim sFinalStr As String = ""

        'GAD Removed the line below and added the following one
        'Dim nVersionIndex = sItem.Length
        Dim nVersionIndex As Integer = sItem.Length

        Dim cDummy As Char = New Char

        If (sItem <> "None") Then
            For nIndex As Integer = (sItem.Length - 1) To 0 Step -1
                If (Char.IsDigit(sItem.Chars(nIndex))) Then
                    nVersionIndex = nIndex
                Else
                    Exit For
                End If
            Next

            sVersion = sItem.Substring(nVersionIndex)
            sFinalStr = sItem.Substring(0, nVersionIndex)
        Else
            sVersion = ""
            sFinalStr = sItem
        End If

        Return sFinalStr
    End Function

    Shared Function GetACADConfig(ByVal sRegion As String, ByVal sClient As String, Optional ByVal sVersion As String = "") As String
        'Return config name, if no version specified return the latest version
        Dim i As Integer
        Dim iVersion As Integer = -1
        Dim sConfig As String = ""

        For i = 0 To UBound(aACADConfigs)
            If UCase(aACADRegions(i)) = UCase(sRegion) Then
                If UCase(aACADClients(i)) = UCase(sClient) Then
                    If sVersion = "" Then
                        'Get latest version
                        If aACADVersions(i) > iVersion Then
                            iVersion = aACADVersions(i)
                            sConfig = aACADConfigs(i)
                        End If
                    Else
                        If aACADVersions(i) = sVersion Then
                            sConfig = aACADConfigs(i)
                            Exit For
                        End If
                    End If
                End If
            End If
        Next
        Return sConfig
    End Function

    Shared Function GetConfigIndex(ByVal sClientConfig As String)
        'Return index of specified client config
        Dim i As Integer
        Dim iReturn As Integer = -1

        For i = 0 To UBound(aACADConfigs)
            If UCase(sClientConfig) = UCase(aACADConfigs(i)) Then
                iReturn = i
                Exit For
            End If
        Next
        Return iReturn
    End Function
End Class
