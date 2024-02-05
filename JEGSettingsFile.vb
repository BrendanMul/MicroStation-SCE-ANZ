Public Class JEGSettingsFile
    Public Shared bLoaded As Boolean = False
    Public Shared sSettingsFile As String

    Public Shared CADBuildVersion As String 'client provided version
    Public Shared Contact As String 'client build owner (jacobs personnel)
    Public Shared Downloadable As Boolean 'client build can be downloaded locally
    Public Shared HideProject As Boolean 'hide project from JME Launcher
    Public Shared IgnoreJMEConfig As Boolean 'do not load any JME configuration variables
    Public Shared Version As String 'incremented version identifier, not related to client provided version

    Public Shared Client As String 'client build
    Public Shared ClientContact As String 'client contact
    Public Shared ClientJacobsContact As String 'client jacobs contact
    Public Shared ClientDocumentation As String 'client documentation
    Public Shared ClientDownloadable As Boolean 'client build downloadable
    Public Shared ClientIgnoreJMEConfig As Boolean 'do not load any JME configuration variables
    Public Shared ClientLocation As String 'client build location
    Public Shared ClientJMEVersion As String 'JME version
    Public Shared ClientVersion As String 'client version
    Public Shared DMS As String

    Private Shared aSettings() As String
    Private Shared aValues() As String

    Public Shared Sub Initialise(ByVal sFilename As String, Optional ByVal sClientName As String = "")
        'load settings file contents
        Dim sFolder As String = ""
        Dim sLine As String
        Dim i As Integer = -1
        Dim aSplit As Array = Nothing

        If sClientName <> "" Then
            Client = ""
            ClientContact = ""
            ClientJacobsContact = ""
            ClientDocumentation = ""
            ClientDownloadable = False
            ClientIgnoreJMEConfig = False
            ClientLocation = ""
            ClientJMEVersion = ""
            ClientVersion = ""
        End If

        sSettingsFile = sFilename
        If My.Computer.FileSystem.FileExists(sSettingsFile) = True Then
            FileOpen(1, sSettingsFile, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
            Do While Not EOF(1)
                sLine = Trim(LineInput(1))
                If SCEVB.IsCommented(sLine) = False Then
                    i = i + 1
                    ReDim Preserve aSettings(i)
                    ReDim Preserve aValues(i)
                    aSplit = Split(sLine, "=")
                    aSplit(0) = Trim(aSplit(0))
                    aSplit(1) = Trim(aSplit(1))
                    aSettings(i) = aSplit(0)
                    aValues(i) = aSplit(1)
                End If

                'Initialise known settings
                Select Case UCase(aSettings(i))
                    Case "CADBUILDVERSION"
                        CADBuildVersion = aSplit(1)
                    Case "CLIENTCONTACT"
                        ClientContact = aValues(i)
                        Client = sClientName
                    Case "JACOBSCONTACT"
                        ClientJacobsContact = aValues(i)
                        Client = sClientName
                    Case "DOCUMENTATION"
                        ClientDocumentation = aValues(i)
                        Client = sClientName
                    Case "CLIENTDOWNLOADABLE"
                        ClientDownloadable = aValues(i)
                        Client = sClientName
                    Case "CLIENTIGNOREJMECONFIG"
                        ClientIgnoreJMEConfig = aValues(i)
                        Client = sClientName
                    Case "JACOBSVERSION"
                        ClientJMEVersion = aValues(i)
                        Client = sClientName
                    Case "CLIENTVERSION"
                        ClientVersion = aValues(i)
                        Client = sClientName
                    Case "CONTACT"
                        Contact = aSplit(1)
                    Case "DMS"
                        DMS = aValues(i)
                    Case "DOWNLOADABLE"
                        If UCase(aValues(i)) = "TRUE" Then
                            Downloadable = True
                        Else
                            Downloadable = False
                        End If
                    Case "HIDEPROJECT"
                        If UCase(aValues(i)) = "TRUE" Then
                            HideProject = True
                        Else
                            HideProject = False
                        End If
                    Case "IGNOREJMECONFIG"
                        If UCase(aValues(i)) = "TRUE" Then
                            IgnoreJMEConfig = True
                        Else
                            IgnoreJMEConfig = False
                        End If
                    Case "VERSION"
                        Version = aValues(i)
                End Select
            Loop
            FileClose(1)
            bLoaded = True
        Else
            bLoaded = False
            sSettingsFile = ""
        End If

    End Sub

    Public Shared Function GetSetting(ByVal sSetting As String) As String
        'return setting value
        Dim sReturn As String = ""
        Dim i As Integer

        If bLoaded = True Then
            For i = 0 To UBound(aSettings)
                If UCase(sSetting) = UCase(aSettings(i)) Then
                    sReturn = aValues(i)
                    Exit For
                End If
            Next
        End If

        Return sReturn
    End Function

    Public Shared Function IsClientInitialised(ByVal sClientName As String) As Boolean
        'return if specified client build settings has been initialised
        Dim bReturn As Boolean = False

        If UCase(Client) = UCase(sClientName) Then
            bReturn = True
        End If

        Return bReturn
    End Function
End Class

