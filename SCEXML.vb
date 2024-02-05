'Imports SCECoreDLL.File
Imports Microsoft.Win32
Imports System.Net

Public Class SCEXML

    Private Shared m_bFoundRemoteFile As Boolean = False
    Private Shared m_updateFile_url As String
    Private Shared m_file_name As String
    Private Shared m_localAppPath As String
    Private Shared m_UserTempPath As String
    Private Shared m_bNetworkIsOK As Boolean = True
    Private Shared m_lastModifiedDate As Date = Nothing
    Private Shared m_localModifiedDate As Date = Nothing

    Public Const HKCU As String = "HKCU"
    Public Const HKLM As String = "HKLM"

    Shared Function GetConfigurationValue(ByVal Key As String) As String 'From AutoCAD SCE
        Dim sXMLFileName As String = JEGCore.XMLFileName

        Try
            Dim xmlStructure As SKM.Common.XML.XmlDocument
            If My.Computer.FileSystem.FileExists(sXMLFileName) = True Then
                sXMLFileName = sXMLFileName
            Else
                If My.Computer.FileSystem.FileExists(JEGCore.LocalJMEPath & "\" & My.Settings.UpdateFile) = True Then
                    sXMLFileName = JEGCore.LocalJMEPath & My.Settings.UpdateFile
                    ''ElseIf My.Computer.FileSystem.FileExists(SCECore.cConfigurationFilePath & "\" & My.Settings.UpdateFile1) = True Then
                    ''    sXMLFileName = SCECore.cConfigurationFilePath & My.Settings.UpdateFile1
                End If
            End If
            xmlStructure = New SKM.Common.XML.XmlDocument(sXMLFileName)
            Dim nodelist As System.Xml.XmlNodeList = xmlStructure.GetElementsByTagName(Key)
            Dim result As String
            result = nodelist.Item(0).FirstChild.Value
            result = Replace(result, "[ProgramFiles]", My.Computer.FileSystem.SpecialDirectories.ProgramFiles, , , Microsoft.VisualBasic.CompareMethod.Text)
            Dim DocumentsAndSettingsFolder1 As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            DocumentsAndSettingsFolder1 = Replace(DocumentsAndSettingsFolder1, Environment.UserName & "\My Documents", "", , , Microsoft.VisualBasic.CompareMethod.Text)
            result = Replace(result, "[DocumentsAndSettings]", DocumentsAndSettingsFolder1)
            result = Replace(result, "[UserTemp]", JEGCore.UserTempPath)
RetryVariable:
            While result.Contains("[")
                Dim sKey As String = ""
                sKey = Mid(result, (InStr(result, "[") + 1), (InStr(result, "]") - (InStr(result, "[") + 1)))
                Dim nodelist1 As System.Xml.XmlNodeList = xmlStructure.GetElementsByTagName(sKey)
                If nodelist1.Count = 0 Then
                    result = Replace(result, "[ProgramFiles]", My.Computer.FileSystem.SpecialDirectories.ProgramFiles, , , Microsoft.VisualBasic.CompareMethod.Text)
                    result = Replace(result, "[DocumentsAndSettings]", DocumentsAndSettingsFolder1)
                    result = Replace(result, "[UserTemp]", JEGCore.UserTempPath)
                    result = Replace(result, "BUNDLECONTENTS", My.Settings.LocalFilePath32Bit)
                    result = Replace(result, "[", "")
                    result = Replace(result, "]", "")
                    GoTo SkipRetry
                Else
                    result = Replace(result, "[" & sKey & "]", nodelist1.Item(0).FirstChild.Value)
                End If
                GoTo RetryVariable
SkipRetry:
            End While
            Return result
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Shared Function GetXMLValue(ByVal sXMLFileName As String, ByVal Key As String) As String 'From AutoCAD SCE
        Try
            Dim xmlStructure As SKM.Common.XML.XmlDocument
            Dim Result As String

            xmlStructure = New SKM.Common.XML.XmlDocument(sXMLFileName)
            Dim nodelist As System.Xml.XmlNodeList = xmlStructure.GetElementsByTagName(Key)
            Result = nodelist.Item(0).FirstChild.Value
            Result = Replace(Result, "[ProgramFiles]", My.Computer.FileSystem.SpecialDirectories.ProgramFiles, , , Microsoft.VisualBasic.CompareMethod.Text)
            Dim DocumentsAndSettingsFolder As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            DocumentsAndSettingsFolder = Replace(DocumentsAndSettingsFolder, Environment.UserName & "\My Documents", "", , , Microsoft.VisualBasic.CompareMethod.Text)
            Result = Replace(Result, "[DocumentsAndSettings]", DocumentsAndSettingsFolder)
RetryVariable:
            While Result.Contains("[")
                Dim sKey As String = ""
                sKey = Mid(Result, (InStr(Result, "[") + 1), (InStr(Result, "]") - (InStr(Result, "[") + 1)))
                Dim nodelist1 As System.Xml.XmlNodeList = xmlStructure.GetElementsByTagName(sKey)
                If nodelist1.Count = 0 Then
                    Result = Replace(Result, "[ProgramFiles]", My.Computer.FileSystem.SpecialDirectories.ProgramFiles, , , Microsoft.VisualBasic.CompareMethod.Text)
                    Result = Replace(Result, "[DocumentsAndSettings]", DocumentsAndSettingsFolder)
                Else
                    Result = Replace(Result, "[" & sKey & "]", nodelist1.Item(0).FirstChild.Value)
                End If
                GoTo RetryVariable
            End While

            Return Result

        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Shared Function GetXMLConfigurationValueForProduct(ByVal product As String, ByVal item As String) As String
        Dim sXMLFilename As String = ""
        If My.Computer.FileSystem.FileExists(JEGCore.XMLFileName) = True Then
            sXMLFilename = JEGCore.XMLFileName
        Else
            If My.Computer.FileSystem.FileExists(My.Settings.UpdateFile) = True Then
                sXMLFilename = My.Application.Info.DirectoryPath & "\" & My.Settings.UpdateFile
                ''ElseIf My.Computer.FileSystem.FileExists(My.Settings.UpdateFile1) = True Then
                ''    sXMLFilename = My.Application.Info.DirectoryPath & "\" & My.Settings.UpdateFile1
            End If
        End If
        Dim xmlStructure As New SKM.Common.XML.XmlDocument(sXMLFilename)
        Dim nodelist As System.Xml.XmlNodeList = xmlStructure.GetElementsByTagName("Product")
        Dim sKeyValue As String = ""

        For Each node As System.Xml.XmlNode In nodelist
            If node.Item("Name").InnerText() = product Then
                sKeyValue = node.Item(item).InnerText()
            End If
        Next

        sKeyValue = Replace(sKeyValue, "[ProgramFiles]", My.Computer.FileSystem.SpecialDirectories.ProgramFiles, , , Microsoft.VisualBasic.CompareMethod.Text)
        Dim DocumentsAndSettingsFolder As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        DocumentsAndSettingsFolder = Replace(DocumentsAndSettingsFolder, Environment.UserName & "\My Documents", "", , , Microsoft.VisualBasic.CompareMethod.Text)
        sKeyValue = Replace(sKeyValue, "[DocumentsAndSettings]", DocumentsAndSettingsFolder)

RetryVariable:
        While sKeyValue.Contains("[")
            Dim sKey As String = ""
            sKey = Mid(sKeyValue, (InStr(sKeyValue, "[") + 1), (InStr(sKeyValue, "]") - (InStr(sKeyValue, "[") + 1)))
            Dim nodelist1 As System.Xml.XmlNodeList = xmlStructure.GetElementsByTagName(sKey)
            If nodelist1.Count = 0 Then
                sKeyValue = Replace(sKeyValue, "[ProgramFiles]", My.Computer.FileSystem.SpecialDirectories.ProgramFiles, , , Microsoft.VisualBasic.CompareMethod.Text)
                sKeyValue = Replace(sKeyValue, "[DocumentsAndSettings]", DocumentsAndSettingsFolder)
            Else
                sKeyValue = Replace(sKeyValue, "[" & sKey & "]", nodelist1.Item(0).FirstChild.Value)
            End If
            GoTo RetryVariable
        End While

        Return sKeyValue
    End Function

    Public Shared Function XMLFileDownloadShouldBeBypassed() As Boolean

        'Dim regVersion As RegistryKey

        '' Get user's local support folder
        'XMLFileDownloadShouldBeBypassed = False

        'Dim bByPassXMLDownload As String = "Software\SKMCAD"
        'regVersion = Registry.LocalMachine.OpenSubKey(bByPassXMLDownload, False)

        'If (Not regVersion Is Nothing) Then
        '    XMLFileDownloadShouldBeBypassed = regVersion.GetValue("ByPassXMLDownload")

        '    regVersion.Close()
        'End If

        Dim regVersion As RegistryKey

        ' Get user's local support folder
        XMLFileDownloadShouldBeBypassed = False

        Dim bByPassXMLDownload As String = "Software\Jacobs\SCE_R17"
        regVersion = Registry.LocalMachine.OpenSubKey(bByPassXMLDownload, False)

        If (Not regVersion Is Nothing) Then
            If UCase(regVersion.GetValue("ByPassXMLDownload").ToString) = "TRUE" Then
                XMLFileDownloadShouldBeBypassed = True
            Else
                XMLFileDownloadShouldBeBypassed = False
            End If
            regVersion.Close()
        End If

    End Function

    Public Shared Sub getLastModifiedDate()
        Dim request As WebRequest

        'GAD Removed the line below and added the following one
        'Dim response As HttpWebResponse
        Dim response As HttpWebResponse = Nothing

        Dim header As String
        Dim oWebclient As New WebClient

        m_updateFile_url = My.Settings.UpdateFileURL
        m_UserTempPath = JEGCore.LocalJMEPath
        m_UserTempPath = Replace(m_UserTempPath, "[ProgramFiles]", My.Computer.FileSystem.SpecialDirectories.ProgramFiles)

        Dim DocumentsAndSettingsFolder As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        DocumentsAndSettingsFolder = Replace(DocumentsAndSettingsFolder, Environment.UserName & "\My Documents", "", , , Microsoft.VisualBasic.CompareMethod.Text)
        m_UserTempPath = Replace(m_UserTempPath, "[DocumentsAndSettings]", DocumentsAndSettingsFolder)

        If My.Computer.FileSystem.FileExists(My.Settings.UpdateFile) = True Then
            m_file_name = My.Settings.UpdateFile
        Else
            ''m_file_name = My.Settings.UpdateFile1
        End If
        m_bNetworkIsOK = True
        m_bFoundRemoteFile = False
        Try
            'Gad added this condition to allow us to modify the local ConfigurationII.xml file for testing purposes
            ' If this reg key exists then we don't download the ConfigurationII.xml file from the server
            If XMLFileDownloadShouldBeBypassed() = False Then
                request = WebRequest.Create(m_updateFile_url & m_file_name)
                request.GetResponse()
                response = CType(request.GetResponse(), HttpWebResponse)
                header = response.GetResponseHeader("last-modified")
                m_lastModifiedDate = CDate(header)
                oWebclient.DownloadFile(m_updateFile_url + m_file_name, m_UserTempPath + m_file_name)
                m_bFoundRemoteFile = True
            Else
                'PJR
                'MsgBox("This machine has been configured not to download the " & My.Settings.UpdateFile & " file for testing purposes." & vbCrLf & _
                '"To remove this restriction please contact CAD Support", MsgBoxStyle.Information, "SCE Information")
            End If

        Catch ex As WebException
            m_bFoundRemoteFile = False
            m_lastModifiedDate = Nothing
            'Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbCrLf + "SKM Notice message will be displayed next time server is available." + vbCrLf + vbCrLf)
            m_bNetworkIsOK = False
        Catch headerEx As ObjectDisposedException
            m_bFoundRemoteFile = False
            m_lastModifiedDate = Nothing

            'Print error message to CAD command
            'Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("Cannot get " & _
            '            "header from URL request.Contact your system adminstrator. Error: " + headerEx.Message)
            m_bNetworkIsOK = False
        Catch createEx As UriFormatException
            m_bFoundRemoteFile = False
            m_lastModifiedDate = Nothing

            'Print error message to CAD command
            'Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("Requesting an " & _
            '            "incorrect URL from remote server.Contact your system adminstrator. Error: " + createEx.Message)
            m_bNetworkIsOK = False
        Catch createEx_1 As ArgumentNullException
            m_bFoundRemoteFile = False
            m_lastModifiedDate = Nothing

            'Print error message to CAD command
            'Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("Argument NULL exception. " & _
            '            "Contact your system adminstrator. Error: " + createEx_1.Message)
            m_bNetworkIsOK = False
        Finally
            If Not response Is Nothing Then
                response.Close()
                response = Nothing
            End If
            request = Nothing
        End Try

    End Sub

    Public Shared Function LoadProductsFromXML() As DataTable
        Dim sXMLFilename As String = JEGCore.XMLFileName
        If My.Computer.FileSystem.FileExists(sXMLFilename) = False Then
            sXMLFilename = JEGCore.LocalJMEPath & SCEFile.FileParse(JEGCore.XMLFileName, SCEFile.FileParser.FullFileName)
        End If
        Dim xmlStructure As New SKM.Common.XML.XmlDocument(sXMLFilename)
        Dim nodelist As System.Xml.XmlNodeList = xmlStructure.GetElementsByTagName("Product")

        'Complex DataBinding accepts as a data source either an IList or an IListSource
        Dim productTable As New DataTable
        productTable.Columns.Add(New DataColumn("Name", GetType(String)))
        productTable.Columns.Add(New DataColumn("Path", GetType(String)))

        Dim Products As System.Collections.IEnumerator = nodelist.GetEnumerator()
        Dim element As System.Xml.XmlElement = Nothing
        Dim selectedRow As DataRow = productTable.NewRow()
        Dim index As Integer = 0
        Dim selectedIndex As Integer = 0
        For Each node As System.Xml.XmlNode In nodelist
            If node.Item("Name").InnerText.Contains("MicroStation") Then
                Dim row As DataRow = productTable.NewRow()
                row("Name") = node.Item("Name").InnerText

                Dim sPath As String
                sPath = Replace(node.Item("Path").InnerText, "[ProgramFiles]", My.Computer.FileSystem.SpecialDirectories.ProgramFiles, , , Microsoft.VisualBasic.CompareMethod.Text)

                Dim DocumentsAndSettingsFolder As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
                DocumentsAndSettingsFolder = Replace(DocumentsAndSettingsFolder, Environment.UserName & "\My Documents", "", , , Microsoft.VisualBasic.CompareMethod.Binary)
                sPath = Replace(sPath, "[DocumentsAndSettings]", DocumentsAndSettingsFolder, , , Microsoft.VisualBasic.CompareMethod.Text)
RetryVariable:
                While sPath.Contains("[")
                    Dim sKey As String = ""
                    sKey = Mid(sPath, (InStr(sPath, "[") + 1), (InStr(sPath, "]") - (InStr(sPath, "[") + 1)))
                    Dim nodelist1 As System.Xml.XmlNodeList = xmlStructure.GetElementsByTagName(sKey)
                    If nodelist1.Count = 0 Then
                        sPath = Replace(sPath, "[ProgramFiles]", My.Computer.FileSystem.SpecialDirectories.ProgramFiles, , , Microsoft.VisualBasic.CompareMethod.Text)
                        sPath = Replace(sPath, "[DocumentsAndSettings]", DocumentsAndSettingsFolder)
                    Else
                        sPath = Replace(sPath, "[" & sKey & "]", nodelist1.Item(0).FirstChild.Value)
                    End If
                    GoTo RetryVariable
                End While
                row("Path") = sPath

                Dim sProductRegKey As String = SCEXML.GetXMLConfigurationValueForProduct(node.Item("Name").InnerText, "CheckKeyForProductExist")
                Dim sRegKey As String = SCEXML.GetXMLConfigurationValueForProduct(node.Item("Name").InnerText, "RegKey")

                If CheckIfRegistryHiveExists(sRegKey, HKLM, sProductRegKey) Then
                    productTable.Rows.Add(row)
                    index = index + 1
                End If
            End If
        Next

        Return productTable
    End Function

    Public Shared Function CheckIfRegistryHiveExists(ByVal regHive As String, ByVal location As String, ByVal key As String) As Boolean
        'Required for LoadProductsFromXML
        Dim regVersion As RegistryKey

        If location = HKCU Then
            regVersion = Registry.CurrentUser.OpenSubKey(regHive, False)
        Else
            regVersion = Registry.LocalMachine.OpenSubKey(regHive, False)
        End If

        If regVersion Is Nothing Then
            Return False
        Else
            Dim sKeyValue As String
            sKeyValue = regVersion.GetValue(key, "")

            If (Not sKeyValue Is Nothing) Then
                Return True
            Else
                Return False
            End If
        End If
    End Function

End Class
