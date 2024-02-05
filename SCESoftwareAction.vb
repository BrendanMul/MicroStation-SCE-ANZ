Public Class SCESoftwareAction
    'automated actions of software provided in Software table of SCECore database

    Shared Sub OpenPDF(ByVal sPDFFullName As String)
        'Launch Adobe Acrobat Reader to specified PDF file
        Dim sFullName As String
        Dim dsData As DataRow

        If My.Computer.FileSystem.FileExists(sPDFFullName) = True Then
            dsData = SCEAccess.GetRowDB(JEGCore.SCECoreDB, "LUSoftware", "Software", "Software", "Acrobat Reader", "", "", True, True)
            sFullName = dsData.Item("Location")
            If My.Computer.FileSystem.FileExists(sFullName) = True Then
                Shell(sFullName & Space(1) & Chr(34) & sPDFFullName & Chr(34), vbNormalFocus)
            Else
                MsgBox("Adobe Acrobat Reader was not found.", MsgBoxStyle.Critical, My.Settings.MsgBoxCaption)
            End If
            dsData = Nothing
            'Else
            '    MsgBox("PDF file does not exist." & Chr(10) & Chr(10) & " PDF file: " & sPDFFullName, MsgBoxStyle.Critical, My.Settings.MsgBoxCaption)
        End If
    End Sub

    Shared Sub Explorer(ByVal sPath As String)
        'Vaildates specified path and opens Windows Explorer to path
        If My.Computer.FileSystem.DirectoryExists(sPath) = True Then
            Shell("Explorer " & sPath, vbNormalFocus)
        End If
    End Sub

End Class
