Public Class SCENetwork

    Shared Function GetADSiteNameII() As String
        'Return Active directory Site Name
        'Author: Corporate CAD
        Dim sReturn As String = ""
        Dim objADSysInfo As Object

        objADSysInfo = CreateObject("ADSystemInfo")
        Try
            sReturn = objADSysInfo.SiteName()
        Finally
            If Not objADSysInfo Is Nothing Then
                objADSysInfo = Nothing
            End If
        End Try

        Return sReturn
    End Function

    '    Shared Function PingSpeed(ByVal sFile As String)
    '        'Return ping return speed from specified file
    '        Dim sLine As String
    '        Dim sReturn As String = ""
    '        Dim aSplit As Array

    '        If My.Computer.FileSystem.FileExists(sFile) = False Then
    '            GoTo ExitSub
    '        End If
    '        FileOpen(1, sFile, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
    '        Do While Not EOF(1)
    '            sLine = LineInput(1)
    '            sLine = UCase(Trim(sLine))
    '            If InStr(sLine, "MINIMUM = ") > 0 Then
    '                aSplit = Split(sLine, ", AVERAGE = ")
    '                sReturn = Replace(aSplit(1), "MS", "")
    '                aSplit = Split(sReturn, Chr(44))
    '                sReturn = aSplit(0)
    '                Exit Do
    '            End If
    '        Loop
    '        FileClose(1)
    'ExitSub:
    '        Return sReturn
    '    End Function

End Class
