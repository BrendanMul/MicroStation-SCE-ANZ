Public Class SCEProcesses

    Shared Function RunningCount(ByVal sProcessName As String, ByVal sFileName As String) As Integer
        'Return if program is already running
        Dim oProcess As Process
        Dim iCount As Integer

        For Each oProcess In Process.GetProcessesByName(sProcessName)
            If UCase(oProcess.MainModule.ModuleName) = UCase(sFileName) Then
                iCount = iCount + 1
            End If
        Next

        Return iCount
    End Function

    Shared Function AlreadyRunning(ByVal sProgramName As String, ByVal sFileName As String) As Boolean
        'Return if program is already running
        If RunningCount(sProgramName, sFileName) < 1 Then
            Return False
        Else
            Return True
        End If
    End Function

    Shared Function QuitSessions(ByVal sProcessName As String, ByVal sFileName As String) As Integer
        'Quit program if running
        Dim oProcess As Process
        Dim iCount As Integer

        For Each oProcess In Process.GetProcessesByName(sProcessName)
            If UCase(oProcess.MainModule.ModuleName) = UCase(sFileName) Then
                oProcess.Kill()
                oProcess.Dispose()
            End If
        Next

        Return iCount
    End Function


    Public Shared Function CoreCount() As Integer
        Return Environment.ProcessorCount
        ''Dim scope As New System.Management.ManagementScope("\\.\root\cimv2")
        ''scope.Connect()
        ''Dim objectQuery As New System.Management.ObjectQuery("SELECT * FROM Win32_Processor")
        ''Dim cpu As System.Management.ManagementObject
        ''Dim searcher As New System.Management.ManagementObjectSearcher(scope, objectQuery)
        ''Dim iReturn As Integer

        ''For Each cpu In searcher.Get()
        ''    iReturn = iReturn + cpu("NumberOfCores")
        ''Next

        ''Return iReturn
    End Function
End Class
