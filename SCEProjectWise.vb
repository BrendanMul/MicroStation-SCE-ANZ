Public Class SCEProjectWise

    Public Enum IntegrationServer
        Australia
        UnitedKingdom
    End Enum

    Public Shared Function GetPWDB(ByVal sIntegrationServer As String, ByVal sDatasource As String, ByVal sStatement As String, ByVal sTable As String) As DataSet
        'Return the Column of a Table
        Dim cnSQL As OleDbConnection
        Dim dsReturn As New DataSet
        Dim da As OleDb.OleDbDataAdapter
        Dim j As Integer = -1
        Dim aReturn() As String = Nothing
        Dim sServer As String

        'Select Case UCase(sIntegrationServer)
        '    Case "AU-GLB-PWI04.SKMCONSULTING.COM"
        '        sServer = "AU-GLB-PWD03"
        '    Case Else
        '        sServer = "AU-GLB-PWD03"
        'End Select

        Select Case UCase(sIntegrationServer)
            Case "AUSYDO-DOC017CS.jacobs.com"
                sServer = "AUSYD0-SQL017"
            Case Else
                sServer = "AUSYD0-SQL017"
        End Select

        Try
            cnSQL = New OleDbConnection("PROVIDER=SQLOLEDB.1;INITIAL CATALOG=" & sDatasource & ";DATA SOURCE=" & sServer & ";User Id=PWReadOnly;Password=Password1") 'USER ID=PWReadOnly;Password=Password1;")
            cnSQL.Open()

            da = New OleDb.OleDbDataAdapter(sStatement, cnSQL)
            da.Fill(dsReturn, sTable)

            cnSQL.Close()
            da = Nothing
            cnSQL.Dispose()
            cnSQL = Nothing
            GoTo ExitFunction
        Catch e As OleDbException
            MsgBox(e.Message, MsgBoxStyle.Critical, "oledb Error")
        Catch e As Exception
            MsgBox(e.Message, MsgBoxStyle.Critical, "General Error")
        End Try
ExitFunction:
        Return dsReturn
    End Function

End Class
