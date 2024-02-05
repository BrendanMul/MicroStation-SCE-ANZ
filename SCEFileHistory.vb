Public Class SCEFileHistory

    Public Shared dsFileHistory As DataSet
    Public Shared Count As Integer
    Public Shared Loaded As Boolean

    Public Shared Sub Initialise()
        'load file history into memory
        If SCEAccess.TableExists(JEGCore.UserDatabase, "RGFileHistory") = True Then
            dsFileHistory = SCEAccess.GetAllDB(JEGCore.UserDatabase, "RGFileHistory", "FullName", JEGCore.Password)
        End If

        If SCEVB.IsNullDataSet(dsFileHistory) = True Then
            Count = 0
            Loaded = True
        Else
            Count = dsFileHistory.Tables(0).Rows.Count
            Loaded = True
        End If
ExitSub:
    End Sub

    Public Shared Function GetIndex(ByVal sFullName As String) As Integer
        'load indexes
        Dim i As Integer
        Dim iReturn As Integer = -1

        For i = 0 To Count - 1
            If UCase(sFullName) = UCase(dsFileHistory.Tables(0).Rows(i).Item("FullName")) Then
                iReturn = i
                Exit For
            End If
        Next

        Return iReturn
    End Function

    Public Shared Function GetIndexes(ByVal sRegion As String, ByVal sClient As String, Optional ByVal sProjectName As String = "") As Array
        'load indexes
        Dim aReturn() As String = Nothing
        Dim i As Integer
        Dim j As Integer = -1

        For i = 0 To Count - 1
            If UCase(sRegion) = UCase(dsFileHistory.Tables(0).Rows(i).Item("Region")) And UCase(sClient) = UCase(dsFileHistory.Tables(0).Rows(i).Item("Client")) Then
                If sProjectName = "" Then
                    j = j + 1
                    ReDim Preserve aReturn(j)
                    aReturn(j) = i
                Else
                    If UCase(sRegion) = UCase(dsFileHistory.Tables(0).Rows(i).Item("Region")) And UCase(sClient) = UCase(dsFileHistory.Tables(0).Rows(i).Item("Client")) And UCase(sProjectName) = UCase(dsFileHistory.Tables(0).Rows(i).Item("ProjectName")) Then
                        j = j + 1
                        ReDim Preserve aReturn(j)
                        aReturn(j) = i
                    End If
                End If
            End If
        Next

        Return aReturn
    End Function

    Public Shared Sub AddFile(ByVal sFullName As String, ByVal sRegion As String, ByVal sClient As String, ByVal sSoftware As String, Optional ByVal sProjectName As String = "", Optional ByVal sAddon As String = "", Optional ByVal sTools As String = "")
        'Save file history
        Dim oRow As DataRow = Nothing
        Dim sFields As String
        Dim sValues As String

        sFields = SCEAccess.GetTableColumns("RGFileHistory")
        sValues = sFullName & "|" & sRegion & "|" & sClient & "|" & sProjectName & "|" & sSoftware & "|" & sAddon & "|" & sTools

        If SCEAccess.TableExists(JEGCore.UserDatabase, "RGFileHistory") = False Then
            SCEAccess.CreateTable(JEGCore.UserDatabase, "RGFileHistory", sFields)
        End If

        Try
            oRow = SCEAccess.GetRowDB(JEGCore.UserDatabase, "RGFileHistory", "FullName", "FullName", sFullName)
        Catch
        End Try

        If Not oRow Is Nothing Then
            SCEAccess.DelDB(JEGCore.UserDatabase, "RGFileHistory", "FullName", sFullName)
        End If

        SCEAccess.SetDB(JEGCore.UserDatabase, "RGFileHistory", sFields, sValues)

        Initialise()
    End Sub

    Public Shared Sub RemoveFile(ByVal sFullName As String)
        'remove file history file
        Dim oRow As DataRow = Nothing
        'Dim sFields As String
        'Dim sValues As String

        'sFields = SCEAccess.GetTableColumns("RGFileHistory")
        'sValues = sFullName & "|" & sRegion & "|" & sClient & "|" & sProjectName & "|" & sSoftware & "|" & sAddon & "|" & sTools

        Try
            oRow = SCEAccess.GetRowDB(JEGCore.UserDatabase, "RGFileHistory", "FullName", "FullName", sFullName)
        Catch
        End Try

        If Not oRow Is Nothing Then
            SCEAccess.DelDB(JEGCore.UserDatabase, "RGFileHistory", "FullName", sFullName)
        End If

        'SCEAccess.SetDB(SCECore.UserDatabase, "RGFileHistory", sFields, sValues)

        Initialise()
    End Sub

End Class
