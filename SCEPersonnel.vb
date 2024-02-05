Public Class SCEPersonnel

    Public Shared dsPersonnel As DataSet
    Public Shared Count As Integer
    Public Shared Loaded As Boolean
    Public Shared iCurrentUserIndex As Integer

    Public Enum tSCETitle
        Administrator
        Developer
    End Enum

    Public Shared Sub Initialise()
        'initialise SCE Personnel
        dsPersonnel = SCEAccess.GetAllDB(JEGCore.SCECoreDB, "RGPersonnel", "UserName", JEGCore.Password)
        If SCEVB.IsNullDataSet(dsPersonnel) = True Then
            iCurrentUserIndex = -1
        Else
            Count = dsPersonnel.Tables(0).Rows.Count
            Loaded = True

            iCurrentUserIndex = GetIndex(Environment.UserName)
        End If
    End Sub

    Public Shared Function GetIndex(ByVal sUserName As String) As Integer
        'get index of specified software
        Dim iReturn As Integer = -1
        Dim i As Integer

        For i = 0 To Count - 1
            If UCase(sUserName) = UCase(dsPersonnel.Tables(0).Rows(i).Item("UserName")) Then
                iReturn = i
                Exit For
            End If
        Next

        Return iReturn
    End Function

    Public Shared Function HasTitle(ByVal oTitle As tSCETitle) As Boolean
        'Return if the specified/current user has a particular title
        Dim bReturn As Boolean = False
        If iCurrentUserIndex = -1 Then
            GoTo ExitFunction
        End If
        If Not dsPersonnel.Tables(0).Rows(iCurrentUserIndex).Item("Titles") Is System.DBNull.Value Then
            If InStr(UCase(dsPersonnel.Tables(0).Rows(iCurrentUserIndex).Item("Titles")), UCase(oTitle.ToString)) > 0 Then
                bReturn = True
            End If
        End If
ExitFunction:
        Return bReturn
    End Function

End Class
