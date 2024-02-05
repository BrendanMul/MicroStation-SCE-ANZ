Imports JMECore.SCEAccess
Imports JMECore

Public Class SCEText

    Public Shared dsClientText As DataSet = Nothing
    Public Shared Count As Integer
    Public Shared Loaded As Boolean

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Public Shared Sub Initialise(ByVal sDBFullName As String)
        dsClientText = GetAllDB(sDBFullName, "LUText", "Height")
        If SCEVB.IsNullDataSet(dsClientText) = True Then
            Count = 0
            Loaded = True
        Else
            Count = dsClientText.Tables(0).Rows.Count
            Loaded = True
        End If
    End Sub

    Public Shared Function GetIndex(ByVal sHeight As String, Optional ByVal sType As String = "", Optional ByVal sFont As String = "") As Long
        'Return index of specified text height
        Dim i As Long
        Dim iReturn As Integer = -1

        For i = 0 To (Count - 1)
            If Not dsClientText.Tables(0).Rows(i).Item("Height") Is System.DBNull.Value Then
                If UCase(sHeight) = UCase(dsClientText.Tables(0).Rows(i).Item("Height")) Then
                    If UCase(sType) = UCase(dsClientText.Tables(0).Rows(i).Item("Type")) Then
                        iReturn = i
                        Exit For
                    End If
                End If
            End If
NextItem:
        Next

        Return iReturn
    End Function

End Class
