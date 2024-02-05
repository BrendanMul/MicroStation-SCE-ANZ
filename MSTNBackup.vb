<ComClass(MSTNBackup.ClassId, MSTNBackup.InterfaceId, MSTNBackup.EventsId)> _
Public Class MSTNBackup

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "b1314ef4-b7f5-4aa0-99f2-e2b1132dd9f3"
    Public Const InterfaceId As String = "094ec840-698e-4e0c-8e6a-a35f1d234b72"
    Public Const EventsId As String = "28c65c0e-1376-49a8-b8e2-e6b8be2fc338"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

End Class


