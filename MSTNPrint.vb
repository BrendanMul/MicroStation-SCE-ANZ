<ComClass(MSTNPrint.ClassId, MSTNPrint.InterfaceId, MSTNPrint.EventsId)> _
Public Class MSTNPrint

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "38a597a0-a8ef-450a-bcd8-f5a9a7a67881"
    Public Const InterfaceId As String = "b0c76205-fca0-4261-9219-d22e018d56cf"
    Public Const EventsId As String = "9b724f28-9894-4cda-9279-55da1f1896ef"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

End Class


