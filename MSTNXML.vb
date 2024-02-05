<ComClass(MSTNXML.ClassId, MSTNXML.InterfaceId, MSTNXML.EventsId)> _
Public Class MSTNXML

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "bd7277aa-3814-46ed-aae7-a6b97363dd7b"
    Public Const InterfaceId As String = "013618e4-1465-4e48-bb92-dc05e2ddbb33"
    Public Const EventsId As String = "babb9527-2efb-4d22-adda-98bd0faf9396"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Shared Function GetXMLValue(ByVal sXMLFileName As String, ByVal sXMLFieldName As String, ByVal sXMLFieldValue As String) As String
        'Return entire XML file contents
        Dim sLine As String
        Dim sReturn As String = ""

        If My.Computer.FileSystem.FileExists(sXMLFileName) = False Then GoTo ExitFunction
        FileOpen(1, sXMLFileName, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
        Do While Not EOF(1)
            sLine = LineInput(1)
            sLine = Trim(sLine)
            If InStr(UCase(sLine), UCase(sXMLFieldName)) > 0 Then
                sLine = LineInput(1)
                sLine = Trim(sLine)
                sReturn = Replace(UCase(sXMLFieldValue), "</" & UCase(sXMLFieldValue) & ">", "")
                sReturn = Replace(UCase(sXMLFieldValue), "<" & UCase(sXMLFieldValue) & ">", "")
                Exit Do
            End If
        Loop
        FileClose(1)
ExitFunction:
        Return sReturn
    End Function

    Shared Function GetXMLSetting(ByVal sXMLFileName As String, ByVal sXMLFieldName As String) As String
        'Return XML setting
        Dim sLine As String
        Dim sReturn As String = ""

        If My.Computer.FileSystem.FileExists(sXMLFileName) = False Then GoTo ExitFunction
        FileOpen(1, sXMLFileName, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
        Do While Not EOF(1)
            sLine = LineInput(1)
            sLine = Trim(sLine)
            If InStr(UCase(sLine), "<" & UCase(sXMLFieldName) & ">") > 0 Then
                sLine = Replace(UCase(sLine), "</" & UCase(sXMLFieldName) & ">", "")
                sReturn = Replace(UCase(sReturn), "<" & UCase(sXMLFieldName) & ">", "")
                sReturn = Replace(UCase(sReturn), Chr(9), "")
                Exit Do
            End If
        Loop
        FileClose(1)
ExitFunction:
        Return sReturn
    End Function

    Shared Function GetXMLFields(ByVal sXMLFileName As String, ByVal sXMLFieldName As String)
        'Return entire XML file contents
        Dim aFields() As String
        Dim i As Long
        Dim sLine As String
        Dim aReturn() As String = Nothing

        If My.Computer.FileSystem.FileExists(sXMLFileName) = False Then
            ReDim aReturn(0)
            GoTo ExitFunction
        End If

        i = -1
        FileOpen(1, sXMLFileName, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
        Do While Not EOF(1)
            sLine = LineInput(1)
            sLine = Trim(sLine)
            If InStr(UCase(sLine), UCase(sXMLFieldName)) > 0 Then
                i = i + 1
                ReDim Preserve aFields(i)
                aReturn(i) = Replace(UCase(sLine), "</" & UCase(sXMLFieldName) & ">", "")
                aReturn(i) = Replace(UCase(aReturn(i)), "<" & UCase(sXMLFieldName) & ">", "")
            End If
        Loop
        FileClose(1)
        If i = -1 Then
            ReDim aReturn(0)
        End If
ExitFunction:
        Return aReturn
    End Function

    Shared Function GetXMLValues(ByVal sXMLFileName As String, ByVal sXMLFieldValueName As String)
        'Return field values
        Dim aReturn() As String = Nothing
        Dim i As Long
        Dim sLine As String

        If My.Computer.FileSystem.FileExists(sXMLFileName) = False Then
            ReDim aReturn(0)
            GoTo ExitFunction
        End If

        i = -1
        FileOpen(1, sXMLFileName, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
        Do While Not EOF(1)
            sLine = LineInput(1)
            sLine = Trim(sLine)
            If InStr(UCase(sLine), UCase(sXMLFieldValueName)) > 0 Then
                i = i + 1
                ReDim Preserve aReturn(i)
                aReturn(i) = Replace(UCase(sLine), "</" & UCase(sXMLFieldValueName) & ">", "")
                aReturn(i) = Replace(UCase(aReturn(i)), "<" & UCase(sXMLFieldValueName) & ">", "")
            End If
        Loop
        FileClose(1)
ExitFunction:
        Return aReturn
    End Function

End Class


