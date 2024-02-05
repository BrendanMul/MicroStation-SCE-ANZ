Public Class SCECommunicator

    Shared Function GetStatusImage(ByVal iStatus As Integer) As String
        'Return image name associated with the presence status index
        Dim sReturn As String

        Select Case iStatus
            Case 1 'Offline
                sReturn = "presence_16-off.png"
            Case 2 'Online
                sReturn = "presence_16-online.png"
            Case 10 'Busy
                sReturn = "presence_16-busy.png"
            Case 14 'Be Right Back
                sReturn = "presence_16-away.png"
            Case 18 'Idle
                sReturn = "presence_16-idle-online.png"
            Case 34 'Away
                sReturn = "presence_16-away.png"
            Case 50 'On Phone
                sReturn = "presence_16-idle-busy.png"
            Case 82 'Meeting
                sReturn = "presence_16-idle-busy.png"
            Case Else 'Unknown
                sReturn = "presence_16-unknown.png"
        End Select

        sReturn = Replace(sReturn, "-", "_")
        sReturn = SCEFile.FileParse(sReturn, SCEFile.FileParser.FileName)
        Return sReturn
    End Function

    Shared Function GetStatusDescription(ByVal iStatus As Integer) As String
        'Return status description
        Dim sReturn As String = ""

        Select Case iStatus
            Case 1 'Offline
                sReturn = "Offline"
            Case 2 'Online
                sReturn = "Online"
            Case 10 'Busy
                sReturn = "Busy"
            Case 14 'Be Right Back
                sReturn = "Be right back"
            Case 18 'Idle
                sReturn = "Idle"
            Case 34 'Away
                sReturn = "Away"
            Case 50 'On Phone
                sReturn = "On the phone"
            Case 82 'Meeting
                sReturn = "In a meeting"
            Case 114 'do not disturb
                sReturn = "Do not disturb"
            Case Else
                sReturn = "Unknown"
        End Select

        Return sReturn
    End Function

    Shared Function GetRealName(ByVal sUser As String) As String
        'get real username from communicator friendly name
        Dim sReturn As String = ""
        Dim sName As String
        Dim aSplit As Array

        aSplit = Split(sUser, ",")
        If UBound(aSplit) = 0 Then
            sName = aSplit(0)
            sName = Replace(UCase(sName), "@RIOTINTO.COM", "")
            sReturn = Replace(UCase(sName), "@GLOBALSKM.COM", "")
        Else
            sName = Trim(Replace(aSplit(1), "(RTIO)", ""))
            sName = Trim(Replace(sName, "(RTIOEP)", ""))
            sName = Trim(Replace(aSplit(1), "(SKM)", ""))
            sReturn = sName & Space(1) & Trim(aSplit(0))
        End If

        Return sReturn
    End Function

    'Shared Sub SendIM(ByVal sToUserName As String)
    'initiate instant message
    ''        Dim oComm As CommunicatorAPI.Messenger
    ''        Dim oContact As CommunicatorAPI.IMessengerContact = Nothing
    ''        Dim bComm As Boolean = True
    ''        Dim bTest As Boolean

    ''        On Error GoTo Err
    ''        If bComm = True Then
    ''RetrySignIn:
    ''            oComm = New CommunicatorAPI.Messenger
    ''            If Not oComm Is Nothing Then
    ''                If oComm.MyStatus <> CommunicatorAPI.MISTATUS.MISTATUS_OFFLINE Then
    ''                    oContact = oComm.GetContact(sToUserName & "@" & SCECore.EmailDomain, oComm.MyServiceId)
    ''                    bComm = True
    ''                Else
    ''                    bComm = False
    ''                End If
    ''            Else
    ''                If bTest = False Then
    ''                    oComm.AutoSignin()
    ''                    bTest = True
    ''                    GoTo RetrySignIn
    ''                Else
    ''                    bComm = False
    ''                End If
    ''            End If
    ''            oComm.InstantMessage(oContact)
    ''            oContact = Nothing
    ''        End If
    ''        oComm = Nothing
    ''        oContact = Nothing
    ''        Exit Sub
    ''Err:
    ''        bComm = False
    ''        Err.Clear()
    ''        oComm = Nothing
    ''        oContact = Nothing
    'End Sub

    'Shared Function GetPhoneNumber(ByVal sUserName As String) As String
    'get specified users phone number
    ''        Dim oComm As Object
    ''        Dim oContact As CommunicatorAPI.IMessengerContact = Nothing
    ''        Dim oCommunicator As CommunicatorAPI.Messenger = Nothing
    ''        Dim bComm As Boolean
    ''        Dim bTest As Boolean
    ''        Dim sReturn As String = ""

    ''        bComm = True
    ''        On Error GoTo Err
    ''        If bComm = True Then
    ''RetrySignIn:
    ''            oComm = New CommunicatorAPI.Messenger
    ''            If Not oComm Is Nothing Then
    ''                If oComm.MyStatus <> "MISTATUS_OFFLINE" Then
    ''                    oContact = oComm.GetContact(sUserName & SCECore.EmailDomain, oComm.MyServiceId)
    ''                    bComm = True
    ''                Else
    ''                    bComm = False
    ''                End If
    ''            Else
    ''                If bTest = False Then
    ''                    oCommunicator.AutoSignin()
    ''                    bTest = True
    ''                    GoTo RetrySignIn
    ''                Else
    ''                    bComm = False
    ''                End If
    ''            End If
    ''            sReturn = oContact.PhoneNumber(CommunicatorAPI.MPHONE_TYPE.MPHONE_TYPE_WORK)
    ''            oContact = Nothing
    ''        End If

    ''        Return sReturn
    ''        oComm = Nothing
    ''        oContact = Nothing
    ''        Exit Function
    ''Err:
    ''        Return sReturn
    ''        bComm = False
    ''        Err.Clear()
    ''        oContact = Nothing
    'End Function

    'Shared Function GetUserContact(ByVal sUserName As String) 'As CommunicatorAPI.IMessengerContact
    'get specified users contact object
    ''        Dim oComm As Object
    ''        Dim oContact As CommunicatorAPI.IMessengerContact = Nothing
    ''        Dim oCommunicator As CommunicatorAPI.Messenger = Nothing
    ''        Dim bComm As Boolean
    ''        Dim bTest As Boolean
    ''        Dim oReturn As CommunicatorAPI.IMessengerContact = Nothing

    ''        bComm = True
    ''        On Error GoTo Err
    ''        If bComm = True Then
    ''RetrySignIn:
    ''            oComm = New CommunicatorAPI.Messenger
    ''            If Not oComm Is Nothing Then
    ''                If oComm.MyStatus <> CommunicatorAPI.MISTATUS.MISTATUS_OFFLINE Then
    ''                    oContact = oComm.GetContact(sUserName & "@" & SCECore.EmailDomain, oComm.MyServiceId)
    ''                    bComm = True
    ''                Else
    ''                    bComm = False
    ''                End If
    ''            Else
    ''                If bTest = False Then
    ''                    oCommunicator.AutoSignin()
    ''                    bTest = True
    ''                    GoTo RetrySignIn
    ''                Else
    ''                    bComm = False
    ''                End If
    ''            End If
    ''        End If
    ''        Return oContact
    ''        oContact = Nothing
    ''        Exit Function
    ''Err:
    ''        Return oReturn
    ''        bComm = False
    ''        Err.Clear()
    ''        oContact = Nothing
    'End Function

End Class
