Public Class JEGLicense
    'class currently not in use

    Private Shared bAuthorise As Boolean
    Const iProductName As Integer = 0
    Const iAddress As Integer = 1
    Const iSuburb As Integer = 2
    Const iCountry As Integer = 3
    Const iExpiry As Integer = 4
    Const iPCName As Integer = 5
    Const iLicensePath As Integer = 6

    Shared Function Authorise()
        'PJR - add Authorise code
        bAuthorise = True

        Return bAuthorise
    End Function

    Shared Sub CreateLicense()
        'Create license file
        Dim i As Integer
        Dim sLicenseLine As String = ""
        Dim sDetails(6) As String
        Dim sExpDay As String = ""
        Dim sExpMonth As String = ""
        Dim sExpYear As String = ""

        MsgBox("The following dialogs require input, leaving field blank will exit the program.", MsgBoxStyle.Information, My.Settings.MsgBoxCaption)
        sDetails(0) = InputBox("Enter Product Name: (e.g. MicroStation)", My.Settings.MsgBoxCaption)
        If Trim(sDetails(0)) = "" Then GoTo ExitSub
        sDetails(1) = InputBox("Enter Office Address: (e.g. 263 Adelaide Terrace)", My.Settings.MsgBoxCaption)
        If Trim(sDetails(1)) = "" Then GoTo ExitSub
        sDetails(2) = InputBox("Enter Office Suburb: (e.g. Perth)", My.Settings.MsgBoxCaption)
        If Trim(sDetails(2)) = "" Then GoTo ExitSub
        sDetails(3) = InputBox("Enter Office Country: (e.g. Australia)", My.Settings.MsgBoxCaption)
        If Trim(sDetails(3)) = "" Then GoTo ExitSub
        sDetails(4) = InputBox("Enter Expiry Period: (number of months, maximum is 12)", My.Settings.MsgBoxCaption)
        If Trim(sDetails(4)) = "" Then GoTo ExitSub
        sDetails(5) = InputBox("Enter PC Name: (e.g. AU-D01234)", My.Settings.MsgBoxCaption)
        If Trim(sDetails(5)) = "" Then GoTo ExitSub
        sDetails(6) = InputBox("Enter path where license file is to be created: (e.g. D:\CADWORK)", My.Settings.MsgBoxCaption)
        If Trim(sDetails(6)) = "" Then GoTo ExitSub

        'Validate data
        For i = 0 To UBound(sDetails)
            sDetails(i) = Trim(sDetails(i))
            If i = iExpiry Then
                If sDetails(iExpiry) > 12 Then
                    GoTo ExitSub
                End If
            End If
        Next

        'Add months expiry to current date
        sExpDay = Format(Now, "dd")
        sExpMonth = CInt(Format(Now, "MM")) + sDetails(iExpiry)
        If sExpMonth > 12 Then
            sExpMonth = sExpMonth - 12
            sExpYear = Format(Now, "yyyy") + 1
        Else
            sExpYear = Format(Now, "yyyy")
        End If
        sDetails(iExpiry) = sExpDay & "\" & sExpMonth & "\" & sExpYear

        'Collate data into license string
        For i = 0 To UBound(sDetails)
            If sLicenseLine = "" Then
                sLicenseLine = sDetails(i)
            Else
                sLicenseLine = sLicenseLine & ";" & sDetails(i)
            End If
        Next

        sLicenseLine = EncryptLicense(sLicenseLine, My.Settings.LicensePassword)
        'Create license file
        My.Computer.FileSystem.CreateDirectory(sDetails(6))
        FileOpen(1, sDetails(6) & "\" & sDetails(0) & ".lic", OpenMode.Output, OpenAccess.Write, OpenShare.LockWrite)
        Print(1, sLicenseLine)
        FileClose(1)
        MsgBox("License has been created." & Chr(10) & Chr(10) & sDetails(6) & "\" & sDetails(0) & ".lic")
ExitSub:
    End Sub

    Shared Function EncryptLicense(ByVal sLine As String, ByVal sPassword As String)
        'Create license
        Dim i As Long
        Dim sChar As String
        Dim lRandom As Long
        Dim sLicense As String = ""
        Dim sRandom As String

        'sLine = "DWS02;1MILLST;PERTH;AUSTRALIA;08/07/2008;Z:\XM_WORKSPACE"
        'sLine = "Software;Street;Suburb;Country;Expiration date;Software Path
        If UCase(sPassword) <> UCase(My.Settings.LicensePassword) Then
            GoTo ExitFunction
        End If
        Randomize()
        lRandom = Int((6 * Rnd()) + 2)
        sRandom = lRandom
        If Len(sRandom) < 2 Then
            sRandom = "0" & lRandom
        End If
        For i = 1 To Len(sLine)
            sChar = Left(sLine, i)
            sLine = Right(sLine, (Len(sLine) - 1))
            sChar = Asc(sChar) * lRandom
            sLicense = sLicense & sChar & Len(sChar)
        Next
        sLicense = sLicense & sRandom & 2
ExitFunction:
        Return sLicense
    End Function

    Shared Function DecryptLicense(ByVal sLicense As String, ByVal cPassword As String) As Object
        'Split License into array/variant
        Dim sCountry As String
        Dim sAirportCode As String
        Dim sStreet As String
        Dim sLine As String = ""
        Dim sDate As String
        Dim sLocation As String
        Dim i As Long
        Dim sChar As String
        Dim lRandom As Long
        Dim sTempLicense As String
        Dim sLenChar As Long

        On Error GoTo Err
        'sLicense = "27233483332319232003236319633083292330433043332333632363320327633283336328832363260334033323336332832603304329232603236319232243188319232203188320031923192322432363360323233683352330833803348331633283300333233203260326832763042"
        sTempLicense = sLicense
        For i = 0 To Len(sLicense)
            If sTempLicense = "" Then Exit For
            sLenChar = Right(sTempLicense, 1)
            sTempLicense = Left(sTempLicense, Len(sTempLicense) - 1) 'minus character length
            sChar = Right(sTempLicense, sLenChar)
            If i = 0 Then
                lRandom = sChar
                sTempLicense = Left(sTempLicense, Len(sTempLicense) - 2) 'minus randomizer
            Else
                sChar = sChar / lRandom
                sChar = Chr(sChar)
                sLine = sChar & sLine
                sTempLicense = Left(sTempLicense, Len(sTempLicense) - sLenChar) 'minus character
            End If
        Next
        Return Split(sLine, ";")
ExitFunction:
        Exit Function
Err:
        MsgBox("Invalid License", vbCritical, My.Settings.LicenseMsgBoxCaption)
        GoTo ExitFunction
    End Function

    Shared Function IsValid(ByVal sLicenseFileName) As Boolean
        'return boolean if license is valid
        Dim bReturn As Boolean = False
        Dim iExpiry As Integer = 4
        Dim sLicenseLine As String = ""
        Dim aDetails As Array
        Dim aPCs As Array
        Dim sPCName As String
        Dim bPCName As Boolean = False

        If My.Computer.FileSystem.FileExists(sLicenseFileName) = True Then
            'Get license string
            FileOpen(1, sLicenseFileName, OpenMode.Input, OpenAccess.Default, OpenShare.Shared)
            sLicenseLine = Trim(LineInput(1))
            FileClose(1)

            aDetails = DecryptLicense(sLicenseLine, My.Settings.LicensePassword)
            If DateDiff(DateInterval.Day, CDate(Replace(aDetails(iExpiry), "\", "/")), Now) < 1 Then
                bReturn = True
            End If

            If bReturn = True Then
                If aDetails(iPCName) <> "" Then
                    aPCs = Split(aDetails(iPCName), "|")
                    If aPCs(0) = "All" Then
                        bPCName = True
                    Else
                        For Each sPCName In aPCs
                            If UCase(My.Computer.Name) = UCase(sPCName) Then
                                bPCName = True
                            End If
                        Next
                    End If
                    bReturn = bPCName
                End If
            End If
        End If

        Return bReturn
    End Function

    Shared Function IsExpiring(ByVal sLicenseFileName) As Boolean
        'return boolean if license is about to expire (one month)
        Dim bReturn As Boolean = False
        Dim iExpiry As Integer = 4
        Dim sLicenseLine As String = ""
        Dim aDetails As Array

        If My.Computer.FileSystem.FileExists(sLicenseFileName) = True Then
            'Get license string
            FileOpen(1, sLicenseFileName, OpenMode.Input, OpenAccess.Default, OpenShare.Shared)
            sLicenseLine = Trim(LineInput(1))
            FileClose(1)

            aDetails = DecryptLicense(sLicenseLine, My.Settings.LicensePassword)
            If DateDiff(DateInterval.Day, CDate(Replace(aDetails(iExpiry), "\", "/")), Now) > 30 Then
                bReturn = True
            End If
        End If

        Return bReturn
    End Function

    Shared Function GetExpiryDate(ByVal sLicenseFileName) As String
        'return expiry date
        Dim sReturn As String = ""
        'Dim iExpiry As Integer = 4
        Dim sLicenseLine As String = ""
        Dim aDetails As Array

        If My.Computer.FileSystem.FileExists(sLicenseFileName) = True Then
            'Get license string
            FileOpen(1, sLicenseFileName, OpenMode.Input, OpenAccess.Default, OpenShare.Shared)
            sLicenseLine = Trim(LineInput(1))
            FileClose(1)

            aDetails = DecryptLicense(sLicenseLine, My.Settings.LicensePassword)
            sReturn = Replace(aDetails(iExpiry), "\", "/")
        End If

        Return sReturn
    End Function
End Class
